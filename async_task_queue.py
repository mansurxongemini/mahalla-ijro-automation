"""
Production-Ready Async Event-Driven Task Queue System
======================================================

Architecture Overview:
----------------------
This system implements a high-performance, thread-safe async task queue using Python's asyncio.

Key Design Decisions for Thread Safety & Race Condition Prevention:
1. Single-Threaded Event Loop: All state mutations occur within the asyncio event loop,
   eliminating traditional threading race conditions. We use asyncio.Lock for critical sections
   where multiple coroutines might access shared state.

2. Atomic Operations: All queue operations (enqueue, dequeue, priority updates) are wrapped
   in async context managers ensuring atomicity.

3. Priority Heap with Aging: Uses a min-heap for O(log n) priority operations. The aging
   mechanism runs as a background task that periodically recalculates effective priorities.

4. Token Bucket Rate Limiting: Each task type has its own token bucket with atomic refill/consume
   operations protected by locks.

5. Circuit Breaker Pattern: Per-task-type circuit breakers track failure counts and open/close
   based on configurable thresholds.

6. Backpressure Handling: When queue exceeds MAX_QUEUE_SIZE, applies rejection policy with
   priority-aware dropping of non-aged tasks.

Components:
- TaskQueue: Main queue with priority heap and aging mechanism
- TokenBucket: Rate limiter per task type
- CircuitBreaker: Failure tracking and fallback execution
- DeadLetterQueue: Failed task storage with state serialization
- TaskProcessor: Worker pool with concurrency control
"""

import asyncio
import time
import json
import heapq
import logging
import random
from dataclasses import dataclass, field, asdict
from enum import Enum, IntEnum
from typing import (
    Any, Callable, Coroutine, Dict, List, Optional, Set, Tuple, TypeVar, Union
)
from collections import defaultdict
from datetime import datetime
from contextlib import asynccontextmanager
import uuid
import traceback

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class Priority(IntEnum):
    """Task priority levels (lower value = higher priority)"""
    HIGH = 0
    MEDIUM = 1
    LOW = 2
    
    def age(self) -> 'Priority':
        """Return the next higher priority (aged)"""
        if self == Priority.LOW:
            return Priority.MEDIUM
        elif self == Priority.MEDIUM:
            return Priority.HIGH
        return self  # HIGH stays HIGH


class TaskStatus(Enum):
    """Task lifecycle states"""
    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    RETRYING = "retrying"
    DEAD_LETTER = "dead_letter"
    REJECTED = "rejected"


class CircuitState(Enum):
    """Circuit breaker states"""
    CLOSED = "closed"
    OPEN = "open"
    HALF_OPEN = "half_open"


@dataclass(order=True)
class Task:
    """
    Task representation with priority ordering.
    
    The order is determined by:
    1. Effective priority (considering aging)
    2. Timestamp (older tasks first for same priority)
    """
    effective_priority: int = field(compare=True)
    task_id: str = field(compare=False, default_factory=lambda: str(uuid.uuid4()))
    task_type: str = field(compare=False, default="default")
    payload: Dict[str, Any] = field(compare=False, default_factory=dict)
    original_priority: Priority = field(compare=False, default=Priority.MEDIUM)
    created_at: float = field(compare=False, default_factory=time.time)
    status: TaskStatus = field(compare=False, default=TaskStatus.PENDING)
    retry_count: int = field(compare=False, default=0)
    max_retries: int = field(compare=False, default=3)
    error_log: List[str] = field(compare=False, default_factory=list)
    metadata: Dict[str, Any] = field(compare=False, default_factory=dict)
    
    def __post_init__(self):
        if not isinstance(self.effective_priority, int):
            self.effective_priority = int(self.original_priority)
    
    def get_effective_priority(self) -> int:
        """Calculate effective priority considering aging"""
        age_seconds = time.time() - self.created_at
        aged_priority = self.original_priority
        
        # Age by 1 level every 10 seconds
        aging_steps = int(age_seconds / 10)
        for _ in range(min(aging_steps, 2)):  # Max 2 aging steps (LOW->MEDIUM->HIGH)
            aged_priority = aged_priority.age()
        
        return int(aged_priority)
    
    def to_dict(self) -> Dict[str, Any]:
        """Serialize task to dictionary"""
        return {
            'task_id': self.task_id,
            'task_type': self.task_type,
            'payload': self.payload,
            'original_priority': self.original_priority.name,
            'effective_priority': self.get_effective_priority(),
            'created_at': self.created_at,
            'status': self.status.value,
            'retry_count': self.retry_count,
            'max_retries': self.max_retries,
            'error_log': self.error_log,
            'metadata': self.metadata
        }
    
    def serialize_state(self) -> str:
        """Complete structural state serialization for DLQ"""
        return json.dumps(self.to_dict(), indent=2, default=str)


@dataclass
class TokenBucket:
    """
    Token bucket rate limiter for a specific task type.
    
    Implements atomic token consumption and refill operations.
    """
    task_type: str
    capacity: int = 100
    tokens: float = field(default=100.0)
    refill_rate: float = field(default=10.0)  # tokens per second
    last_refill: float = field(default_factory=time.time)
    lock: asyncio.Lock = field(default_factory=asyncio.Lock, repr=False)
    
    async def consume(self, tokens: int = 1) -> bool:
        """
        Attempt to consume tokens from the bucket.
        Returns True if successful, False otherwise.
        """
        async with self.lock:
            self._refill()
            if self.tokens >= tokens:
                self.tokens -= tokens
                return True
            return False
    
    async def wait_for_token(self, tokens: int = 1, timeout: float = 30.0) -> bool:
        """Wait until tokens are available or timeout"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            if await self.consume(tokens):
                return True
            await asyncio.sleep(0.1)
        return False
    
    def _refill(self):
        """Refill tokens based on elapsed time"""
        now = time.time()
        elapsed = now - self.last_refill
        self.tokens = min(self.capacity, self.tokens + elapsed * self.refill_rate)
        self.last_refill = now
    
    def update_rate_based_on_load(self, cpu_usage: float, memory_usage: float):
        """Dynamically adjust refill rate based on system load"""
        # Reduce rate when load is high
        load_factor = max(cpu_usage, memory_usage) / 100.0
        if load_factor > 0.8:
            self.refill_rate = max(1.0, self.refill_rate * 0.5)
        elif load_factor < 0.3:
            self.refill_rate = min(100.0, self.refill_rate * 1.2)


@dataclass
class CircuitBreaker:
    """
    Circuit breaker for a specific task type.
    
    States:
    - CLOSED: Normal operation, failures are tracked
    - OPEN: Circuit tripped, all requests fail immediately
    - HALF_OPEN: Testing if service recovered
    """
    task_type: str
    failure_threshold: int = 3
    recovery_timeout: float = 30.0
    failure_count: int = field(default=0)
    state: CircuitState = field(default=CircuitState.CLOSED)
    last_failure_time: Optional[float] = field(default=None)
    lock: asyncio.Lock = field(default_factory=asyncio.Lock, repr=False)
    fallback_handler: Optional[Callable] = field(default=None, repr=False)
    
    async def record_success(self):
        """Record successful execution, reset failure count"""
        async with self.lock:
            self.failure_count = 0
            self.state = CircuitState.CLOSED
    
    async def record_failure(self) -> bool:
        """
        Record failure and potentially open circuit.
        Returns True if circuit is now open.
        """
        async with self.lock:
            self.failure_count += 1
            self.last_failure_time = time.time()
            
            if self.failure_count >= self.failure_threshold:
                self.state = CircuitState.OPEN
                logger.warning(f"Circuit breaker OPEN for {self.task_type}")
                return True
            return False
    
    async def can_execute(self) -> bool:
        """Check if execution is allowed"""
        async with self.lock:
            if self.state == CircuitState.CLOSED:
                return True
            
            if self.state == CircuitState.OPEN:
                # Check if recovery timeout has passed
                if self.last_failure_time and \
                   time.time() - self.last_failure_time > self.recovery_timeout:
                    self.state = CircuitState.HALF_OPEN
                    logger.info(f"Circuit breaker HALF_OPEN for {self.task_type}")
                    return True
                return False
            
            # HALF_OPEN: allow one test request
            return True
    
    async def execute_with_fallback(
        self,
        task: Task,
        handler: Callable[[Task], Coroutine[Any, Any, Any]]
    ) -> Any:
        """Execute handler with fallback if circuit is open"""
        if not await self.can_execute():
            logger.warning(f"Circuit open for {task.task_type}, executing fallback")
            if self.fallback_handler:
                return await self.fallback_handler(task)
            raise Exception(f"Circuit breaker open for {task.task_type}")
        
        try:
            result = await handler(task)
            await self.record_success()
            return result
        except Exception as e:
            circuit_opened = await self.record_failure()
            if circuit_opened and self.fallback_handler:
                logger.info(f"Executing fallback due to circuit break for {task.task_type}")
                return await self.fallback_handler(task)
            raise


@dataclass
class DeadLetterQueue:
    """
    Dead Letter Queue for failed tasks.
    
    Stores complete task state serialization for later analysis/replay.
    """
    max_size: int = 10000
    tasks: List[Task] = field(default_factory=list)
    lock: asyncio.Lock = field(default_factory=asyncio.Lock, repr=False)
    storage_path: Optional[str] = None
    
    async def add(self, task: Task, reason: str):
        """Add failed task to DLQ with reason"""
        async with self.lock:
            task.status = TaskStatus.DEAD_LETTER
            task.metadata['dlq_reason'] = reason
            task.metadata['dlq_timestamp'] = time.time()
            
            # Persist to file if path configured
            if self.storage_path:
                await self._persist_task(task)
            
            self.tasks.append(task)
            
            # Enforce max size (drop oldest)
            if len(self.tasks) > self.max_size:
                dropped = self.tasks.pop(0)
                logger.warning(f"Dropped oldest DLQ task: {dropped.task_id}")
            
            logger.error(f"Task {task.task_id} moved to DLQ: {reason}")
    
    async def _persist_task(self, task: Task):
        """Persist task state to file"""
        try:
            import aiofiles
            async with aiofiles.open(self.storage_path, 'a') as f:
                await f.write(task.serialize_state() + '\n')
        except ImportError:
            # Fallback to sync file write
            with open(self.storage_path, 'a') as f:
                f.write(task.serialize_state() + '\n')
        except Exception as e:
            logger.error(f"Failed to persist DLQ task: {e}")
    
    async def get_all(self) -> List[Dict[str, Any]]:
        """Get all DLQ tasks as dictionaries"""
        async with self.lock:
            return [task.to_dict() for task in self.tasks]
    
    async def replay_task(self, task_id: str) -> Optional[Task]:
        """Remove and return a task for replay"""
        async with self.lock:
            for i, task in enumerate(self.tasks):
                if task.task_id == task_id:
                    return self.tasks.pop(i)
        return None
    
    def size(self) -> int:
        return len(self.tasks)


class SystemLoadMonitor:
    """Simulated system load monitor for rate limiting decisions"""
    
    def __init__(self):
        self.cpu_usage: float = 0.0
        self.memory_usage: float = 0.0
        self.lock = asyncio.Lock()
    
    async def update_load(self, cpu: float, memory: float):
        """Update current system load metrics"""
        async with self.lock:
            self.cpu_usage = min(100.0, max(0.0, cpu))
            self.memory_usage = min(100.0, max(0.0, memory))
    
    async def get_load(self) -> Tuple[float, float]:
        """Get current CPU and memory usage"""
        async with self.lock:
            return self.cpu_usage, self.memory_usage
    
    async def simulate_load_fluctuation(self):
        """Background task to simulate realistic load changes"""
        while True:
            new_cpu = random.gauss(50, 20)
            new_memory = random.gauss(60, 15)
            await self.update_load(new_cpu, new_memory)
            await asyncio.sleep(5)


class TaskQueue:
    """
    Main async task queue with priority handling, aging, and backpressure.
    
    Thread Safety Guarantees:
    - All mutations protected by asyncio.Lock
    - Heap operations are atomic within lock context
    - Aging mechanism runs as separate task but respects locks
    """
    
    MAX_QUEUE_SIZE = 10000
    AGING_INTERVAL = 1.0  # Check for aging every second
    
    def __init__(self):
        self._heap: List[Tuple[int, float, Task]] = []  # (priority, timestamp, task)
        self._task_index: Dict[str, Task] = {}  # Fast lookup by task_id
        self._lock = asyncio.Lock()
        self._not_empty = asyncio.Condition(self._lock)
        self._not_full = asyncio.Condition(self._lock)
        self._aging_task: Optional[asyncio.Task] = None
        self._running = False
        
        # Statistics
        self.enqueued_count = 0
        self.dequeued_count = 0
        self.rejected_count = 0
        self.aged_count = 0
    
    async def start(self):
        """Start background tasks"""
        self._running = True
        self._aging_task = asyncio.create_task(self._aging_loop())
        logger.info("TaskQueue started")
    
    async def stop(self):
        """Stop background tasks"""
        self._running = False
        if self._aging_task:
            self._aging_task.cancel()
            try:
                await self._aging_task
            except asyncio.CancelledError:
                pass
        logger.info("TaskQueue stopped")
    
    async def _aging_loop(self):
        """Background task to handle priority aging"""
        while self._running:
            try:
                await self._process_aging()
                await asyncio.sleep(self.AGING_INTERVAL)
            except asyncio.CancelledError:
                break
            except Exception as e:
                logger.error(f"Aging loop error: {e}")
    
    async def _process_aging(self):
        """Recalculate effective priorities for all pending tasks"""
        async with self._lock:
            updated = False
            new_heap = []
            
            for priority, timestamp, task in self._heap:
                new_priority = task.get_effective_priority()
                if new_priority != priority:
                    task.effective_priority = new_priority
                    updated = True
                    self.aged_count += 1
                
                heapq.heappush(new_heap, (new_priority, timestamp, task))
            
            if updated:
                self._heap = new_heap
                logger.debug(f"Aging processed, {self.aged_count} tasks aged so far")
    
    async def enqueue(
        self,
        task: Task,
        timeout: Optional[float] = None
    ) -> bool:
        """
        Add task to queue with backpressure handling.
        
        Returns True if enqueued, False if rejected.
        Blocks if queue is full (unless timeout specified).
        """
        async with self._not_full:
            # Check if we need to apply backpressure
            if len(self._heap) >= self.MAX_QUEUE_SIZE:
                if timeout is not None and timeout > 0:
                    try:
                        await asyncio.wait_for(
                            self._not_full.wait(),
                            timeout=timeout
                        )
                    except asyncio.TimeoutError:
                        return await self._apply_rejection_policy(task)
                else:
                    # Apply rejection policy immediately if no timeout or zero timeout
                    return await self._apply_rejection_policy(task)
            
            # Add task to heap
            entry = (task.effective_priority, task.created_at, task)
            heapq.heappush(self._heap, entry)
            self._task_index[task.task_id] = task
            self.enqueued_count += 1
            
            logger.debug(f"Enqueued task {task.task_id} with priority {task.effective_priority}")
            self._not_empty.notify()
            return True
    
    async def _apply_rejection_policy(self, task: Task) -> bool:
        """
        Apply custom rejection policy when queue is full.
        
        Policy: Drop lowest priority tasks that haven't aged yet.
        If no such tasks exist, reject the incoming task.
        """
        async with self._lock:
            # Find lowest priority non-aged tasks
            candidates = []
            for i, (priority, timestamp, t) in enumerate(self._heap):
                age_seconds = time.time() - t.created_at
                if age_seconds < 10 and priority >= task.effective_priority:
                    candidates.append((i, priority, t))
            
            if candidates:
                # Remove the lowest priority candidate
                candidates.sort(key=lambda x: x[1], reverse=True)
                idx, _, removed_task = candidates[0]
                
                # Remove from heap (inefficient but safe)
                self._heap[idx] = self._heap[-1]
                self._heap.pop()
                if self._heap and idx < len(self._heap):
                    heapq.heapify(self._heap)
                
                del self._task_index[removed_task.task_id]
                removed_task.status = TaskStatus.REJECTED
                self.rejected_count += 1
                
                logger.warning(
                    f"Backpressure: dropped task {removed_task.task_id} "
                    f"to accept {task.task_id}"
                )
                
                # Now add the new task
                entry = (task.effective_priority, task.created_at, task)
                heapq.heappush(self._heap, entry)
                self._task_index[task.task_id] = task
                self.enqueued_count += 1
                return True
            
            # Reject incoming task
            task.status = TaskStatus.REJECTED
            self.rejected_count += 1
            logger.warning(f"Backpressure: rejected task {task.task_id}")
            return False
    
    async def dequeue(self, timeout: Optional[float] = None) -> Optional[Task]:
        """
        Remove and return highest priority task.
        
        Returns None if timeout expires with empty queue.
        """
        async with self._not_empty:
            while not self._heap:
                if timeout is not None:
                    try:
                        await asyncio.wait_for(
                            self._not_empty.wait(),
                            timeout=timeout
                        )
                    except asyncio.TimeoutError:
                        return None
                else:
                    await self._not_empty.wait()
            
            # Get highest priority task
            _, _, task = heapq.heappop(self._heap)
            del self._task_index[task.task_id]
            self.dequeued_count += 1
            
            self._not_full.notify()
            return task
    
    async def peek(self) -> Optional[Task]:
        """Peek at highest priority task without removing"""
        async with self._lock:
            if self._heap:
                return self._heap[0][2]
            return None
    
    def size(self) -> int:
        """Current queue size"""
        return len(self._heap)
    
    async def get_stats(self) -> Dict[str, Any]:
        """Get queue statistics"""
        async with self._lock:
            priority_counts = defaultdict(int)
            for priority, _, _ in self._heap:
                priority_counts[priority] += 1
            
            return {
                'size': len(self._heap),
                'max_size': self.MAX_QUEUE_SIZE,
                'enqueued': self.enqueued_count,
                'dequeued': self.dequeued_count,
                'rejected': self.rejected_count,
                'aged': self.aged_count,
                'by_priority': dict(priority_counts),
                'utilization': len(self._heap) / self.MAX_QUEUE_SIZE
            }


class TaskProcessor:
    """
    Main task processor with worker pool, rate limiting, and circuit breakers.
    
    Concurrency Control:
    - Configurable worker count
    - Per-task-type rate limiting via token buckets
    - Circuit breakers prevent cascade failures
    """
    
    def __init__(
        self,
        num_workers: int = 10,
        default_bucket_capacity: int = 100,
        default_refill_rate: float = 10.0
    ):
        self.num_workers = num_workers
        self.queue = TaskQueue()
        self.dlq = DeadLetterQueue()
        self.load_monitor = SystemLoadMonitor()
        
        # Rate limiters per task type
        self._token_buckets: Dict[str, TokenBucket] = {}
        self._bucket_lock = asyncio.Lock()
        self._default_capacity = default_bucket_capacity
        self._default_refill_rate = default_refill_rate
        
        # Circuit breakers per task type
        self._circuit_breakers: Dict[str, CircuitBreaker] = {}
        self._cb_lock = asyncio.Lock()
        
        # Task handlers
        self._handlers: Dict[str, Callable[[Task], Coroutine[Any, Any, Any]]] = {}
        self._fallback_handlers: Dict[str, Callable[[Task], Coroutine[Any, Any, Any]]] = {}
        
        # Worker management
        self._workers: List[asyncio.Task] = []
        self._running = False
        self._tasks_processing = 0
        self._processing_lock = asyncio.Lock()
        
        # Statistics
        self.completed_count = 0
        self.failed_count = 0
        self.fallback_count = 0
    
    def register_handler(
        self,
        task_type: str,
        handler: Callable[[Task], Coroutine[Any, Any, Any]],
        fallback: Optional[Callable[[Task], Coroutine[Any, Any, Any]]] = None
    ):
        """Register a handler for a specific task type"""
        self._handlers[task_type] = handler
        if fallback:
            self._fallback_handlers[task_type] = fallback
    
    async def _get_token_bucket(self, task_type: str) -> TokenBucket:
        """Get or create token bucket for task type"""
        async with self._bucket_lock:
            if task_type not in self._token_buckets:
                self._token_buckets[task_type] = TokenBucket(
                    task_type=task_type,
                    capacity=self._default_capacity,
                    refill_rate=self._default_refill_rate
                )
            return self._token_buckets[task_type]
    
    async def _get_circuit_breaker(self, task_type: str) -> CircuitBreaker:
        """Get or create circuit breaker for task type"""
        async with self._cb_lock:
            if task_type not in self._circuit_breakers:
                cb = CircuitBreaker(
                    task_type=task_type,
                    failure_threshold=3,
                    recovery_timeout=30.0
                )
                if task_type in self._fallback_handlers:
                    cb.fallback_handler = self._fallback_handlers[task_type]
                self._circuit_breakers[task_type] = cb
            return self._circuit_breakers[task_type]
    
    async def _update_rate_limits(self):
        """Background task to update rate limits based on system load"""
        while self._running:
            cpu, memory = await self.load_monitor.get_load()
            
            async with self._bucket_lock:
                for bucket in self._token_buckets.values():
                    bucket.update_rate_based_on_load(cpu, memory)
            
            await asyncio.sleep(5)
    
    async def _worker(self, worker_id: int):
        """Worker coroutine that processes tasks from the queue"""
        logger.info(f"Worker {worker_id} started")
        
        while self._running:
            try:
                # Dequeue task with timeout
                task = await self.queue.dequeue(timeout=1.0)
                if not task:
                    continue
                
                async with self._processing_lock:
                    self._tasks_processing += 1
                
                try:
                    await self._process_task(task, worker_id)
                finally:
                    async with self._processing_lock:
                        self._tasks_processing -= 1
                
            except asyncio.CancelledError:
                break
            except Exception as e:
                logger.error(f"Worker {worker_id} error: {e}")
        
        logger.info(f"Worker {worker_id} stopped")
    
    async def _process_task(self, task: Task, worker_id: int):
        """Process a single task with rate limiting and circuit breaker"""
        task.status = TaskStatus.RUNNING
        logger.debug(f"Worker {worker_id} processing task {task.task_id}")
        
        # Get rate limiter and circuit breaker
        bucket = await self._get_token_bucket(task.task_type)
        circuit_breaker = await self._get_circuit_breaker(task.task_type)
        
        # Wait for rate limit token
        if not await bucket.wait_for_token(timeout=30.0):
            logger.warning(f"Rate limit exceeded for {task.task_type}, requeuing")
            task.status = TaskStatus.PENDING
            await self.queue.enqueue(task)
            return
        
        # Get handler
        handler = self._handlers.get(task.task_type)
        if not handler:
            handler = self._handlers.get('default', self._default_handler)
        
        # Execute with circuit breaker
        try:
            result = await circuit_breaker.execute_with_fallback(task, handler)
            task.status = TaskStatus.COMPLETED
            self.completed_count += 1
            logger.info(f"Task {task.task_id} completed successfully")
            
        except Exception as e:
            task.retry_count += 1
            task.error_log.append(f"{datetime.now()}: {str(e)}\n{traceback.format_exc()}")
            
            if task.retry_count >= task.max_retries:
                # Move to DLQ
                await self.dlq.add(task, f"Max retries exceeded: {str(e)}")
                self.failed_count += 1
                logger.error(f"Task {task.task_id} failed permanently")
            else:
                # Retry with exponential backoff
                task.status = TaskStatus.RETRYING
                backoff = min(30.0, 2 ** task.retry_count)
                await asyncio.sleep(backoff)
                await self.queue.enqueue(task)
                logger.info(f"Task {task.task_id} scheduled for retry #{task.retry_count}")
    
    async def _default_handler(self, task: Task) -> Any:
        """Default task handler (simulates work)"""
        await asyncio.sleep(random.uniform(0.1, 0.5))
        return {'status': 'processed', 'task_id': task.task_id}
    
    async def start(self):
        """Start the processor and all workers"""
        self._running = True
        
        # Start queue
        await self.queue.start()
        
        # Start workers
        for i in range(self.num_workers):
            worker = asyncio.create_task(self._worker(i))
            self._workers.append(worker)
        
        # Start background tasks
        asyncio.create_task(self._update_rate_limits())
        asyncio.create_task(self.load_monitor.simulate_load_fluctuation())
        
        logger.info(f"TaskProcessor started with {self.num_workers} workers")
    
    async def stop(self):
        """Stop the processor and all workers"""
        self._running = False
        
        # Stop queue
        await self.queue.stop()
        
        # Stop workers
        for worker in self._workers:
            worker.cancel()
        
        if self._workers:
            await asyncio.gather(*self._workers, return_exceptions=True)
        
        self._workers.clear()
        logger.info("TaskProcessor stopped")
    
    async def submit_task(
        self,
        task_type: str,
        payload: Dict[str, Any],
        priority: Priority = Priority.MEDIUM,
        timeout: Optional[float] = None
    ) -> Optional[str]:
        """Submit a new task to the queue"""
        task = Task(
            task_type=task_type,
            payload=payload,
            original_priority=priority,
            effective_priority=int(priority)
        )
        
        success = await self.queue.enqueue(task, timeout=timeout)
        if success:
            logger.info(f"Submitted task {task.task_id} of type {task_type}")
            return task.task_id
        return None
    
    async def get_stats(self) -> Dict[str, Any]:
        """Get comprehensive system statistics"""
        queue_stats = await self.queue.get_stats()
        dlq_tasks = await self.dlq.get_all()
        
        async with self._processing_lock:
            tasks_processing = self._tasks_processing
        
        return {
            'queue': queue_stats,
            'dlq_size': len(dlq_tasks),
            'workers': {
                'count': self.num_workers,
                'active': len([w for w in self._workers if not w.done()]),
                'processing': tasks_processing
            },
            'counters': {
                'completed': self.completed_count,
                'failed': self.failed_count,
                'fallback_executed': self.fallback_count
            },
            'circuit_breakers': {
                k: v.state.value for k, v in self._circuit_breakers.items()
            },
            'rate_limiters': {
                k: {'tokens': v.tokens, 'rate': v.refill_rate}
                for k, v in self._token_buckets.items()
            }
        }


# Example usage and demonstration
async def example_task_handler(task: Task) -> Dict[str, Any]:
    """Example task handler that simulates different behaviors"""
    # Simulate occasional failures
    if random.random() < 0.2:
        raise Exception("Simulated random failure")
    
    # Simulate processing time
    await asyncio.sleep(random.uniform(0.1, 0.3))
    
    return {
        'result': f"Processed {task.task_type}",
        'payload_size': len(str(task.payload)),
        'timestamp': time.time()
    }


async def example_fallback_handler(task: Task) -> Dict[str, Any]:
    """Fallback handler when circuit breaker opens"""
    logger.warning(f"Fallback executed for task {task.task_id}")
    return {'fallback': True, 'task_id': task.task_id}


async def main():
    """Demonstration of the task queue system"""
    print("=" * 70)
    print("Async Event-Driven Task Queue System Demo")
    print("=" * 70)
    
    # Create processor
    processor = TaskProcessor(num_workers=5)
    
    # Register handlers
    processor.register_handler('email', example_task_handler, example_fallback_handler)
    processor.register_handler('notification', example_task_handler)
    processor.register_handler('data_processing', example_task_handler, example_fallback_handler)
    processor.register_handler('default', example_task_handler)
    
    # Start processor
    await processor.start()
    
    # Submit various tasks
    print("\nSubmitting tasks...")
    task_ids = []
    
    # High priority tasks
    for i in range(5):
        tid = await processor.submit_task(
            'email',
            {'to': f'user{i}@example.com', 'subject': f'Test {i}'},
            priority=Priority.HIGH
        )
        if tid:
            task_ids.append(tid)
    
    # Medium priority tasks
    for i in range(10):
        tid = await processor.submit_task(
            'notification',
            {'user_id': i, 'message': f'Notification {i}'},
            priority=Priority.MEDIUM
        )
        if tid:
            task_ids.append(tid)
    
    # Low priority tasks
    for i in range(15):
        tid = await processor.submit_task(
            'data_processing',
            {'data': list(range(100)), 'operation': 'transform'},
            priority=Priority.LOW
        )
        if tid:
            task_ids.append(tid)
    
    print(f"Submitted {len(task_ids)} tasks")
    
    # Let tasks process
    print("\nProcessing tasks...")
    for i in range(10):
        await asyncio.sleep(1)
        stats = await processor.get_stats()
        print(f"\n[Second {i+1}] Queue Stats:")
        print(f"  Size: {stats['queue']['size']}/{stats['queue']['max_size']}")
        print(f"  Completed: {stats['counters']['completed']}")
        print(f"  Failed: {stats['counters']['failed']}")
        print(f"  DLQ Size: {stats['dlq_size']}")
        print(f"  Workers Active: {stats['workers']['active']}")
        print(f"  Tasks Processing: {stats['workers']['processing']}")
        
        if stats['circuit_breakers']:
            print(f"  Circuit Breakers: {stats['circuit_breakers']}")
    
    # Show final statistics
    print("\n" + "=" * 70)
    print("Final Statistics")
    print("=" * 70)
    final_stats = await processor.get_stats()
    print(json.dumps(final_stats, indent=2, default=str))
    
    # Show DLQ contents if any
    dlq_tasks = await processor.dlq.get_all()
    if dlq_tasks:
        print(f"\nDead Letter Queue ({len(dlq_tasks)} tasks):")
        for task_dict in dlq_tasks[:5]:  # Show first 5
            print(f"  - {task_dict['task_id']}: {task_dict.get('metadata', {}).get('dlq_reason', 'Unknown')}")
    
    # Stop processor
    await processor.stop()
    print("\nProcessor stopped.")


if __name__ == '__main__':
    asyncio.run(main())
