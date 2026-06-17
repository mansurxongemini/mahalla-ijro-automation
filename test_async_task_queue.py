"""
Comprehensive Test Suite for Async Task Queue System
=====================================================

Tests all major features:
1. Priority handling and aging mechanism
2. Rate limiting with token buckets
3. Circuit breaker pattern
4. Dead Letter Queue
5. Backpressure handling
6. Concurrency and thread safety
"""

import asyncio
import time
import pytest
from async_task_queue import (
    Task, TaskQueue, TaskProcessor, TaskStatus, Priority,
    TokenBucket, CircuitBreaker, CircuitState, DeadLetterQueue,
    SystemLoadMonitor
)


class TestPriority:
    """Test priority levels and aging"""
    
    def test_priority_ordering(self):
        """Verify priority ordering (lower value = higher priority)"""
        assert Priority.HIGH < Priority.MEDIUM < Priority.LOW
        assert int(Priority.HIGH) == 0
        assert int(Priority.MEDIUM) == 1
        assert int(Priority.LOW) == 2
    
    def test_priority_aging(self):
        """Test priority aging mechanism"""
        assert Priority.LOW.age() == Priority.MEDIUM
        assert Priority.MEDIUM.age() == Priority.HIGH
        assert Priority.HIGH.age() == Priority.HIGH  # HIGH stays HIGH


class TestTask:
    """Test Task dataclass functionality"""
    
    def test_task_creation(self):
        """Test basic task creation"""
        task = Task(
            effective_priority=1,
            task_type="test",
            payload={"key": "value"},
            original_priority=Priority.MEDIUM
        )
        assert task.task_type == "test"
        assert task.payload == {"key": "value"}
        assert task.status == TaskStatus.PENDING
        assert task.retry_count == 0
    
    def test_task_serialization(self):
        """Test task serialization to dictionary"""
        task = Task(
            effective_priority=0,
            task_type="email",
            payload={"to": "user@example.com"},
            original_priority=Priority.HIGH
        )
        task_dict = task.to_dict()
        
        assert task_dict['task_type'] == "email"
        assert task_dict['original_priority'] == "HIGH"
        assert 'task_id' in task_dict
        assert 'created_at' in task_dict
    
    def test_task_state_serialization(self):
        """Test complete state serialization for DLQ"""
        task = Task(
            effective_priority=1,
            task_type="test",
            payload={"data": [1, 2, 3]},
            original_priority=Priority.MEDIUM
        )
        task.error_log.append("Error 1")
        task.metadata['custom'] = 'value'
        
        serialized = task.serialize_state()
        assert '"task_type": "test"' in serialized
        assert '"Error 1"' in serialized
        assert '"custom": "value"' in serialized
    
    @pytest.mark.asyncio
    async def test_task_aging_calculation(self):
        """Test effective priority calculation with aging"""
        task = Task(
            effective_priority=2,  # LOW
            task_type="test",
            original_priority=Priority.LOW
        )
        
        # Initially LOW priority
        assert task.get_effective_priority() == 2
        
        # Simulate aging by manipulating created_at
        task.created_at = time.time() - 15  # 15 seconds ago
        
        # Should age to MEDIUM after 10+ seconds
        effective = task.get_effective_priority()
        assert effective == 1  # MEDIUM
        
        task.created_at = time.time() - 25  # 25 seconds ago
        effective = task.get_effective_priority()
        assert effective == 0  # HIGH (aged twice)


class TestTokenBucket:
    """Test token bucket rate limiter"""
    
    @pytest.mark.asyncio
    async def test_token_consumption(self):
        """Test basic token consumption"""
        bucket = TokenBucket(task_type="test", capacity=10, tokens=10.0)
        
        # Should be able to consume tokens
        assert await bucket.consume(5) is True
        # Allow small tolerance for timing
        assert abs(bucket.tokens - 5.0) < 0.1
        
        # Should fail when not enough tokens
        assert await bucket.consume(6) is False
        assert await bucket.consume(5) is True
        # Tokens should be near zero (allowing for refill)
        assert bucket.tokens < 0.5
    
    @pytest.mark.asyncio
    async def test_token_refill(self):
        """Test automatic token refill"""
        bucket = TokenBucket(
            task_type="test",
            capacity=10,
            tokens=0,
            refill_rate=100.0  # Fast refill for testing
        )
        
        # Wait for refill
        await asyncio.sleep(0.1)
        
        # Should have refilled
        bucket._refill()
        assert bucket.tokens > 0
    
    @pytest.mark.asyncio
    async def test_wait_for_token(self):
        """Test waiting for tokens"""
        bucket = TokenBucket(
            task_type="test",
            capacity=10,
            tokens=0,
            refill_rate=50.0
        )
        
        # Should eventually get token
        result = await bucket.wait_for_token(timeout=1.0)
        assert result is True
    
    @pytest.mark.asyncio
    async def test_load_based_rate_adjustment(self):
        """Test dynamic rate adjustment based on load"""
        bucket = TokenBucket(task_type="test", refill_rate=10.0)
        
        # High load should reduce rate
        bucket.update_rate_based_on_load(90.0, 85.0)
        assert bucket.refill_rate < 10.0
        
        # Low load should increase rate
        initial_rate = bucket.refill_rate
        bucket.update_rate_based_on_load(10.0, 15.0)
        assert bucket.refill_rate > initial_rate


class TestCircuitBreaker:
    """Test circuit breaker pattern"""
    
    @pytest.mark.asyncio
    async def test_circuit_closed_state(self):
        """Test normal operation with closed circuit"""
        cb = CircuitBreaker(task_type="test", failure_threshold=3)
        
        assert cb.state == CircuitState.CLOSED
        assert await cb.can_execute() is True
    
    @pytest.mark.asyncio
    async def test_circuit_opens_after_failures(self):
        """Test circuit opens after threshold failures"""
        cb = CircuitBreaker(task_type="test", failure_threshold=3)
        
        # Record failures
        assert await cb.record_failure() is False  # 1
        assert await cb.record_failure() is False  # 2
        assert await cb.record_failure() is True   # 3 - opens!
        
        assert cb.state == CircuitState.OPEN
        assert await cb.can_execute() is False
    
    @pytest.mark.asyncio
    async def test_circuit_half_open(self):
        """Test circuit transitions to half-open after timeout"""
        cb = CircuitBreaker(
            task_type="test",
            failure_threshold=1,
            recovery_timeout=0.1
        )
        
        # Open the circuit
        await cb.record_failure()
        assert cb.state == CircuitState.OPEN
        
        # Wait for recovery timeout
        await asyncio.sleep(0.15)
        
        # Should transition to half-open
        assert await cb.can_execute() is True
        assert cb.state == CircuitState.HALF_OPEN
    
    @pytest.mark.asyncio
    async def test_success_resets_circuit(self):
        """Test successful execution resets circuit"""
        cb = CircuitBreaker(task_type="test", failure_threshold=3)
        
        await cb.record_failure()
        await cb.record_failure()
        assert cb.failure_count == 2
        
        await cb.record_success()
        assert cb.failure_count == 0
        assert cb.state == CircuitState.CLOSED
    
    @pytest.mark.asyncio
    async def test_execute_with_fallback(self):
        """Test fallback execution when circuit open"""
        fallback_called = False
        
        async def handler(task):
            raise Exception("Handler failed")
        
        async def fallback(task):
            nonlocal fallback_called
            fallback_called = True
            return {"fallback": True}
        
        cb = CircuitBreaker(
            task_type="test",
            failure_threshold=1,
            recovery_timeout=30.0
        )
        cb.fallback_handler = fallback
        
        task = Task(effective_priority=1, task_type="test")
        
        # First call opens circuit and executes fallback
        result = await cb.execute_with_fallback(task, handler)
        assert result == {"fallback": True}
        assert fallback_called is True
        assert cb.state == CircuitState.OPEN


class TestDeadLetterQueue:
    """Test dead letter queue functionality"""
    
    @pytest.mark.asyncio
    async def test_add_to_dlq(self):
        """Test adding tasks to DLQ"""
        dlq = DeadLetterQueue()
        task = Task(effective_priority=1, task_type="test")
        
        await dlq.add(task, "Test failure reason")
        
        assert dlq.size() == 1
        assert task.status == TaskStatus.DEAD_LETTER
        assert task.metadata['dlq_reason'] == "Test failure reason"
    
    @pytest.mark.asyncio
    async def test_dlq_serialization(self):
        """Test DLQ task retrieval as dictionaries"""
        dlq = DeadLetterQueue()
        task = Task(
            effective_priority=1,
            task_type="email",
            payload={"to": "user@example.com"}
        )
        
        await dlq.add(task, "Max retries exceeded")
        
        tasks = await dlq.get_all()
        assert len(tasks) == 1
        assert tasks[0]['task_type'] == "email"
        assert tasks[0]['status'] == "dead_letter"
    
    @pytest.mark.asyncio
    async def test_dlq_replay(self):
        """Test removing task from DLQ for replay"""
        dlq = DeadLetterQueue()
        task = Task(effective_priority=1, task_type="test")
        
        await dlq.add(task, "Temporary failure")
        assert dlq.size() == 1
        
        replayed = await dlq.replay_task(task.task_id)
        assert replayed is not None
        assert replayed.task_id == task.task_id
        assert dlq.size() == 0
    
    @pytest.mark.asyncio
    async def test_dlq_max_size(self):
        """Test DLQ enforces max size"""
        dlq = DeadLetterQueue(max_size=5)
        
        for i in range(10):
            task = Task(effective_priority=1, task_type="test")
            await dlq.add(task, f"Failure {i}")
        
        # Should only keep last 5
        assert dlq.size() == 5


class TestTaskQueue:
    """Test main task queue functionality"""
    
    @pytest.mark.asyncio
    async def test_enqueue_dequeue(self):
        """Test basic enqueue and dequeue"""
        queue = TaskQueue()
        await queue.start()
        
        try:
            task = Task(effective_priority=1, task_type="test")
            result = await queue.enqueue(task)
            
            assert result is True
            assert queue.size() == 1
            
            dequeued = await queue.dequeue(timeout=1.0)
            assert dequeued is not None
            assert dequeued.task_id == task.task_id
            assert queue.size() == 0
        finally:
            await queue.stop()
    
    @pytest.mark.asyncio
    async def test_priority_ordering(self):
        """Test tasks are dequeued in priority order"""
        queue = TaskQueue()
        await queue.start()
        
        try:
            # Add tasks in reverse priority order
            low_task = Task(effective_priority=2, task_type="test", original_priority=Priority.LOW)
            high_task = Task(effective_priority=0, task_type="test", original_priority=Priority.HIGH)
            medium_task = Task(effective_priority=1, task_type="test", original_priority=Priority.MEDIUM)
            
            await queue.enqueue(low_task)
            await queue.enqueue(high_task)
            await queue.enqueue(medium_task)
            
            # Should dequeue in priority order
            first = await queue.dequeue(timeout=1.0)
            second = await queue.dequeue(timeout=1.0)
            third = await queue.dequeue(timeout=1.0)
            
            assert first.effective_priority == 0  # HIGH
            assert second.effective_priority == 1  # MEDIUM
            assert third.effective_priority == 2  # LOW
        finally:
            await queue.stop()
    
    @pytest.mark.asyncio
    async def test_aging_mechanism(self):
        """Test priority aging over time"""
        queue = TaskQueue()
        queue.AGING_INTERVAL = 0.1  # Fast aging for testing
        await queue.start()
        
        try:
            # Create old low-priority task
            old_task = Task(
                effective_priority=2,
                task_type="test",
                original_priority=Priority.LOW
            )
            old_task.created_at = time.time() - 15  # 15 seconds old
            
            # Create new high-priority task
            new_task = Task(
                effective_priority=0,
                task_type="test",
                original_priority=Priority.HIGH
            )
            
            await queue.enqueue(old_task)
            await queue.enqueue(new_task)
            
            # Wait for aging to process
            await asyncio.sleep(0.2)
            
            # Old task should have aged to higher priority
            stats = await queue.get_stats()
            assert stats['aged'] > 0
        finally:
            await queue.stop()
    
    @pytest.mark.asyncio
    async def test_backpressure_rejection(self):
        """Test backpressure when queue is full"""
        queue = TaskQueue()
        queue.MAX_QUEUE_SIZE = 5
        await queue.start()
        
        try:
            # Fill the queue
            for i in range(5):
                task = Task(effective_priority=1, task_type="test")
                await queue.enqueue(task)
            
            # Try to add more with zero timeout (immediate rejection)
            new_task = Task(effective_priority=1, task_type="test")
            result = await queue.enqueue(new_task, timeout=0.0)
            
            # Should be rejected (or policy applied)
            assert result is False or queue.size() <= 5
        finally:
            await queue.stop()
    
    @pytest.mark.asyncio
    async def test_queue_statistics(self):
        """Test queue statistics tracking"""
        queue = TaskQueue()
        await queue.start()
        
        try:
            for i in range(10):
                task = Task(effective_priority=i % 3, task_type="test")
                await queue.enqueue(task)
            
            for i in range(5):
                await queue.dequeue(timeout=1.0)
            
            stats = await queue.get_stats()
            
            assert stats['size'] == 5
            assert stats['enqueued'] == 10
            assert stats['dequeued'] == 5
            assert stats['utilization'] == 0.0005
        finally:
            await queue.stop()


class TestTaskProcessor:
    """Test integrated task processor"""
    
    @pytest.mark.asyncio
    async def test_processor_lifecycle(self):
        """Test processor start and stop"""
        processor = TaskProcessor(num_workers=2)
        
        await processor.start()
        assert processor._running is True
        assert len(processor._workers) == 2
        
        await processor.stop()
        assert processor._running is False
    
    @pytest.mark.asyncio
    async def test_task_submission_and_processing(self):
        """Test end-to-end task processing"""
        processor = TaskProcessor(num_workers=2)
        
        processed_tasks = []
        
        async def test_handler(task):
            processed_tasks.append(task.task_id)
            return {"processed": True}
        
        processor.register_handler('test', test_handler)
        
        await processor.start()
        
        try:
            # Submit task
            task_id = await processor.submit_task(
                'test',
                {'data': 'value'},
                priority=Priority.HIGH
            )
            
            assert task_id is not None
            
            # Wait for processing
            await asyncio.sleep(1)
            
            assert len(processed_tasks) == 1
            assert processed_tasks[0] == task_id
            
            stats = await processor.get_stats()
            assert stats['counters']['completed'] == 1
        finally:
            await processor.stop()
    
    @pytest.mark.asyncio
    async def test_retry_mechanism(self):
        """Test task retry on failure"""
        processor = TaskProcessor(num_workers=2, default_bucket_capacity=100)
        
        attempt_count = 0
        
        async def failing_handler(task):
            nonlocal attempt_count
            attempt_count += 1
            if attempt_count < 3:
                raise Exception(f"Intentional failure #{attempt_count}")
            return {"success": True}
        
        processor.register_handler('retry_test', failing_handler)
        
        await processor.start()
        
        try:
            task_id = await processor.submit_task(
                'retry_test',
                {},
                priority=Priority.HIGH
            )
            
            # Wait for retries (with exponential backoff: 2^1 + 2^2 = 6 seconds max)
            await asyncio.sleep(8)
            
            # Should succeed on 3rd attempt
            assert attempt_count >= 3
            
            stats = await processor.get_stats()
            assert stats['counters']['completed'] >= 1
        finally:
            await processor.stop()
    
    @pytest.mark.asyncio
    async def test_circuit_breaker_integration(self):
        """Test circuit breaker triggers after failures"""
        processor = TaskProcessor(num_workers=2)
        
        fallback_executed = False
        
        async def always_fails(task):
            raise Exception("Always fails")
        
        async def fallback(task):
            nonlocal fallback_executed
            fallback_executed = True
            return {"fallback": True}
        
        processor.register_handler('failing', always_fails, fallback)
        
        await processor.start()
        
        try:
            # Submit multiple tasks to trigger circuit breaker
            for i in range(5):
                await processor.submit_task('failing', {'id': i})
            
            # Wait for processing and circuit breaker to trigger
            await asyncio.sleep(3)
            
            stats = await processor.get_stats()
            
            # Circuit should be open
            assert stats['circuit_breakers'].get('failing') == 'open'
            
            # Fallback should have been executed at least once
            assert fallback_executed is True
        finally:
            await processor.stop()
    
    @pytest.mark.asyncio
    async def test_dlq_integration(self):
        """Test tasks move to DLQ after max retries"""
        processor = TaskProcessor(num_workers=2, default_bucket_capacity=100)
        
        async def always_fails(task):
            raise Exception("Permanent failure")
        
        processor.register_handler('permanent_fail', always_fails)
        
        await processor.start()
        
        try:
            task_id = await processor.submit_task(
                'permanent_fail',
                {'test': 'data'}
            )
            
            # Wait for all retries and DLQ movement
            await asyncio.sleep(10)
            
            stats = await processor.get_stats()
            
            # Should have failed permanently
            assert stats['counters']['failed'] >= 1
            assert stats['dlq_size'] >= 1
            
            # Verify task in DLQ
            dlq_tasks = await processor.dlq.get_all()
            assert len(dlq_tasks) >= 1
            assert dlq_tasks[0]['status'] == 'dead_letter'
        finally:
            await processor.stop()
    
    @pytest.mark.asyncio
    async def test_rate_limiting_per_task_type(self):
        """Test per-task-type rate limiting"""
        processor = TaskProcessor(
            num_workers=5,
            default_bucket_capacity=5,
            default_refill_rate=1.0
        )
        
        processed_count = 0
        
        async def quick_handler(task):
            nonlocal processed_count
            processed_count += 1
            await asyncio.sleep(0.01)
            return {}
        
        processor.register_handler('rate_limited', quick_handler)
        
        await processor.start()
        
        try:
            # Submit many tasks quickly
            for i in range(10):
                await processor.submit_task('rate_limited', {'id': i})
            
            # Wait a bit
            await asyncio.sleep(1)
            
            stats = await processor.get_stats()
            
            # Should have rate limiter for task type
            assert 'rate_limited' in stats['rate_limiters']
        finally:
            await processor.stop()


class TestConcurrencySafety:
    """Test thread safety and concurrency"""
    
    @pytest.mark.asyncio
    async def test_concurrent_enqueuing(self):
        """Test safe concurrent task submission"""
        queue = TaskQueue()
        await queue.start()
        
        try:
            # Submit many tasks concurrently
            async def submit_task(i):
                task = Task(effective_priority=i % 3, task_type="test")
                return await queue.enqueue(task)
            
            tasks = [submit_task(i) for i in range(100)]
            results = await asyncio.gather(*tasks)
            
            # All should succeed
            assert all(results)
            assert queue.size() == 100
            assert queue.enqueued_count == 100
        finally:
            await queue.stop()
    
    @pytest.mark.asyncio
    async def test_concurrent_dequeuing(self):
        """Test safe concurrent task retrieval"""
        queue = TaskQueue()
        await queue.start()
        
        try:
            # Pre-populate queue
            for i in range(50):
                task = Task(effective_priority=i % 3, task_type="test")
                await queue.enqueue(task)
            
            # Concurrently dequeue
            async def get_task():
                return await queue.dequeue(timeout=1.0)
            
            tasks = [get_task() for _ in range(50)]
            results = await asyncio.gather(*tasks)
            
            # All should get a task
            valid_results = [r for r in results if r is not None]
            assert len(valid_results) == 50
            assert queue.size() == 0
        finally:
            await queue.stop()
    
    @pytest.mark.asyncio
    async def test_worker_pool_concurrency(self):
        """Test multiple workers process tasks safely"""
        processor = TaskProcessor(num_workers=10)
        
        processed_ids = []
        lock = asyncio.Lock()
        
        async def handler(task):
            async with lock:
                processed_ids.append(task.task_id)
            await asyncio.sleep(0.01)
            return {}
        
        processor.register_handler('concurrent', handler)
        
        await processor.start()
        
        try:
            # Submit many tasks
            for i in range(50):
                await processor.submit_task('concurrent', {'id': i})
            
            # Wait for processing
            await asyncio.sleep(3)
            
            stats = await processor.get_stats()
            
            # All tasks should be processed
            assert stats['counters']['completed'] == 50
            assert len(processed_ids) == 50
            
            # No duplicates
            assert len(set(processed_ids)) == 50
        finally:
            await processor.stop()


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
