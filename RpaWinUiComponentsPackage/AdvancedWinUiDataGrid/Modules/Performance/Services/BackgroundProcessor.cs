using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Services;

/// <summary>
/// Background processing service s cancellation tokens pre non-critical operations
/// Implementuje Background Processing stratégiu z newProject.md
/// </summary>
public class BackgroundProcessor : IDisposable
{
    private readonly ILogger? _logger;
    private readonly GridThrottlingConfig _config;
    private readonly ConcurrentQueue<BackgroundTask> _taskQueue = new();
    private readonly SemaphoreSlim _processingSemaphore;
    private readonly CancellationTokenSource _globalCancellation = new();
    private readonly Task _processingTask;
    private readonly Timer _monitoringTimer;
    private volatile bool _disposed;
    
    // Statistics
    private long _totalTasksProcessed;
    private long _totalTasksFailed;
    private long _totalTasksCancelled;

    public BackgroundProcessor(GridThrottlingConfig config, ILogger? logger = null)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _logger = logger;
        
        // Limit concurrent background tasks
        _processingSemaphore = new SemaphoreSlim(Environment.ProcessorCount, Environment.ProcessorCount);
        
        // Start background processing task
        _processingTask = Task.Run(ProcessTaskQueueAsync);
        
        // Setup monitoring timer
        _monitoringTimer = new Timer(LogStatistics, null, 
            TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
        
        _logger?.LogInformation("⚡ BACKGROUND: BackgroundProcessor initialized with {MaxConcurrency} concurrent tasks",
            Environment.ProcessorCount);
    }

    /// <summary>
    /// Pridá task do background queue
    /// </summary>
    public void EnqueueTask(BackgroundTask task)
    {
        try
        {
            if (_disposed) throw new ObjectDisposedException(nameof(BackgroundProcessor));
            if (task == null) throw new ArgumentNullException(nameof(task));

            if (!_config.EnableBackgroundProcessing)
            {
                _logger?.LogTrace("⚡ BACKGROUND: Background processing disabled, skipping task '{TaskName}'", task.Name);
                return;
            }

            task.QueuedAt = DateTime.UtcNow;
            _taskQueue.Enqueue(task);
            
            _logger?.LogTrace("⚡ BACKGROUND: Task '{TaskName}' queued - Priority: {Priority}, Queue size: {QueueSize}",
                task.Name, task.Priority, _taskQueue.Count);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 BACKGROUND ERROR: Failed to enqueue task - TaskName: '{TaskName}', QueueSize: {QueueSize}", 
                task?.Name ?? "null", _taskQueue.Count);
            throw;
        }
    }

    /// <summary>
    /// Pridá async lambda task do background queue
    /// </summary>
    public void EnqueueTask(
        string name, 
        Func<CancellationToken, Task> taskFunc, 
        BackgroundTaskPriority priority = BackgroundTaskPriority.Normal)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(BackgroundProcessor));

        var task = new BackgroundTask
        {
            Name = name,
            TaskFunction = taskFunc,
            Priority = priority,
            MaxRetries = 3
        };

        EnqueueTask(task);
    }

    /// <summary>
    /// Pridá sync lambda task do background queue
    /// </summary>
    public void EnqueueTask(
        string name, 
        Action<CancellationToken> taskAction, 
        BackgroundTaskPriority priority = BackgroundTaskPriority.Normal)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(BackgroundProcessor));

        var task = new BackgroundTask
        {
            Name = name,
            TaskFunction = async (ct) => 
            {
                await Task.Run(() => taskAction(ct), ct);
            },
            Priority = priority,
            MaxRetries = 3
        };

        EnqueueTask(task);
    }

    /// <summary>
    /// Vyčakajú completion všetkých pending tasks
    /// </summary>
    public async Task WaitForCompletionAsync(TimeSpan timeout = default)
    {
        if (_disposed) return;

        var timeoutCts = new CancellationTokenSource(timeout == default ? TimeSpan.FromMinutes(5) : timeout);
        
        try
        {
            while (!_taskQueue.IsEmpty && !timeoutCts.Token.IsCancellationRequested)
            {
                await Task.Delay(100, timeoutCts.Token);
            }
            
            _logger?.LogInformation("⚡ BACKGROUND: All tasks completed");
        }
        catch (OperationCanceledException)
        {
            _logger?.LogWarning("⚠️ BACKGROUND: Wait for completion timed out - Remaining tasks: {RemainingCount}",
                _taskQueue.Count);
        }
    }

    /// <summary>
    /// Získa aktuálne background processing statistics
    /// </summary>
    public BackgroundProcessorStatistics GetStatistics()
    {
        return new BackgroundProcessorStatistics
        {
            QueuedTasks = _taskQueue.Count,
            TotalTasksProcessed = Interlocked.Read(ref _totalTasksProcessed),
            TotalTasksFailed = Interlocked.Read(ref _totalTasksFailed),
            TotalTasksCancelled = Interlocked.Read(ref _totalTasksCancelled),
            IsEnabled = _config.EnableBackgroundProcessing,
            MaxConcurrentTasks = Environment.ProcessorCount
        };
    }

    /// <summary>
    /// Main processing loop pre background tasks
    /// </summary>
    private async Task ProcessTaskQueueAsync()
    {
        _logger?.LogInformation("⚡ BACKGROUND: Processing loop started");

        try
        {
            while (!_globalCancellation.Token.IsCancellationRequested)
            {
                if (_taskQueue.TryDequeue(out var task))
                {
                    // Wait for available processing slot
                    await _processingSemaphore.WaitAsync(_globalCancellation.Token);
                    
                    // Process task in background (fire and forget)
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await ProcessTaskAsync(task);
                        }
                        finally
                        {
                            _processingSemaphore.Release();
                        }
                    });
                }
                else
                {
                    // No tasks available, wait a bit
                    await Task.Delay(50, _globalCancellation.Token);
                }
            }
        }
        catch (OperationCanceledException)
        {
            _logger?.LogInformation("⚡ BACKGROUND: Processing loop cancelled");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 BACKGROUND ERROR: Processing loop failed");
        }
    }

    /// <summary>
    /// Process individual background task
    /// </summary>
    private async Task ProcessTaskAsync(BackgroundTask task)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        task.StartedAt = DateTime.UtcNow;
        
        _logger?.LogTrace("⚡ BACKGROUND: Starting task '{TaskName}' - Priority: {Priority}, " +
            "Queued for: {QueuedMs}ms", task.Name, task.Priority, 
            (task.StartedAt - task.QueuedAt).TotalMilliseconds);

        var attempt = 0;
        while (attempt <= task.MaxRetries && !_globalCancellation.Token.IsCancellationRequested)
        {
            attempt++;
            
            try
            {
                using var taskCts = CancellationTokenSource.CreateLinkedTokenSource(_globalCancellation.Token);
                
                // Set timeout for individual task
                if (task.TimeoutMs > 0)
                {
                    taskCts.CancelAfter(TimeSpan.FromMilliseconds(task.TimeoutMs));
                }

                await task.TaskFunction(taskCts.Token);
                
                // Success
                task.CompletedAt = DateTime.UtcNow;
                stopwatch.Stop();
                
                Interlocked.Increment(ref _totalTasksProcessed);
                
                _logger?.LogTrace("⚡ BACKGROUND: Task '{TaskName}' completed successfully in {ElapsedMs}ms " +
                    "(attempt {Attempt}/{MaxRetries})", task.Name, stopwatch.ElapsedMilliseconds, 
                    attempt, task.MaxRetries + 1);
                
                return; // Success, exit retry loop
            }
            catch (OperationCanceledException) when (_globalCancellation.Token.IsCancellationRequested)
            {
                // Global cancellation
                Interlocked.Increment(ref _totalTasksCancelled);
                _logger?.LogTrace("⚡ BACKGROUND: Task '{TaskName}' cancelled globally", task.Name);
                return;
            }
            catch (OperationCanceledException)
            {
                // Task timeout
                if (attempt <= task.MaxRetries)
                {
                    var delay = TimeSpan.FromMilliseconds(1000 * attempt); // Exponential backoff
                    _logger?.LogWarning("⚠️ BACKGROUND: Task '{TaskName}' timed out (attempt {Attempt}/{MaxRetries}) - " +
                        "Retrying in {DelayMs}ms", task.Name, attempt, task.MaxRetries + 1, delay.TotalMilliseconds);
                    
                    await Task.Delay(delay, _globalCancellation.Token);
                }
                else
                {
                    Interlocked.Increment(ref _totalTasksFailed);
                    _logger?.LogError("🚨 BACKGROUND ERROR: Task '{TaskName}' failed after {Attempts} attempts - timeout",
                        task.Name, attempt);
                }
            }
            catch (Exception ex)
            {
                if (attempt <= task.MaxRetries)
                {
                    var delay = TimeSpan.FromMilliseconds(1000 * attempt); // Exponential backoff
                    _logger?.LogWarning(ex, "⚠️ BACKGROUND: Task '{TaskName}' failed (attempt {Attempt}/{MaxRetries}) - " +
                        "Retrying in {DelayMs}ms", task.Name, attempt, task.MaxRetries + 1, delay.TotalMilliseconds);
                    
                    await Task.Delay(delay, _globalCancellation.Token);
                }
                else
                {
                    Interlocked.Increment(ref _totalTasksFailed);
                    _logger?.LogError(ex, "🚨 BACKGROUND ERROR: Task '{TaskName}' failed after {Attempts} attempts",
                        task.Name, attempt);
                }
            }
        }
        
        task.CompletedAt = DateTime.UtcNow;
    }

    /// <summary>
    /// Periodic statistics logging
    /// </summary>
    private void LogStatistics(object? state)
    {
        if (_disposed) return;

        try
        {
            var stats = GetStatistics();
            
            if (stats.TotalTasksProcessed > 0 || stats.QueuedTasks > 0)
            {
                _logger?.LogInformation("⚡ BACKGROUND: Statistics - " +
                    "Queued: {Queued}, Processed: {Processed}, Failed: {Failed}, Cancelled: {Cancelled}, " +
                    "Success rate: {SuccessRate:P2}",
                    stats.QueuedTasks, stats.TotalTasksProcessed, stats.TotalTasksFailed, stats.TotalTasksCancelled,
                    stats.SuccessRate);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 BACKGROUND ERROR: Statistics logging failed");
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            _logger?.LogInformation("⚡ BACKGROUND: Shutting down BackgroundProcessor...");
            
            // Cancel all tasks
            _globalCancellation.Cancel();
            
            // Wait for processing task to complete
            _processingTask?.Wait(TimeSpan.FromSeconds(5));
            
            // Dispose resources
            _monitoringTimer?.Dispose();
            _globalCancellation?.Dispose();
            _processingTask?.Dispose();
            _processingSemaphore?.Dispose();
            
            var remainingTasks = _taskQueue.Count;
            if (remainingTasks > 0)
            {
                _logger?.LogWarning("⚠️ BACKGROUND: Disposed with {RemainingTasks} unprocessed tasks", remainingTasks);
            }
            
            _logger?.LogInformation("⚡ BACKGROUND: BackgroundProcessor disposed");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 BACKGROUND ERROR: BackgroundProcessor disposal failed");
        }
    }
}

/// <summary>
/// Background task definition
/// </summary>
public class BackgroundTask
{
    public string Name { get; set; } = string.Empty;
    public Func<CancellationToken, Task> TaskFunction { get; set; } = _ => Task.CompletedTask;
    public BackgroundTaskPriority Priority { get; set; } = BackgroundTaskPriority.Normal;
    public int MaxRetries { get; set; } = 3;
    public int TimeoutMs { get; set; } = 30000; // 30 seconds default
    
    // Timestamps
    public DateTime QueuedAt { get; set; }
    public DateTime StartedAt { get; set; }
    public DateTime CompletedAt { get; set; }
    
    // Computed properties
    public TimeSpan QueueTime => StartedAt - QueuedAt;
    public TimeSpan ExecutionTime => CompletedAt - StartedAt;
    public bool IsCompleted => CompletedAt != default;
}

/// <summary>
/// Background task priority
/// </summary>
public enum BackgroundTaskPriority
{
    Low,
    Normal,
    High,
    Critical
}

/// <summary>
/// Background processor statistics
/// </summary>
public class BackgroundProcessorStatistics
{
    public int QueuedTasks { get; init; }
    public long TotalTasksProcessed { get; init; }
    public long TotalTasksFailed { get; init; }
    public long TotalTasksCancelled { get; init; }
    public bool IsEnabled { get; init; }
    public int MaxConcurrentTasks { get; init; }
    
    public long TotalTasksAttempted => TotalTasksProcessed + TotalTasksFailed + TotalTasksCancelled;
    public double SuccessRate => TotalTasksAttempted > 0 ? (double)TotalTasksProcessed / TotalTasksAttempted : 0.0;
    public double FailureRate => TotalTasksAttempted > 0 ? (double)TotalTasksFailed / TotalTasksAttempted : 0.0;
}