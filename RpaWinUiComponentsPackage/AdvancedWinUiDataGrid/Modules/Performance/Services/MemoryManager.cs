using System.Collections.Concurrent;
using Microsoft.Extensions.ObjectPool;
using Microsoft.Extensions.Logging;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Services;

/// <summary>
/// Advanced memory management pre DataGrid s ObjectPool, aggressive GC a weak references
/// Implementuje Intelligent Windowing stratégiu podľa newProject.md
/// </summary>
public class MemoryManager : IDisposable
{
    private readonly ILogger? _logger;
    private readonly ObjectPool<Cell> _cellPool;
    private readonly WeakReferenceCache _weakCache;
    private readonly Timer _memoryCleanupTimer;
    private readonly GridThrottlingConfig _config;
    private readonly SemaphoreSlim _gcSemaphore = new(1, 1);
    private volatile bool _disposed;

    // Memory monitoring 
    private DateTime _lastCleanupTime = DateTime.UtcNow;

    public MemoryManager(GridThrottlingConfig config, ILogger? logger = null)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _logger = logger;
        _weakCache = new WeakReferenceCache();
        
        // Configure ObjectPool for Cell objects
        var poolProvider = new DefaultObjectPoolProvider();
        var cellPoolPolicy = new CellPoolPolicy();
        _cellPool = poolProvider.Create(cellPoolPolicy);
        
        // Setup memory cleanup timer
        _memoryCleanupTimer = new Timer(PerformMemoryCleanup, null, 
            TimeSpan.FromMilliseconds(_config.MemoryCleanupIntervalMs),
            TimeSpan.FromMilliseconds(_config.MemoryCleanupIntervalMs));
        
        _logger?.LogInformation("⚡ MEMORY: MemoryManager initialized with cleanup interval {IntervalMs}ms", 
            _config.MemoryCleanupIntervalMs);
    }

    /// <summary>
    /// Získa Cell z ObjectPool pre optimalizovanú performance
    /// </summary>
    public Cell GetCell()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(MemoryManager));
        
        var cell = _cellPool.Get();
        _logger?.LogTrace("⚡ MEMORY: Cell retrieved from pool");
        return cell;
    }

    /// <summary>
    /// Vráti Cell do ObjectPool
    /// </summary>
    public void ReturnCell(Cell cell)
    {
        if (_disposed || cell == null) return;
        
        try
        {
            _cellPool.Return(cell);
            _logger?.LogTrace("⚡ MEMORY: Cell returned to pool");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 MEMORY ERROR: Failed to return cell to pool");
        }
    }

    /// <summary>
    /// Optimalizuje memory usage - aggressive GC + cache cleanup
    /// </summary>
    public async Task OptimizeMemoryAsync()
    {
        if (_disposed) return;

        try
        {
            await _gcSemaphore.WaitAsync();
            
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var beforeMemory = GC.GetTotalMemory(false);
            
            _logger?.LogDebug("⚡ MEMORY: Starting memory optimization - Before: {BeforeMemoryMB}MB", 
                beforeMemory / 1024 / 1024);
            
            // Clean weak reference cache
            var cleanedItems = _weakCache.Cleanup();
            
            // Force garbage collection if aggressive mode enabled
            if (_config.EnableAggressiveMemoryManagement)
            {
                await ForceGarbageCollectionAsync();
            }
            
            var afterMemory = GC.GetTotalMemory(false);
            var savedMemory = beforeMemory - afterMemory;
            
            stopwatch.Stop();
            _logger?.LogInformation("⚡ MEMORY: Optimization completed in {ElapsedMs}ms - " +
                "Before: {BeforeMemoryMB}MB, After: {AfterMemoryMB}MB, Saved: {SavedMemoryMB}MB, " +
                "Cache cleaned: {CleanedItems} items",
                stopwatch.ElapsedMilliseconds, beforeMemory / 1024 / 1024, 
                afterMemory / 1024 / 1024, savedMemory / 1024 / 1024, cleanedItems);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 MEMORY ERROR: Memory optimization failed");
        }
        finally
        {
            _gcSemaphore.Release();
        }
    }

    /// <summary>
    /// Získa detailný memory usage report
    /// </summary>
    public async Task<MemoryReport> GetMemoryUsageAsync()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(MemoryManager));
        
        var totalMemory = GC.GetTotalMemory(false);
        var gen0Collections = GC.CollectionCount(0);
        var gen1Collections = GC.CollectionCount(1);
        var gen2Collections = GC.CollectionCount(2);
        
        var report = new MemoryReport
        {
            TotalMemoryBytes = totalMemory,
            TotalMemoryMB = totalMemory / 1024.0 / 1024.0,
            Gen0Collections = gen0Collections,
            Gen1Collections = gen1Collections,
            Gen2Collections = gen2Collections,
            WeakCacheAliveReferences = _weakCache.GetAliveCount(),
            WeakCacheDeadReferences = _weakCache.GetDeadCount(),
            LastCleanupTime = _lastCleanupTime,
            IsAggressiveModeEnabled = _config.EnableAggressiveMemoryManagement,
            CleanupIntervalMs = _config.MemoryCleanupIntervalMs
        };
        
        _logger?.LogDebug("⚡ MEMORY: Memory report generated - " +
            "Total: {MemoryMB}MB, GC Gen0/1/2: {Gen0}/{Gen1}/{Gen2}, " +
            "WeakCache alive/dead: {Alive}/{Dead}",
            report.TotalMemoryMB, gen0Collections, gen1Collections, gen2Collections,
            report.WeakCacheAliveReferences, report.WeakCacheDeadReferences);
        
        return report;
    }

    /// <summary>
    /// Aggressive garbage collection pre kritické memory situácie
    /// </summary>
    public async Task ForceGarbageCollectionAsync()
    {
        if (_disposed) return;
        
        try
        {
            await Task.Run(() =>
            {
                var beforeMemory = GC.GetTotalMemory(false);
                
                // Force collection of all generations
                GC.Collect(2, GCCollectionMode.Forced, true);
                GC.WaitForPendingFinalizers();
                GC.Collect(2, GCCollectionMode.Forced, true);
                
                // Compact large object heap if available (.NET 4.5.1+)
                GC.Collect();
                
                var afterMemory = GC.GetTotalMemory(false);
                var savedMemory = beforeMemory - afterMemory;
                
                _logger?.LogInformation("⚡ MEMORY: Forced GC completed - " +
                    "Before: {BeforeMemoryMB}MB, After: {AfterMemoryMB}MB, Freed: {SavedMemoryMB}MB",
                    beforeMemory / 1024 / 1024, afterMemory / 1024 / 1024, savedMemory / 1024 / 1024);
            });
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 MEMORY ERROR: Forced garbage collection failed");
        }
    }

    /// <summary>
    /// Periodic memory cleanup callback
    /// </summary>
    private async void PerformMemoryCleanup(object? state)
    {
        if (_disposed) return;
        
        try
        {
            _lastCleanupTime = DateTime.UtcNow;
            await OptimizeMemoryAsync();
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 MEMORY ERROR: Periodic memory cleanup failed");
        }
    }

    /// <summary>
    /// WeakReference cache pre optimalizáciu memory
    /// </summary>
    public WeakReferenceCache WeakCache => _weakCache;

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        
        try
        {
            _memoryCleanupTimer?.Dispose();
            _weakCache?.Dispose();
            _gcSemaphore?.Dispose();
            
            _logger?.LogInformation("⚡ MEMORY: MemoryManager disposed");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 MEMORY ERROR: MemoryManager disposal failed");
        }
    }
}

/// <summary>
/// ObjectPool policy pre Cell objekty
/// </summary>
internal class CellPoolPolicy : IPooledObjectPolicy<Cell>
{
    public Cell Create()
    {
        return new Cell();
    }

    public bool Return(Cell obj)
    {
        if (obj == null) return false;
        
        // Reset cell state pre reuse
        obj.Reset();
        return true;
    }
}

/// <summary>
/// Memory report pre monitoring
/// </summary>
public class MemoryReport
{
    public long TotalMemoryBytes { get; init; }
    public double TotalMemoryMB { get; init; }
    public int Gen0Collections { get; init; }
    public int Gen1Collections { get; init; }
    public int Gen2Collections { get; init; }
    public int WeakCacheAliveReferences { get; init; }
    public int WeakCacheDeadReferences { get; init; }
    public DateTime LastCleanupTime { get; init; }
    public bool IsAggressiveModeEnabled { get; init; }
    public int CleanupIntervalMs { get; init; }
}

/// <summary>
/// Jednoduchý Cell objekt pre ObjectPool
/// </summary>
public class Cell
{
    public object? Value { get; set; }
    public bool IsModified { get; set; }
    public string? FormattedValue { get; set; }
    public Dictionary<string, object>? Metadata { get; set; }

    /// <summary>
    /// Reset cell state pre ObjectPool reuse
    /// </summary>
    public void Reset()
    {
        Value = null;
        IsModified = false;
        FormattedValue = null;
        Metadata?.Clear();
    }
}