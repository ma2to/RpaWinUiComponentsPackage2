using Microsoft.Extensions.Logging;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Services;

/// <summary>
/// Main Performance Module integrating all performance services
/// Implementuje kompletný Performance Module z newProject.md dokumentácie
/// </summary>
public class PerformanceModule : IDisposable
{
    private readonly ILogger? _logger;
    private readonly GridThrottlingConfig _config;
    private readonly Lazy<MemoryManager> _memoryManager;
    private readonly Lazy<CacheManager> _cacheManager;
    private readonly Lazy<LargeFileOptimizer> _largeFileOptimizer;
    private readonly Lazy<BackgroundProcessor> _backgroundProcessor;
    private volatile bool _disposed;

    public PerformanceModule(GridThrottlingConfig? config = null, ILogger? logger = null)
    {
        _config = config ?? GridThrottlingConfig.Default;
        _logger = logger;

        // Initialize services lazily for optimal startup performance
        _memoryManager = new Lazy<MemoryManager>(() => new MemoryManager(_config, _logger));
        _cacheManager = new Lazy<CacheManager>(() => new CacheManager(_config, _logger));
        _largeFileOptimizer = new Lazy<LargeFileOptimizer>(() => new LargeFileOptimizer(_config, _logger));
        _backgroundProcessor = new Lazy<BackgroundProcessor>(() => new BackgroundProcessor(_config, _logger));

        _logger?.LogInformation("⚡ PERFORMANCE: PerformanceModule initialized with config: {ConfigType}", 
            GetConfigurationName());
    }

    /// <summary>
    /// Memory management service s ObjectPool a aggressive GC
    /// </summary>
    public MemoryManager Memory => _disposed ? throw new ObjectDisposedException(nameof(PerformanceModule)) : _memoryManager.Value;

    /// <summary>
    /// Multi-level cache manager (L1/L2/L3)
    /// </summary>
    public CacheManager Cache => _disposed ? throw new ObjectDisposedException(nameof(PerformanceModule)) : _cacheManager.Value;

    /// <summary>
    /// Large file optimization s streaming operations
    /// </summary>
    public LargeFileOptimizer LargeFile => _disposed ? throw new ObjectDisposedException(nameof(PerformanceModule)) : _largeFileOptimizer.Value;

    /// <summary>
    /// Background processing s cancellation tokens
    /// </summary>
    public BackgroundProcessor Background => _disposed ? throw new ObjectDisposedException(nameof(PerformanceModule)) : _backgroundProcessor.Value;

    /// <summary>
    /// Aktuálna performance konfigurácia
    /// </summary>
    public GridThrottlingConfig Configuration => _config;

    /// <summary>
    /// Získa kompletný performance report zo všetkých subsystémov
    /// </summary>
    public async Task<PerformanceReport> GetPerformanceReportAsync()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(PerformanceModule));

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            // Gather statistics from all subsystems
            var memoryReport = _memoryManager.IsValueCreated ? await _memoryManager.Value.GetMemoryUsageAsync() : null;
            var cacheStats = _cacheManager.IsValueCreated ? await _cacheManager.Value.GetStatisticsAsync() : null;
            var backgroundStats = _backgroundProcessor.IsValueCreated ? _backgroundProcessor.Value.GetStatistics() : null;

            stopwatch.Stop();

            var report = new PerformanceReport
            {
                GeneratedAt = DateTime.UtcNow,
                ConfigurationName = GetConfigurationName(),
                MemoryReport = memoryReport,
                CacheStatistics = cacheStats,
                BackgroundProcessorStatistics = backgroundStats,
                ReportGenerationTimeMs = stopwatch.ElapsedMilliseconds,
                
                // Module initialization status
                IsMemoryManagerInitialized = _memoryManager.IsValueCreated,
                IsCacheManagerInitialized = _cacheManager.IsValueCreated,
                IsLargeFileOptimizerInitialized = _largeFileOptimizer.IsValueCreated,
                IsBackgroundProcessorInitialized = _backgroundProcessor.IsValueCreated
            };

            _logger?.LogInformation("⚡ PERFORMANCE: Performance report generated in {ElapsedMs}ms - " +
                "Memory: {MemoryMB}MB, Cache entries: {CacheEntries}, Background tasks: {BackgroundTasks}",
                stopwatch.ElapsedMilliseconds,
                memoryReport?.TotalMemoryMB ?? 0,
                cacheStats?.TotalEntries ?? 0,
                backgroundStats?.QueuedTasks ?? 0);

            return report;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 PERFORMANCE ERROR: Failed to generate performance report");
            throw;
        }
    }

    /// <summary>
    /// Optimalizuje všetky performance subsystémy naraz
    /// </summary>
    public async Task OptimizeAllAsync()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(PerformanceModule));

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        
        try
        {
            _logger?.LogInformation("⚡ PERFORMANCE: Starting comprehensive optimization...");

            var optimizationTasks = new List<Task>();

            // Memory optimization
            if (_memoryManager.IsValueCreated)
            {
                optimizationTasks.Add(_memoryManager.Value.OptimizeMemoryAsync());
            }

            // Cache cleanup
            if (_cacheManager.IsValueCreated)
            {
                optimizationTasks.Add(_cacheManager.Value.CleanupAsync());
            }

            // Wait for completion of all background tasks
            if (_backgroundProcessor.IsValueCreated)
            {
                optimizationTasks.Add(_backgroundProcessor.Value.WaitForCompletionAsync(TimeSpan.FromMinutes(2)));
            }

            // Execute all optimizations in parallel
            await Task.WhenAll(optimizationTasks);

            stopwatch.Stop();
            
            _logger?.LogInformation("⚡ PERFORMANCE: Comprehensive optimization completed in {ElapsedMs}ms",
                stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 PERFORMANCE ERROR: Comprehensive optimization failed");
            throw;
        }
    }

    /// <summary>
    /// Warm-up všetkých performance subsystémov pre optimal startup
    /// </summary>
    public async Task WarmUpAsync()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(PerformanceModule));

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            _logger?.LogInformation("⚡ PERFORMANCE: Starting warm-up sequence...");

            // Initialize all services (triggers lazy loading)
            _ = Memory;
            _ = Cache;
            _ = LargeFile;
            _ = Background;

            // Pre-populate cache with some test data to warm up L1/L2/L3
            await Cache.SetAsync("warmup-test", "initialized", TimeSpan.FromMinutes(1));
            _ = await Cache.GetAsync<string>("warmup-test");

            // Trigger initial memory report
            _ = await Memory.GetMemoryUsageAsync();

            stopwatch.Stop();
            
            _logger?.LogInformation("⚡ PERFORMANCE: Warm-up completed in {ElapsedMs}ms - All subsystems ready",
                stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 PERFORMANCE ERROR: Warm-up failed");
            throw;
        }
    }

    /// <summary>
    /// Factory method pre rôzne performance scenáre
    /// </summary>
    public static PerformanceModule CreateForScenario(PerformanceScenario scenario, ILogger? logger = null)
    {
        var config = scenario switch
        {
            PerformanceScenario.HighPerformance => GridThrottlingConfig.HighPerformance,
            PerformanceScenario.BatterySaver => GridThrottlingConfig.BatterySaver,
            PerformanceScenario.LargeDataset => GridThrottlingConfig.LargeDataset,
            _ => GridThrottlingConfig.Default
        };

        logger?.LogInformation("⚡ PERFORMANCE: Creating PerformanceModule for scenario: {Scenario}", scenario);
        return new PerformanceModule(config, logger);
    }

    /// <summary>
    /// Určí názov aktuálnej konfigurácie
    /// </summary>
    private string GetConfigurationName()
    {
        if (_config == GridThrottlingConfig.HighPerformance) return "HighPerformance";
        if (_config == GridThrottlingConfig.BatterySaver) return "BatterySaver";  
        if (_config == GridThrottlingConfig.LargeDataset) return "LargeDataset";
        return "Default";
    }

    /// <summary>
    /// Diagnostické informácie pre debugging
    /// </summary>
    public async Task<string> GetDiagnosticInfoAsync()
    {
        if (_disposed) return "PerformanceModule: DISPOSED";

        try
        {
            var report = await GetPerformanceReportAsync();
            
            return $"""
                ⚡ PERFORMANCE MODULE DIAGNOSTICS ⚡
                Generated: {report.GeneratedAt:yyyy-MM-dd HH:mm:ss} UTC
                Configuration: {report.ConfigurationName}
                Report Generation Time: {report.ReportGenerationTimeMs}ms
                
                📊 MEMORY:
                - Manager Initialized: {report.IsMemoryManagerInitialized}
                - Total Memory: {report.MemoryReport?.TotalMemoryMB:F2} MB
                - GC Gen0/1/2: {report.MemoryReport?.Gen0Collections}/{report.MemoryReport?.Gen1Collections}/{report.MemoryReport?.Gen2Collections}
                - Weak Cache Alive/Dead: {report.MemoryReport?.WeakCacheAliveReferences}/{report.MemoryReport?.WeakCacheDeadReferences}
                - Aggressive Mode: {report.MemoryReport?.IsAggressiveModeEnabled}
                
                📈 CACHE:
                - Manager Initialized: {report.IsCacheManagerInitialized}
                - Total Entries: {report.CacheStatistics?.TotalEntries}
                - L1 Active: {report.CacheStatistics?.L1Statistics.ActiveEntries}
                - L2 Active: {report.CacheStatistics?.L2Statistics.ActiveEntries} 
                - L3 Active: {report.CacheStatistics?.L3Statistics.ActiveEntries}
                - Total Memory Usage: {report.CacheStatistics?.TotalMemoryUsageMB:F2} MB
                
                🔄 BACKGROUND:
                - Processor Initialized: {report.IsBackgroundProcessorInitialized}
                - Queued Tasks: {report.BackgroundProcessorStatistics?.QueuedTasks}
                - Processed: {report.BackgroundProcessorStatistics?.TotalTasksProcessed}
                - Failed: {report.BackgroundProcessorStatistics?.TotalTasksFailed}
                - Success Rate: {report.BackgroundProcessorStatistics?.SuccessRate:P2}
                - Max Concurrent: {report.BackgroundProcessorStatistics?.MaxConcurrentTasks}
                
                📁 LARGE FILE:
                - Optimizer Initialized: {report.IsLargeFileOptimizerInitialized}
                
                ⚙️ CONFIGURATION:
                - UI Update Interval: {_config.UIUpdateIntervalMs}ms
                - Validation Update Interval: {_config.ValidationUpdateIntervalMs}ms  
                - Bulk Operation Batch Size: {_config.BulkOperationBatchSize}
                - Virtualization Buffer: {_config.VirtualizationBufferSize}
                - Memory Cleanup Interval: {_config.MemoryCleanupIntervalMs}ms
                - Cache Cleanup Interval: {_config.CacheCleanupIntervalMs}ms
                - Multi-level Caching: {_config.EnableMultiLevelCaching}
                - Background Processing: {_config.EnableBackgroundProcessing}
                - Aggressive Memory Management: {_config.EnableAggressiveMemoryManagement}
                """;
        }
        catch (Exception ex)
        {
            return $"⚡ PERFORMANCE MODULE DIAGNOSTICS ⚡\nERROR: {ex.Message}";
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            _logger?.LogInformation("⚡ PERFORMANCE: Disposing PerformanceModule...");

            // Dispose services in reverse order of dependencies
            if (_backgroundProcessor.IsValueCreated)
                _backgroundProcessor.Value.Dispose();
                
            if (_largeFileOptimizer.IsValueCreated)
                _largeFileOptimizer.Value.Dispose();
                
            if (_cacheManager.IsValueCreated)
                _cacheManager.Value.Dispose();
                
            if (_memoryManager.IsValueCreated)
                _memoryManager.Value.Dispose();

            _logger?.LogInformation("⚡ PERFORMANCE: PerformanceModule disposed successfully");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 PERFORMANCE ERROR: PerformanceModule disposal failed");
        }
    }
}

/// <summary>
/// Performance scenáre pre factory method
/// </summary>
public enum PerformanceScenario
{
    Default,
    HighPerformance,
    BatterySaver,
    LargeDataset
}

/// <summary>
/// Kompletný performance report zo všetkých subsystémov
/// </summary>
public class PerformanceReport
{
    public DateTime GeneratedAt { get; init; }
    public string ConfigurationName { get; init; } = string.Empty;
    public long ReportGenerationTimeMs { get; init; }
    
    // Subsystem reports
    public MemoryReport? MemoryReport { get; init; }
    public CacheStatistics? CacheStatistics { get; init; }
    public BackgroundProcessorStatistics? BackgroundProcessorStatistics { get; init; }
    
    // Initialization status
    public bool IsMemoryManagerInitialized { get; init; }
    public bool IsCacheManagerInitialized { get; init; }
    public bool IsLargeFileOptimizerInitialized { get; init; }
    public bool IsBackgroundProcessorInitialized { get; init; }
    
    // Overall health
    public bool IsHealthy => (MemoryReport?.TotalMemoryMB ?? 0) < 1000 && // < 1GB memory usage
                            (BackgroundProcessorStatistics?.SuccessRate ?? 1.0) > 0.95; // > 95% success rate
}