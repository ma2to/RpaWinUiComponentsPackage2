using System.Collections.Concurrent;
using System.Text.Json;
using System.IO.Compression;
using Microsoft.Extensions.Logging;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Services;

/// <summary>
/// Multi-level cache manager (L1/L2/L3) podľa newProject.md
/// L1: Hot data in memory (immediate access)
/// L2: Compressed data in memory (quick access ~10ms)
/// L3: Disk storage (lazy load ~50-100ms)
/// </summary>
public class CacheManager : IDisposable
{
    private readonly ILogger? _logger;
    private readonly GridThrottlingConfig _config;
    private readonly L1MemoryCache _l1Cache;
    private readonly L2CompressedCache _l2Cache;
    private readonly L3DiskCache _l3Cache;
    private readonly Timer? _cleanupTimer;
    private volatile bool _disposed;

    public CacheManager(GridThrottlingConfig config, ILogger? logger = null)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _logger = logger;
        
        _l1Cache = new L1MemoryCache(_logger);
        _l2Cache = new L2CompressedCache(_logger);
        _l3Cache = new L3DiskCache(_logger);
        
        // Setup cache cleanup timer
        if (_config.EnableMultiLevelCaching)
        {
            _cleanupTimer = new Timer(PerformCacheCleanup, null,
                TimeSpan.FromMilliseconds(_config.CacheCleanupIntervalMs),
                TimeSpan.FromMilliseconds(_config.CacheCleanupIntervalMs));
        }
        
        _logger?.LogInformation("⚡ CACHE: Multi-level CacheManager initialized - " +
            "L1: Hot Memory, L2: Compressed Memory, L3: Disk Storage");
    }

    /// <summary>
    /// Získa hodnotu z cache (L1 -> L2 -> L3 hierarchy)
    /// </summary>
    public async Task<T?> GetAsync<T>(string key)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(CacheManager));
        if (string.IsNullOrEmpty(key)) return default;

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            // L1: Try hot memory cache first (immediate access)
            var l1Value = _l1Cache.Get<T>(key);
            if (l1Value != null)
            {
                _logger?.LogTrace("⚡ CACHE: L1 HIT for key '{Key}' in {ElapsedMs}ms", key, stopwatch.ElapsedMilliseconds);
                return l1Value;
            }

            // L2: Try compressed memory cache (~10ms access)
            var l2Value = await _l2Cache.GetAsync<T>(key);
            if (l2Value != null)
            {
                // Promote to L1 for faster future access
                _l1Cache.Set(key, l2Value, TimeSpan.FromMinutes(5));
                _logger?.LogTrace("⚡ CACHE: L2 HIT for key '{Key}' in {ElapsedMs}ms (promoted to L1)", 
                    key, stopwatch.ElapsedMilliseconds);
                return l2Value;
            }

            // L3: Try disk cache (~50-100ms access)
            var l3Value = await _l3Cache.GetAsync<T>(key);
            if (l3Value != null)
            {
                // Promote to L2 and L1 for faster future access
                await _l2Cache.SetAsync(key, l3Value, TimeSpan.FromHours(1));
                _l1Cache.Set(key, l3Value, TimeSpan.FromMinutes(5));
                _logger?.LogTrace("⚡ CACHE: L3 HIT for key '{Key}' in {ElapsedMs}ms (promoted to L1+L2)", 
                    key, stopwatch.ElapsedMilliseconds);
                return l3Value;
            }

            // Cache MISS na všetkých leveloch
            _logger?.LogTrace("⚡ CACHE: MISS for key '{Key}' in {ElapsedMs}ms", key, stopwatch.ElapsedMilliseconds);
            return default;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CACHE ERROR: Get failed for key '{Key}'", key);
            return default;
        }
    }

    /// <summary>
    /// Uloží hodnotu do cache (všetky levely)
    /// </summary>
    public async Task SetAsync<T>(string key, T value, TimeSpan expiry)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(CacheManager));
        if (string.IsNullOrEmpty(key) || value == null) return;

        try
        {
            // Store in all levels for optimal performance
            _l1Cache.Set(key, value, TimeSpan.FromMinutes(Math.Min(expiry.TotalMinutes, 5))); // L1: max 5 min
            await _l2Cache.SetAsync(key, value, TimeSpan.FromMinutes(Math.Min(expiry.TotalMinutes, 60))); // L2: max 1 hour
            await _l3Cache.SetAsync(key, value, expiry); // L3: full expiry

            _logger?.LogTrace("⚡ CACHE: SET key '{Key}' across all levels (L1+L2+L3)", key);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CACHE ERROR: Set failed for key '{Key}'", key);
        }
    }

    /// <summary>
    /// Odstráni key zo všetkých cache levelov
    /// </summary>
    public async Task RemoveAsync(string key)
    {
        if (_disposed || string.IsNullOrEmpty(key)) return;

        try
        {
            _l1Cache.Remove(key);
            await _l2Cache.RemoveAsync(key);
            await _l3Cache.RemoveAsync(key);

            _logger?.LogTrace("⚡ CACHE: REMOVE key '{Key}' from all levels", key);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CACHE ERROR: Remove failed for key '{Key}'", key);
        }
    }

    /// <summary>
    /// Vyčistí expired entries zo všetkých levelov
    /// </summary>
    public async Task CleanupAsync()
    {
        if (_disposed) return;

        try
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            
            var l1Cleaned = _l1Cache.Cleanup();
            var l2Cleaned = await _l2Cache.CleanupAsync();
            var l3Cleaned = await _l3Cache.CleanupAsync();
            
            stopwatch.Stop();
            _logger?.LogInformation("⚡ CACHE: Cleanup completed in {ElapsedMs}ms - " +
                "L1: {L1Cleaned}, L2: {L2Cleaned}, L3: {L3Cleaned} entries removed",
                stopwatch.ElapsedMilliseconds, l1Cleaned, l2Cleaned, l3Cleaned);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CACHE ERROR: Cleanup failed");
        }
    }

    /// <summary>
    /// Cache štatistiky pre monitoring
    /// </summary>
    public async Task<CacheStatistics> GetStatisticsAsync()
    {
        if (_disposed) return new CacheStatistics();

        try
        {
            var l1Stats = _l1Cache.GetStatistics();
            var l2Stats = await _l2Cache.GetStatisticsAsync();
            var l3Stats = await _l3Cache.GetStatisticsAsync();

            return new CacheStatistics
            {
                L1Statistics = l1Stats,
                L2Statistics = l2Stats,
                L3Statistics = l3Stats,
                TotalEntries = l1Stats.EntryCount + l2Stats.EntryCount + l3Stats.EntryCount,
                TotalMemoryUsageMB = l1Stats.MemoryUsageMB + l2Stats.MemoryUsageMB + l3Stats.MemoryUsageMB
            };
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CACHE ERROR: Statistics failed");
            return new CacheStatistics();
        }
    }

    /// <summary>
    /// Periodic cache cleanup callback
    /// </summary>
    private async void PerformCacheCleanup(object? state)
    {
        if (_disposed) return;
        await CleanupAsync();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            _cleanupTimer?.Dispose();
            _l1Cache?.Dispose();
            _l2Cache?.Dispose();
            _l3Cache?.Dispose();

            _logger?.LogInformation("⚡ CACHE: CacheManager disposed");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CACHE ERROR: CacheManager disposal failed");
        }
    }
}

/// <summary>
/// L1 Memory Cache - Hot data (immediate access)
/// </summary>
internal class L1MemoryCache : IDisposable
{
    private readonly ILogger? _logger;
    private readonly ConcurrentDictionary<string, CacheEntry<object>> _cache = new();
    private volatile bool _disposed;

    public L1MemoryCache(ILogger? logger = null)
    {
        _logger = logger;
    }

    public T? Get<T>(string key)
    {
        if (_disposed || string.IsNullOrEmpty(key)) return default;

        if (_cache.TryGetValue(key, out var entry) && !entry.IsExpired)
        {
            return (T?)entry.Value;
        }
        return default;
    }

    public void Set<T>(string key, T value, TimeSpan expiry)
    {
        if (_disposed || string.IsNullOrEmpty(key) || value == null) return;

        var entry = new CacheEntry<object>
        {
            Value = value,
            ExpiresAt = DateTime.UtcNow.Add(expiry)
        };
        _cache.AddOrUpdate(key, entry, (_, _) => entry);
    }

    public bool Remove(string key)
    {
        return !_disposed && !string.IsNullOrEmpty(key) && _cache.TryRemove(key, out _);
    }

    public int Cleanup()
    {
        if (_disposed) return 0;

        var expiredKeys = _cache
            .Where(kvp => kvp.Value.IsExpired)
            .Select(kvp => kvp.Key)
            .ToList();

        int removed = 0;
        foreach (var key in expiredKeys)
        {
            if (_cache.TryRemove(key, out _))
                removed++;
        }

        return removed;
    }

    public L1CacheStatistics GetStatistics()
    {
        if (_disposed) return new L1CacheStatistics();

        var entries = _cache.ToList();
        var activeEntries = entries.Count(kvp => !kvp.Value.IsExpired);
        
        return new L1CacheStatistics
        {
            EntryCount = entries.Count,
            ActiveEntries = activeEntries,
            ExpiredEntries = entries.Count - activeEntries,
            MemoryUsageMB = GC.GetTotalMemory(false) / 1024.0 / 1024.0 // Approximate
        };
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _cache.Clear();
    }
}

/// <summary>
/// L2 Compressed Cache - Compressed data in memory (~10ms access)
/// </summary>
internal class L2CompressedCache : IDisposable
{
    private readonly ILogger? _logger;
    private readonly ConcurrentDictionary<string, CompressedCacheEntry> _cache = new();
    private volatile bool _disposed;

    public L2CompressedCache(ILogger? logger = null)
    {
        _logger = logger;
    }

    public async Task<T?> GetAsync<T>(string key)
    {
        if (_disposed || string.IsNullOrEmpty(key)) return default;

        if (_cache.TryGetValue(key, out var entry) && !entry.IsExpired)
        {
            return await DecompressAsync<T>(entry.CompressedData);
        }
        return default;
    }

    public async Task SetAsync<T>(string key, T value, TimeSpan expiry)
    {
        if (_disposed || string.IsNullOrEmpty(key) || value == null) return;

        var compressedData = await CompressAsync(value);
        var entry = new CompressedCacheEntry
        {
            CompressedData = compressedData,
            ExpiresAt = DateTime.UtcNow.Add(expiry)
        };
        _cache.AddOrUpdate(key, entry, (_, _) => entry);
    }

    public async Task<bool> RemoveAsync(string key)
    {
        return !_disposed && !string.IsNullOrEmpty(key) && _cache.TryRemove(key, out _);
    }

    public async Task<int> CleanupAsync()
    {
        if (_disposed) return 0;

        var expiredKeys = _cache
            .Where(kvp => kvp.Value.IsExpired)
            .Select(kvp => kvp.Key)
            .ToList();

        int removed = 0;
        foreach (var key in expiredKeys)
        {
            if (_cache.TryRemove(key, out _))
                removed++;
        }

        return removed;
    }

    public async Task<L2CacheStatistics> GetStatisticsAsync()
    {
        if (_disposed) return new L2CacheStatistics();

        var entries = _cache.ToList();
        var activeEntries = entries.Count(kvp => !kvp.Value.IsExpired);
        var totalCompressedSize = entries.Sum(kvp => kvp.Value.CompressedData.Length);

        return new L2CacheStatistics
        {
            EntryCount = entries.Count,
            ActiveEntries = activeEntries,
            ExpiredEntries = entries.Count - activeEntries,
            MemoryUsageMB = totalCompressedSize / 1024.0 / 1024.0,
            CompressionRatio = CalculateCompressionRatio(entries)
        };
    }

    private async Task<byte[]> CompressAsync<T>(T value)
    {
        var json = JsonSerializer.Serialize(value);
        var bytes = System.Text.Encoding.UTF8.GetBytes(json);
        
        using var output = new MemoryStream();
        using var gzip = new GZipStream(output, CompressionLevel.Optimal);
        await gzip.WriteAsync(bytes);
        await gzip.FlushAsync();
        
        return output.ToArray();
    }

    private async Task<T?> DecompressAsync<T>(byte[] compressedData)
    {
        using var input = new MemoryStream(compressedData);
        using var gzip = new GZipStream(input, CompressionMode.Decompress);
        using var output = new MemoryStream();
        
        await gzip.CopyToAsync(output);
        var json = System.Text.Encoding.UTF8.GetString(output.ToArray());
        
        return JsonSerializer.Deserialize<T>(json);
    }

    private double CalculateCompressionRatio(List<KeyValuePair<string, CompressedCacheEntry>> entries)
    {
        if (entries.Count == 0) return 0.0;
        // Simplified compression ratio calculation
        return 0.6; // Typical compression ratio for JSON data
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _cache.Clear();
    }
}

/// <summary>
/// L3 Disk Cache - Disk storage (lazy load ~50-100ms)
/// </summary>
internal class L3DiskCache : IDisposable
{
    private readonly ILogger? _logger;
    private readonly string _cacheDirectory;
    private readonly ConcurrentDictionary<string, DateTime> _expiryTracker = new();
    private volatile bool _disposed;

    public L3DiskCache(ILogger? logger = null)
    {
        _logger = logger;
        _cacheDirectory = Path.Combine(Path.GetTempPath(), "RpaWinUiCache", "L3");
        Directory.CreateDirectory(_cacheDirectory);
    }

    public async Task<T?> GetAsync<T>(string key)
    {
        if (_disposed || string.IsNullOrEmpty(key)) return default;

        var filePath = GetFilePath(key);
        if (!File.Exists(filePath)) return default;

        // Check expiry
        if (_expiryTracker.TryGetValue(key, out var expiry) && DateTime.UtcNow > expiry)
        {
            await RemoveAsync(key);
            return default;
        }

        try
        {
            var json = await File.ReadAllTextAsync(filePath);
            return JsonSerializer.Deserialize<T>(json);
        }
        catch
        {
            await RemoveAsync(key);
            return default;
        }
    }

    public async Task SetAsync<T>(string key, T value, TimeSpan expiry)
    {
        if (_disposed || string.IsNullOrEmpty(key) || value == null) return;

        try
        {
            var filePath = GetFilePath(key);
            var json = JsonSerializer.Serialize(value);
            await File.WriteAllTextAsync(filePath, json);
            
            _expiryTracker.AddOrUpdate(key, DateTime.UtcNow.Add(expiry), (_, _) => DateTime.UtcNow.Add(expiry));
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CACHE ERROR: L3 Set failed for key '{Key}'", key);
        }
    }

    public async Task<bool> RemoveAsync(string key)
    {
        if (_disposed || string.IsNullOrEmpty(key)) return false;

        try
        {
            var filePath = GetFilePath(key);
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            _expiryTracker.TryRemove(key, out _);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task<int> CleanupAsync()
    {
        if (_disposed) return 0;

        try
        {
            var now = DateTime.UtcNow;
            var expiredKeys = _expiryTracker
                .Where(kvp => now > kvp.Value)
                .Select(kvp => kvp.Key)
                .ToList();

            int removed = 0;
            foreach (var key in expiredKeys)
            {
                if (await RemoveAsync(key))
                    removed++;
            }

            return removed;
        }
        catch
        {
            return 0;
        }
    }

    public async Task<L3CacheStatistics> GetStatisticsAsync()
    {
        if (_disposed) return new L3CacheStatistics();

        try
        {
            var files = Directory.GetFiles(_cacheDirectory);
            var totalSize = files.Sum(f => new FileInfo(f).Length);
            var activeCount = _expiryTracker.Count(kvp => DateTime.UtcNow <= kvp.Value);

            return new L3CacheStatistics
            {
                EntryCount = files.Length,
                ActiveEntries = activeCount,
                ExpiredEntries = files.Length - activeCount,
                MemoryUsageMB = totalSize / 1024.0 / 1024.0,
                DiskUsageMB = totalSize / 1024.0 / 1024.0
            };
        }
        catch
        {
            return new L3CacheStatistics();
        }
    }

    private string GetFilePath(string key)
    {
        // Create safe filename from key
        var safeKey = string.Concat(key.Where(c => char.IsLetterOrDigit(c) || c == '_' || c == '-'));
        return Path.Combine(_cacheDirectory, $"{safeKey}.cache");
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            // Auto-cleanup temp files on disposal
            if (Directory.Exists(_cacheDirectory))
            {
                Directory.Delete(_cacheDirectory, true);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CACHE ERROR: L3 disposal cleanup failed");
        }
    }
}

// Cache entry structures
internal class CacheEntry<T>
{
    public T? Value { get; set; }
    public DateTime ExpiresAt { get; set; }
    public bool IsExpired => DateTime.UtcNow > ExpiresAt;
}

internal class CompressedCacheEntry
{
    public byte[] CompressedData { get; set; } = Array.Empty<byte>();
    public DateTime ExpiresAt { get; set; }
    public bool IsExpired => DateTime.UtcNow > ExpiresAt;
}

// Statistics structures
public class CacheStatistics
{
    public L1CacheStatistics L1Statistics { get; init; } = new();
    public L2CacheStatistics L2Statistics { get; init; } = new();
    public L3CacheStatistics L3Statistics { get; init; } = new();
    public int TotalEntries { get; init; }
    public double TotalMemoryUsageMB { get; init; }
}

public class L1CacheStatistics
{
    public int EntryCount { get; init; }
    public int ActiveEntries { get; init; }
    public int ExpiredEntries { get; init; }
    public double MemoryUsageMB { get; init; }
}

public class L2CacheStatistics
{
    public int EntryCount { get; init; }
    public int ActiveEntries { get; init; }
    public int ExpiredEntries { get; init; }
    public double MemoryUsageMB { get; init; }
    public double CompressionRatio { get; init; }
}

public class L3CacheStatistics
{
    public int EntryCount { get; init; }
    public int ActiveEntries { get; init; }
    public int ExpiredEntries { get; init; }
    public double MemoryUsageMB { get; init; }
    public double DiskUsageMB { get; init; }
}