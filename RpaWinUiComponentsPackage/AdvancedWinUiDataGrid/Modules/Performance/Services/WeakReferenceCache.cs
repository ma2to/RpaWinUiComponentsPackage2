using System.Collections.Concurrent;
using Microsoft.Extensions.Logging;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Services;

/// <summary>
/// WeakReference cache pre optimalizáciu memory usage
/// Implementuje multi-level memory management stratégiu z newProject.md
/// </summary>
public class WeakReferenceCache : IDisposable
{
    private readonly ConcurrentDictionary<string, WeakReference> _cache = new();
    private readonly object _lockObject = new();
    private volatile bool _disposed;

    /// <summary>
    /// Uloží objekt do weak reference cache
    /// </summary>
    public void Set<T>(string key, T value) where T : class
    {
        try
        {
            if (_disposed) throw new ObjectDisposedException(nameof(WeakReferenceCache));
            if (string.IsNullOrEmpty(key)) throw new ArgumentException("Key cannot be null or empty", nameof(key));
            if (value == null) throw new ArgumentNullException(nameof(value));

            _cache.AddOrUpdate(key, new WeakReference(value), (_, _) => new WeakReference(value));
        }
        catch (Exception ex)
        {
            // Log the error but don't rethrow to prevent cache errors from breaking application flow
            // This is a performance optimization feature, not critical functionality
            System.Diagnostics.Debug.WriteLine($"🚨 CACHE ERROR: Failed to set cache value - Key: {key}, Type: {typeof(T).Name}, Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Získa objekt z weak reference cache
    /// </summary>
    public T? Get<T>(string key) where T : class
    {
        try
        {
            if (_disposed) throw new ObjectDisposedException(nameof(WeakReferenceCache));
            if (string.IsNullOrEmpty(key)) return null;

            if (_cache.TryGetValue(key, out var weakRef) && weakRef.IsAlive)
            {
                return weakRef.Target as T;
            }

            return null;
        }
        catch (Exception ex)
        {
            // Log the error but return null to prevent cache errors from breaking application flow
            System.Diagnostics.Debug.WriteLine($"🚨 CACHE ERROR: Failed to get cache value - Key: {key}, Type: {typeof(T).Name}, Error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Skontroluje či existuje a je alive
    /// </summary>
    public bool ContainsAlive(string key)
    {
        if (_disposed || string.IsNullOrEmpty(key)) return false;
        
        return _cache.TryGetValue(key, out var weakRef) && weakRef.IsAlive;
    }

    /// <summary>
    /// Odstráni key z cache
    /// </summary>
    public bool Remove(string key)
    {
        if (_disposed || string.IsNullOrEmpty(key)) return false;
        
        return _cache.TryRemove(key, out _);
    }

    /// <summary>
    /// Vyčistí všetky dead references
    /// </summary>
    public int Cleanup()
    {
        if (_disposed) return 0;

        lock (_lockObject)
        {
            var deadKeys = new List<string>();
            
            foreach (var kvp in _cache)
            {
                if (!kvp.Value.IsAlive)
                {
                    deadKeys.Add(kvp.Key);
                }
            }

            foreach (var deadKey in deadKeys)
            {
                _cache.TryRemove(deadKey, out _);
            }

            return deadKeys.Count;
        }
    }

    /// <summary>
    /// Počet živých references
    /// </summary>
    public int GetAliveCount()
    {
        if (_disposed) return 0;

        int aliveCount = 0;
        foreach (var weakRef in _cache.Values)
        {
            if (weakRef.IsAlive) aliveCount++;
        }
        return aliveCount;
    }

    /// <summary>
    /// Počet dead references
    /// </summary>
    public int GetDeadCount()
    {
        if (_disposed) return 0;

        int deadCount = 0;
        foreach (var weakRef in _cache.Values)
        {
            if (!weakRef.IsAlive) deadCount++;
        }
        return deadCount;
    }

    /// <summary>
    /// Celkový počet entries
    /// </summary>
    public int GetTotalCount()
    {
        return _disposed ? 0 : _cache.Count;
    }

    /// <summary>
    /// Vyčistí celý cache
    /// </summary>
    public void Clear()
    {
        if (_disposed) return;
        
        _cache.Clear();
    }

    /// <summary>
    /// Získa všetky keys pre debugging
    /// </summary>
    public IEnumerable<string> GetKeys()
    {
        if (_disposed) return Enumerable.Empty<string>();
        
        return _cache.Keys.ToList();
    }

    /// <summary>
    /// Cache statistics pre monitoring
    /// </summary>
    public WeakCacheStatistics GetStatistics()
    {
        if (_disposed) 
        {
            return new WeakCacheStatistics
            {
                TotalEntries = 0,
                AliveReferences = 0,
                DeadReferences = 0,
                HitRatio = 0.0
            };
        }

        var alive = GetAliveCount();
        var total = GetTotalCount();
        
        return new WeakCacheStatistics
        {
            TotalEntries = total,
            AliveReferences = alive,
            DeadReferences = total - alive,
            HitRatio = total > 0 ? (double)alive / total : 0.0
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
/// Štatistiky weak cache
/// </summary>
public class WeakCacheStatistics
{
    public int TotalEntries { get; init; }
    public int AliveReferences { get; init; }
    public int DeadReferences { get; init; }
    public double HitRatio { get; init; }
}