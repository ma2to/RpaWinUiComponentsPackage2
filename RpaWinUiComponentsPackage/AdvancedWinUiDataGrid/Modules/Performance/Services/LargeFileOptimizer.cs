using System.IO.Compression;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Services;

/// <summary>
/// Large file optimization s streaming, progressive loading a virtualization
/// Implementuje Large File Optimization stratégiu z newProject.md
/// </summary>
public class LargeFileOptimizer : IDisposable
{
    private readonly ILogger? _logger;
    private readonly GridThrottlingConfig _config;
    private readonly SemaphoreSlim _streamingSemaphore;
    private volatile bool _disposed;

    public LargeFileOptimizer(GridThrottlingConfig config, ILogger? logger = null)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _logger = logger;
        
        // Limit concurrent streaming operations
        _streamingSemaphore = new SemaphoreSlim(2, 2);
        
        _logger?.LogInformation("⚡ OPTIMIZER: LargeFileOptimizer initialized with batch size {BatchSize}",
            _config.BulkOperationBatchSize);
    }

    /// <summary>
    /// Streaming import pre large datasets s progress reportingom
    /// </summary>
    public async Task<ImportResult> StreamingImportAsync(
        Stream dataStream, 
        IProgress<ImportProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(LargeFileOptimizer));
        if (dataStream == null) throw new ArgumentNullException(nameof(dataStream));

        await _streamingSemaphore.WaitAsync(cancellationToken);
        
        try
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var result = new ImportResult();
            
            _logger?.LogInformation("⚡ OPTIMIZER: Starting streaming import - Stream length: {StreamLength} bytes",
                dataStream.CanSeek ? dataStream.Length : -1);

            using var reader = new StreamReader(dataStream);
            var batch = new List<Dictionary<string, object?>>();
            var totalProcessed = 0;
            var batchNumber = 0;
            
            string? line;
            while ((line = await reader.ReadLineAsync()) != null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                try
                {
                    // Parse line (assuming JSON format for this example)
                    var rowData = JsonSerializer.Deserialize<Dictionary<string, object?>>(line);
                    if (rowData != null)
                    {
                        batch.Add(rowData);
                        totalProcessed++;
                    }
                }
                catch (JsonException ex)
                {
                    result.ErrorCount++;
                    _logger?.LogWarning("⚠️ OPTIMIZER: Failed to parse line {LineNumber}: {Error}", 
                        totalProcessed + 1, ex.Message);
                    continue;
                }

                // Process batch when full
                if (batch.Count >= _config.BulkOperationBatchSize)
                {
                    await ProcessBatch(batch, batchNumber++, result, cancellationToken);
                    batch.Clear();
                    
                    // Report progress
                    progress?.Report(new ImportProgress
                    {
                        ProcessedRows = totalProcessed,
                        CurrentBatch = batchNumber,
                        ElapsedTime = stopwatch.Elapsed,
                        EstimatedTimeRemaining = EstimateRemainingTime(stopwatch, totalProcessed, dataStream)
                    });
                    
                    // Yield control to avoid UI blocking
                    await Task.Yield();
                }
            }

            // Process final batch if any remaining
            if (batch.Count > 0)
            {
                await ProcessBatch(batch, batchNumber++, result, cancellationToken);
            }

            stopwatch.Stop();
            result.TotalProcessed = totalProcessed;
            result.TotalBatches = batchNumber;
            result.ElapsedTime = stopwatch.Elapsed;
            result.ThroughputRowsPerSecond = stopwatch.Elapsed.TotalSeconds > 0 
                ? totalProcessed / stopwatch.Elapsed.TotalSeconds 
                : 0;

            _logger?.LogInformation("⚡ OPTIMIZER: Streaming import completed - " +
                "Processed: {ProcessedRows} rows, Batches: {Batches}, " +
                "Elapsed: {ElapsedMs}ms, Throughput: {Throughput:F2} rows/sec, " +
                "Errors: {Errors}",
                totalProcessed, batchNumber, stopwatch.ElapsedMilliseconds, 
                result.ThroughputRowsPerSecond, result.ErrorCount);

            return result;
        }
        catch (OperationCanceledException)
        {
            _logger?.LogWarning("⚠️ OPTIMIZER: Streaming import cancelled");
            throw;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 OPTIMIZER ERROR: Streaming import failed");
            throw;
        }
        finally
        {
            _streamingSemaphore.Release();
        }
    }

    /// <summary>
    /// Streaming export pre large datasets
    /// </summary>
    public async Task<Stream> StreamingExportAsync(
        IAsyncEnumerable<Dictionary<string, object?>> data,
        ExportFormat format,
        IProgress<ExportProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(LargeFileOptimizer));
        if (data == null) throw new ArgumentNullException(nameof(data));

        await _streamingSemaphore.WaitAsync(cancellationToken);

        try
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var outputStream = new MemoryStream();
            
            _logger?.LogInformation("⚡ OPTIMIZER: Starting streaming export - Format: {Format}", format);

            using var writer = new StreamWriter(outputStream, leaveOpen: true);
            var processedRows = 0;
            var batchCount = 0;

            await foreach (var row in data.WithCancellation(cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();

                // Format row according to export format
                var formattedRow = FormatRowForExport(row, format);
                await writer.WriteLineAsync(formattedRow);
                
                processedRows++;

                // Progress reporting every batch
                if (processedRows % _config.BulkOperationBatchSize == 0)
                {
                    batchCount++;
                    await writer.FlushAsync();
                    
                    progress?.Report(new ExportProgress
                    {
                        ProcessedRows = processedRows,
                        CurrentBatch = batchCount,
                        ElapsedTime = stopwatch.Elapsed,
                        ThroughputRowsPerSecond = stopwatch.Elapsed.TotalSeconds > 0 
                            ? processedRows / stopwatch.Elapsed.TotalSeconds 
                            : 0
                    });

                    // Yield control
                    await Task.Yield();
                }
            }

            await writer.FlushAsync();
            outputStream.Seek(0, SeekOrigin.Begin);

            stopwatch.Stop();
            _logger?.LogInformation("⚡ OPTIMIZER: Streaming export completed - " +
                "Processed: {ProcessedRows} rows, Batches: {Batches}, " +
                "Elapsed: {ElapsedMs}ms, Size: {SizeMB:F2}MB",
                processedRows, batchCount, stopwatch.ElapsedMilliseconds,
                outputStream.Length / 1024.0 / 1024.0);

            return outputStream;
        }
        catch (OperationCanceledException)
        {
            _logger?.LogWarning("⚠️ OPTIMIZER: Streaming export cancelled");
            throw;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 OPTIMIZER ERROR: Streaming export failed");
            throw;
        }
        finally
        {
            _streamingSemaphore.Release();
        }
    }

    /// <summary>
    /// Progressive loading pre large datasets s chunking
    /// </summary>
    public async IAsyncEnumerable<DataChunk> LoadProgressivelyAsync(
        Stream dataStream,
        int chunkSize,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(LargeFileOptimizer));
        if (dataStream == null) throw new ArgumentNullException(nameof(dataStream));

        _logger?.LogInformation("⚡ OPTIMIZER: Starting progressive loading - Chunk size: {ChunkSize}", chunkSize);

        using var reader = new StreamReader(dataStream);
        var chunkNumber = 0;
        var totalProcessed = 0;

        while (!reader.EndOfStream)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var chunk = new DataChunk
            {
                ChunkNumber = chunkNumber++,
                Data = new List<Dictionary<string, object?>>()
            };

            // Load chunk
            for (int i = 0; i < chunkSize && !reader.EndOfStream; i++)
            {
                var line = await reader.ReadLineAsync();
                if (line == null) break;

                try
                {
                    var rowData = JsonSerializer.Deserialize<Dictionary<string, object?>>(line);
                    if (rowData != null)
                    {
                        chunk.Data.Add(rowData);
                        totalProcessed++;
                    }
                }
                catch (JsonException ex)
                {
                    _logger?.LogWarning("⚠️ OPTIMIZER: Failed to parse line in chunk {ChunkNumber}: {Error}", 
                        chunkNumber, ex.Message);
                }
            }

            chunk.TotalRowsProcessed = totalProcessed;
            
            if (chunk.Data.Count > 0)
            {
                _logger?.LogTrace("⚡ OPTIMIZER: Yielding chunk {ChunkNumber} with {RowCount} rows", 
                    chunkNumber, chunk.Data.Count);
                yield return chunk;
            }

            // Yield control between chunks
            await Task.Yield();
        }

        _logger?.LogInformation("⚡ OPTIMIZER: Progressive loading completed - " +
            "Total chunks: {ChunkCount}, Total rows: {TotalRows}", chunkNumber, totalProcessed);
    }

    /// <summary>
    /// Memory-efficient file compression pre large exports
    /// </summary>
    public async Task<Stream> CompressStreamAsync(Stream inputStream, CancellationToken cancellationToken = default)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(LargeFileOptimizer));
        if (inputStream == null) throw new ArgumentNullException(nameof(inputStream));

        try
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var outputStream = new MemoryStream();
            
            using (var gzip = new GZipStream(outputStream, CompressionLevel.Optimal, leaveOpen: true))
            {
                await inputStream.CopyToAsync(gzip, 81920, cancellationToken); // 80KB buffer
                await gzip.FlushAsync(cancellationToken);
            }

            var originalSize = inputStream.Length;
            var compressedSize = outputStream.Length;
            var compressionRatio = originalSize > 0 ? (double)compressedSize / originalSize : 0;

            stopwatch.Stop();
            outputStream.Seek(0, SeekOrigin.Begin);

            _logger?.LogInformation("⚡ OPTIMIZER: Stream compression completed - " +
                "Original: {OriginalMB:F2}MB, Compressed: {CompressedMB:F2}MB, " +
                "Ratio: {Ratio:P2}, Elapsed: {ElapsedMs}ms",
                originalSize / 1024.0 / 1024.0, compressedSize / 1024.0 / 1024.0,
                compressionRatio, stopwatch.ElapsedMilliseconds);

            return outputStream;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 OPTIMIZER ERROR: Stream compression failed");
            throw;
        }
    }

    /// <summary>
    /// Process batch of data
    /// </summary>
    private async Task ProcessBatch(
        List<Dictionary<string, object?>> batch, 
        int batchNumber, 
        ImportResult result,
        CancellationToken cancellationToken)
    {
        try
        {
            _logger?.LogTrace("⚡ OPTIMIZER: Processing batch {BatchNumber} with {RowCount} rows", 
                batchNumber, batch.Count);

            // Simulate batch processing (replace with actual implementation)
            await Task.Delay(10, cancellationToken); // Simulate processing time
            
            result.SuccessCount += batch.Count;
        }
        catch (Exception ex)
        {
            result.ErrorCount += batch.Count;
            _logger?.LogError(ex, "🚨 OPTIMIZER ERROR: Batch {BatchNumber} processing failed", batchNumber);
        }
    }

    /// <summary>
    /// Format row for specific export format
    /// </summary>
    private string FormatRowForExport(Dictionary<string, object?> row, ExportFormat format)
    {
        return format switch
        {
            ExportFormat.Json => JsonSerializer.Serialize(row),
            ExportFormat.Csv => string.Join(",", row.Values.Select(v => $"\"{v}\"")),
            ExportFormat.Xml => ConvertToXml(row),
            _ => JsonSerializer.Serialize(row)
        };
    }

    /// <summary>
    /// Simple XML conversion
    /// </summary>
    private string ConvertToXml(Dictionary<string, object?> row)
    {
        var xml = "<Row>";
        foreach (var kvp in row)
        {
            xml += $"<{kvp.Key}>{kvp.Value}</{kvp.Key}>";
        }
        xml += "</Row>";
        return xml;
    }

    /// <summary>
    /// Estimate remaining time based on current progress
    /// </summary>
    private TimeSpan EstimateRemainingTime(System.Diagnostics.Stopwatch stopwatch, int processed, Stream stream)
    {
        if (!stream.CanSeek || processed == 0) return TimeSpan.Zero;

        var totalBytes = stream.Length;
        var processedBytes = stream.Position;
        var remainingBytes = totalBytes - processedBytes;
        
        if (remainingBytes <= 0) return TimeSpan.Zero;

        var elapsedSeconds = stopwatch.Elapsed.TotalSeconds;
        var bytesPerSecond = processedBytes / elapsedSeconds;
        var remainingSeconds = remainingBytes / bytesPerSecond;

        return TimeSpan.FromSeconds(remainingSeconds);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            _streamingSemaphore?.Dispose();
            _logger?.LogInformation("⚡ OPTIMIZER: LargeFileOptimizer disposed");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 OPTIMIZER ERROR: LargeFileOptimizer disposal failed");
        }
    }
}

/// <summary>
/// Import result statistics
/// </summary>
public class ImportResult
{
    public int TotalProcessed { get; set; }
    public int SuccessCount { get; set; }
    public int ErrorCount { get; set; }
    public int TotalBatches { get; set; }
    public TimeSpan ElapsedTime { get; set; }
    public double ThroughputRowsPerSecond { get; set; }
    public bool IsSuccess => ErrorCount == 0;
}

/// <summary>
/// Import progress reporting
/// </summary>
public class ImportProgress
{
    public int ProcessedRows { get; init; }
    public int CurrentBatch { get; init; }
    public TimeSpan ElapsedTime { get; init; }
    public TimeSpan EstimatedTimeRemaining { get; init; }
}

/// <summary>
/// Export progress reporting
/// </summary>
public class ExportProgress
{
    public int ProcessedRows { get; init; }
    public int CurrentBatch { get; init; }
    public TimeSpan ElapsedTime { get; init; }
    public double ThroughputRowsPerSecond { get; init; }
}

/// <summary>
/// Data chunk pre progressive loading
/// </summary>
public class DataChunk
{
    public int ChunkNumber { get; init; }
    public List<Dictionary<string, object?>> Data { get; init; } = new();
    public int TotalRowsProcessed { get; set; }
}

/// <summary>
/// Export formats
/// </summary>
public enum ExportFormat
{
    Json,
    Csv,
    Xml
}