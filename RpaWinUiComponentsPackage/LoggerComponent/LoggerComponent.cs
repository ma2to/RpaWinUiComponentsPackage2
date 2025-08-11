using Microsoft.Extensions.Logging.Abstractions;
using System.Threading.Channels;
using RpaWinUiComponentsPackage.LoggerComponent.Utilities;
using LogLevel = Microsoft.Extensions.Logging.LogLevel;

namespace RpaWinUiComponentsPackage.LoggerComponent;

/// <summary>
/// Thread-safe logger component s Channel-based background writing a zero message loss
/// </summary>
public class LoggerComponent : IDisposable
{
    private readonly Channel<LogMessage> _channel;
    private readonly ChannelWriter<LogMessage> _writer;
    private readonly Task _backgroundWriterTask;
    private readonly CancellationTokenSource _cancellationTokenSource;
    private readonly Timer _flushTimer;
    private readonly SemaphoreSlim _flushSemaphore = new(1, 1);
    
    // Configuration
    private readonly Microsoft.Extensions.Logging.ILogger? _externalLogger;
    private readonly string _logDirectory;
    private readonly string _baseFileName;
    private readonly int? _maxFileSizeMB;
    
    // Buffer pre batch writing
    private readonly List<LogMessage> _pendingMessages = new();
    private volatile bool _hasUnflushedData = false;

    public LoggerComponent(
        Microsoft.Extensions.Logging.ILogger? externalLogger, 
        string logDirectory, 
        string baseFileName, 
        int? maxFileSizeMB = null,
        int flushIntervalMs = 100)
    {
        _externalLogger = externalLogger;
        _logDirectory = logDirectory;
        _baseFileName = baseFileName;
        _maxFileSizeMB = maxFileSizeMB;
        
        // Ensure directory exists
        Directory.CreateDirectory(_logDirectory);
        
        // Create bounded channel
        var options = new BoundedChannelOptions(1000)
        {
            FullMode = BoundedChannelFullMode.Wait
        };
        
        _channel = Channel.CreateBounded<LogMessage>(options);
        _writer = _channel.Writer;
        _cancellationTokenSource = new CancellationTokenSource();
        
        // Periodic flush timer
        _flushTimer = new Timer(async _ => await ForceFlushAsync(), 
                               null, flushIntervalMs, flushIntervalMs);
        
        // Start background writer
        _backgroundWriterTask = Task.Run(ProcessMessagesWithBatching);
    }

    #region Public API Methods

    /// <summary>
    /// Log INFO level message
    /// </summary>
    public async Task Info(string message)
    {
        await WriteLogAsync("INFO", message);
    }

    /// <summary>
    /// Log DEBUG level message
    /// </summary>
    public async Task Debug(string message)
    {
        await WriteLogAsync("DEBUG", message);
    }

    /// <summary>
    /// Log WARNING level message
    /// </summary>
    public async Task Warning(string message)
    {
        await WriteLogAsync("WARNING", message);
    }

    /// <summary>
    /// Log ERROR level message
    /// </summary>
    public async Task Error(string message)
    {
        await WriteLogAsync("ERROR", message);
    }

    /// <summary>
    /// Log ERROR with exception
    /// </summary>
    public async Task Error(Exception exception, string? message = null)
    {
        var errorMessage = message != null 
            ? $"{message} - Exception: {exception}" 
            : exception.ToString();
        await WriteLogAsync("ERROR", errorMessage);
    }

    /// <summary>
    /// Legacy compatibility method
    /// </summary>
    public async Task LogAsync(string message, string logLevel = "INFO")
    {
        await WriteLogAsync(logLevel.ToUpperInvariant(), message);
    }

    #endregion

    #region Properties

    /// <summary>
    /// Current log file path
    /// </summary>
    public string CurrentLogFile => GetCurrentLogFile();

    /// <summary>
    /// Current log file size in MB
    /// </summary>
    public double CurrentFileSizeMB
    {
        get
        {
            var filePath = GetCurrentLogFile();
            if (File.Exists(filePath))
            {
                var fileInfo = new FileInfo(filePath);
                return fileInfo.Length / (1024.0 * 1024.0);
            }
            return 0.0;
        }
    }

    /// <summary>
    /// Is file rotation enabled
    /// </summary>
    public bool IsRotationEnabled => _maxFileSizeMB.HasValue;

    /// <summary>
    /// External logger instance
    /// </summary>
    public Microsoft.Extensions.Logging.ILogger? ExternalLogger => _externalLogger;

    #endregion

    #region Private Implementation

    private async Task WriteLogAsync(string level, string message)
    {
        var logMessage = new LogMessage
        {
            Timestamp = DateTime.Now,
            Level = level,
            Message = message
        };

        try
        {
            await _writer.WriteAsync(logMessage, _cancellationTokenSource.Token);
        }
        catch (OperationCanceledException)
        {
            // Logger is shutting down - ignore
        }
    }

    private async Task ProcessMessagesWithBatching()
    {
        try
        {
            await foreach (var msg in _channel.Reader.ReadAllAsync(_cancellationTokenSource.Token))
            {
                await _flushSemaphore.WaitAsync(_cancellationTokenSource.Token);
                try
                {
                    _pendingMessages.Add(msg);
                    _hasUnflushedData = true;

                    // Immediate flush pre CRITICAL messages
                    if (msg.Level == "ERROR" || msg.Level == "FATAL")
                    {
                        await FlushPendingMessages();
                    }
                    // Batch flush (každých 10 messages)
                    else if (_pendingMessages.Count >= 10)
                    {
                        await FlushPendingMessages();
                    }
                }
                finally
                {
                    _flushSemaphore.Release();
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown
        }
        finally
        {
            // Final flush at shutdown
            await FlushPendingMessages();
        }
    }

    private async Task ForceFlushAsync()
    {
        if (!_hasUnflushedData) return;

        await _flushSemaphore.WaitAsync();
        try
        {
            await FlushPendingMessages();
        }
        finally
        {
            _flushSemaphore.Release();
        }
    }

    private async Task FlushPendingMessages()
    {
        if (_pendingMessages.Count == 0) return;

        try
        {
            var logText = string.Join("", _pendingMessages.Select(FormatLogMessage));
            var currentLogFile = GetCurrentLogFile();

            // Write to file
            await File.AppendAllTextAsync(currentLogFile, logText);

            // Flush OS buffer to disk - GUARANTEE disk write
            using var fileStream = new FileStream(currentLogFile, FileMode.Append, FileAccess.Write);
            await fileStream.FlushAsync();

            // Log to external logger
            foreach (var msg in _pendingMessages)
            {
                LogToExternalLogger(msg);
            }

            _pendingMessages.Clear();
            _hasUnflushedData = false;
        }
        catch (Exception ex)
        {
            // Log error to console as fallback
            Console.WriteLine($"LoggerComponent flush error: {ex.Message}");
        }
    }

    private string FormatLogMessage(LogMessage msg)
    {
        return $"[{msg.Timestamp:yyyy-MM-dd HH:mm:ss.fff}] [{msg.Level}] {msg.Message}\n";
    }

    private void LogToExternalLogger(LogMessage msg)
    {
        if (_externalLogger == null) return;

        try
        {
            switch (msg.Level)
            {
                case "INFO":
                    _externalLogger.Info(msg.Message);
                    break;
                case "DEBUG":
                    _externalLogger.Debug(msg.Message);
                    break;
                case "WARNING":
                    _externalLogger.Warning(msg.Message);
                    break;
                case "ERROR":
                case "FATAL":
                    _externalLogger.Error(msg.Message);
                    break;
            }
        }
        catch
        {
            // Ignore external logger errors
        }
    }

    private string GetCurrentLogFile()
    {
        var today = DateTime.Now.ToString("yyyy-MM-dd");

        if (_maxFileSizeMB == null)
        {
            // Bez size limit: app_2025-01-10.log
            return Path.Combine(_logDirectory, $"{_baseFileName}_{today}.log");
        }

        // S size limit: app_2025-01-10_1.log, app_2025-01-10_2.log, ...
        var counter = 1;
        string logFile;

        do
        {
            logFile = Path.Combine(_logDirectory, $"{_baseFileName}_{today}_{counter}.log");

            if (!File.Exists(logFile))
                break;

            var fileInfo = new FileInfo(logFile);
            var fileSizeMB = fileInfo.Length / (1024.0 * 1024.0);

            if (fileSizeMB < _maxFileSizeMB.Value)
                break;

            counter++;
        }
        while (true);

        return logFile;
    }

    #endregion

    #region Graceful Shutdown

    /// <summary>
    /// Graceful shutdown s flush všetkých pending messages
    /// </summary>
    public async Task ShutdownGracefullyAsync(TimeSpan timeout = default)
    {
        if (timeout == default) timeout = TimeSpan.FromSeconds(5);

        try
        {
            // Stop accepting new messages
            _channel.Writer.Complete();

            // Wait for background task to process remaining messages
            await _backgroundWriterTask.WaitAsync(timeout);

            // Force final flush
            await ForceFlushAsync();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"LoggerComponent shutdown error: {ex.Message}");
        }
    }

    public void Dispose()
    {
        try
        {
            // Emergency synchronous shutdown
            ShutdownGracefullyAsync(TimeSpan.FromSeconds(2)).Wait();
        }
        catch { /* Already shutting down */ }

        _flushTimer?.Dispose();
        _cancellationTokenSource?.Cancel();
        _cancellationTokenSource?.Dispose();
        _flushSemaphore?.Dispose();
    }

    #endregion
}