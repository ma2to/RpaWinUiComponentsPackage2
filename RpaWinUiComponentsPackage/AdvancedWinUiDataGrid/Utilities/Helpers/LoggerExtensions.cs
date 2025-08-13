using Microsoft.Extensions.Logging.Abstractions;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Utilities.Helpers;

/// <summary>
/// Helper metódy pre logovanie v balíku (používa iba Abstractions, nie full Extensions.Logging)
/// Balík môže používať INOU Microsoft.Extensions.Logging.Abstractions
/// </summary>
public static class LoggerExtensions
{
    // Log level konštanty pre kratší zápis
    private static readonly Microsoft.Extensions.Logging.LogLevel InfoLevel = Microsoft.Extensions.Logging.LogLevel.Information;
    private static readonly Microsoft.Extensions.Logging.LogLevel ErrorLevel = Microsoft.Extensions.Logging.LogLevel.Error;
    private static readonly Microsoft.Extensions.Logging.LogLevel DebugLevel = Microsoft.Extensions.Logging.LogLevel.Debug;
    private static readonly Microsoft.Extensions.Logging.LogLevel WarnLevel = Microsoft.Extensions.Logging.LogLevel.Warning;

    /// <summary>
    /// Log Information message - ekvivalent LogInformation() ale pre Abstractions
    /// </summary>
    public static void Info(this Microsoft.Extensions.Logging.ILogger? logger, string message, params object?[] args)
    {
        if (logger == null) return;
        // Use safe string formatting that handles both structured logging and string.Format
        logger.Log(InfoLevel, default, message, null, (msg, ex) => SafeFormat(msg, args));
    }

    /// <summary>
    /// Log Error message - ekvivalent LogError() ale pre Abstractions
    /// </summary>
    public static void Error(this Microsoft.Extensions.Logging.ILogger? logger, string message, params object?[] args)
    {
        if (logger == null) return;
        logger.Log(ErrorLevel, default, message, null, (msg, ex) => SafeFormat(msg, args));
    }

    /// <summary>
    /// Log Error with exception - ekvivalent LogError(ex, message) ale pre Abstractions
    /// </summary>
    public static void Error(this Microsoft.Extensions.Logging.ILogger? logger, Exception exception, string message, params object?[] args)
    {
        if (logger == null) return;
        logger.Log(ErrorLevel, default, message, exception, (msg, ex) => SafeFormat(msg, args));
    }

    /// <summary>
    /// Log Debug message - ekvivalent LogDebug() ale pre Abstractions
    /// </summary>
    public static void Debug(this Microsoft.Extensions.Logging.ILogger? logger, string message, params object?[] args)
    {
        if (logger == null) return;
        logger.Log(DebugLevel, default, message, null, (msg, ex) => SafeFormat(msg, args));
    }

    /// <summary>
    /// Log Warning message - ekvivalent LogWarning() ale pre Abstractions
    /// </summary>
    public static void Warning(this Microsoft.Extensions.Logging.ILogger? logger, string message, params object?[] args)
    {
        if (logger == null) return;
        logger.Log(WarnLevel, default, message, null, (msg, ex) => SafeFormat(msg, args));
    }

    /// <summary>
    /// Log method entry with parameters and timestamp - rozsiahle logovanie
    /// </summary>
    public static void LogMethodEntry(this Microsoft.Extensions.Logging.ILogger? logger, string methodName, params object?[] parameters)
    {
        if (logger == null) return;
        var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
        var paramString = parameters?.Length > 0 ? string.Join(", ", parameters) : "no parameters";
        logger.Info($"🚀 METHOD ENTRY [{timestamp}]: {methodName}({paramString})");
    }

    /// <summary>
    /// Log method exit with execution time and result - rozsiahle logovanie
    /// </summary>
    public static void LogMethodExit(this Microsoft.Extensions.Logging.ILogger? logger, string methodName, TimeSpan executionTime, object? result = null)
    {
        if (logger == null) return;
        var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
        var resultString = result != null ? $" → Result: {result}" : "";
        var executionMs = Math.Round(executionTime.TotalMilliseconds, 2);
        logger.Info($"✅ METHOD EXIT [{timestamp}]: {methodName} completed in {executionMs}ms{resultString}");
    }

    /// <summary>
    /// Log data details - pre rozsiahle logovanie dát
    /// </summary>
    public static void LogDataDetails(this Microsoft.Extensions.Logging.ILogger? logger, string operationType, object data, int? count = null)
    {
        if (logger == null) return;
        var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
        var countInfo = count.HasValue ? $" Count: {count}" : "";
        
        // Log data type and sample
        var dataType = data?.GetType().Name ?? "null";
        var sampleData = GetSampleData(data, 100); // Max 100 chars sample
        
        logger.Info($"📊 DATA [{timestamp}]: {operationType} - Type: {dataType}{countInfo} Sample: {sampleData}");
    }

    /// <summary>
    /// Log performance metrics - pre monitoring výkonu
    /// </summary>
    public static void LogPerformance(this Microsoft.Extensions.Logging.ILogger? logger, string operationType, TimeSpan executionTime, int? itemCount = null, long? memoryUsed = null)
    {
        if (logger == null) return;
        var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
        var itemInfo = itemCount.HasValue ? $" Items: {itemCount}" : "";
        var memoryInfo = memoryUsed.HasValue ? $" Memory: {memoryUsed / 1024}KB" : "";
        var executionMs = Math.Round(executionTime.TotalMilliseconds, 2);
        
        logger.Info($"⚡ PERFORMANCE [{timestamp}]: {operationType} - Time: {executionMs}ms{itemInfo}{memoryInfo}");
    }

    /// <summary>
    /// Helper method to get sample data string (max length)
    /// </summary>
    private static string GetSampleData(object? data, int maxLength)
    {
        if (data == null) return "null";
        
        try
        {
            var dataString = data.ToString() ?? "null";
            if (dataString.Length <= maxLength) return dataString;
            return dataString.Substring(0, maxLength) + "...";
        }
        catch
        {
            return $"[{data.GetType().Name}]";
        }
    }

    /// <summary>
    /// Safe string formatting that handles both regular string.Format and prevents format exceptions
    /// </summary>
    private static string SafeFormat(string message, params object?[] args)
    {
        if (args == null || args.Length == 0)
            return message;
            
        try
        {
            // If the message contains {0}, {1}, etc. use string.Format
            // If it contains {SomeName}, it's structured logging, just return the message
            if (message.Contains("{0}") || message.Contains("{1}") || message.Contains("{2}") || 
                (args.Length > 0 && System.Text.RegularExpressions.Regex.IsMatch(message, @"\{\d+\}")))
            {
                return string.Format(message, args);
            }
            else
            {
                // For structured logging or plain messages, return as-is
                return message;
            }
        }
        catch
        {
            // If formatting fails, return the original message
            return message + " [FORMAT ERROR]";
        }
    }
}