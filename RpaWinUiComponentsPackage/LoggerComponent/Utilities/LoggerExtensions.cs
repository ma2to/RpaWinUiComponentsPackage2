using Microsoft.Extensions.Logging.Abstractions;

namespace RpaWinUiComponentsPackage.LoggerComponent.Utilities;

/// <summary>
/// Helper metódy pre logovanie v LoggerComponent (používa iba Abstractions)
/// Nezávislé na AdvancedWinUiDataGrid komponente
/// </summary>
public static class LoggerExtensions
{
    // Log level konštanty pre kratší zápis
    private static readonly Microsoft.Extensions.Logging.LogLevel InfoLevel = Microsoft.Extensions.Logging.LogLevel.Information;
    private static readonly Microsoft.Extensions.Logging.LogLevel ErrorLevel = Microsoft.Extensions.Logging.LogLevel.Error;
    private static readonly Microsoft.Extensions.Logging.LogLevel DebugLevel = Microsoft.Extensions.Logging.LogLevel.Debug;
    private static readonly Microsoft.Extensions.Logging.LogLevel WarnLevel = Microsoft.Extensions.Logging.LogLevel.Warning;

    /// <summary>
    /// Log Information message
    /// </summary>
    public static void Info(this Microsoft.Extensions.Logging.ILogger? logger, string message, params object?[] args)
    {
        if (logger == null) return;
        logger.Log(InfoLevel, default, message, null, (msg, ex) => string.Format(msg, args));
    }

    /// <summary>
    /// Log Error message
    /// </summary>
    public static void Error(this Microsoft.Extensions.Logging.ILogger? logger, string message, params object?[] args)
    {
        if (logger == null) return;
        logger.Log(ErrorLevel, default, message, null, (msg, ex) => string.Format(msg, args));
    }

    /// <summary>
    /// Log Error with exception
    /// </summary>
    public static void Error(this Microsoft.Extensions.Logging.ILogger? logger, Exception exception, string message, params object?[] args)
    {
        if (logger == null) return;
        logger.Log(ErrorLevel, default, message, exception, (msg, ex) => string.Format(msg, args));
    }

    /// <summary>
    /// Log Debug message
    /// </summary>
    public static void Debug(this Microsoft.Extensions.Logging.ILogger? logger, string message, params object?[] args)
    {
        if (logger == null) return;
        logger.Log(DebugLevel, default, message, null, (msg, ex) => string.Format(msg, args));
    }

    /// <summary>
    /// Log Warning message
    /// </summary>
    public static void Warning(this Microsoft.Extensions.Logging.ILogger? logger, string message, params object?[] args)
    {
        if (logger == null) return;
        logger.Log(WarnLevel, default, message, null, (msg, ex) => string.Format(msg, args));
    }
}