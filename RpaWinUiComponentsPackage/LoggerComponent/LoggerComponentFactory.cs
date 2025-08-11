using Microsoft.Extensions.Logging;

namespace RpaWinUiComponentsPackage.LoggerComponent;

/// <summary>
/// Factory methods pre LoggerComponent initialization
/// </summary>
public static class LoggerComponentFactory
{
    /// <summary>
    /// Create LoggerComponent from ILoggerFactory
    /// </summary>
    public static LoggerComponent FromLoggerFactory<T>(
        ILoggerFactory factory,
        string logDirectory,
        string baseFileName = "app",
        int? maxFileSizeMB = null,
        int flushIntervalMs = 100)
    {
        var logger = factory.CreateLogger<T>();
        return new LoggerComponent(logger, logDirectory, baseFileName, maxFileSizeMB, flushIntervalMs);
    }

    /// <summary>
    /// Create LoggerComponent bez file rotation
    /// </summary>
    public static LoggerComponent WithoutRotation(
        ILogger? logger,
        string logDirectory,
        string baseFileName = "app",
        int flushIntervalMs = 100)
    {
        return new LoggerComponent(logger, logDirectory, baseFileName, null, flushIntervalMs);
    }

    /// <summary>
    /// Create LoggerComponent s file rotation
    /// </summary>
    public static LoggerComponent WithRotation(
        ILogger? logger,
        string logDirectory,
        string baseFileName = "app",
        int maxFileSizeMB = 10,
        int flushIntervalMs = 100)
    {
        return new LoggerComponent(logger, logDirectory, baseFileName, maxFileSizeMB, flushIntervalMs);
    }

    /// <summary>
    /// Create LoggerComponent pre testing (temp directory)
    /// </summary>
    public static LoggerComponent ForTesting(
        ILogger? logger = null,
        string baseFileName = "test",
        int flushIntervalMs = 50)
    {
        var tempDirectory = Path.Combine(Path.GetTempPath(), "LoggerComponentTests");
        return new LoggerComponent(logger, tempDirectory, baseFileName, null, flushIntervalMs);
    }
}