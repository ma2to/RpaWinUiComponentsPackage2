namespace RpaWinUiComponentsPackage.LoggerComponent;

/// <summary>
/// Internal log message model pre Channel-based logging
/// </summary>
internal class LogMessage
{
    /// <summary>
    /// Timestamp kedy bola správa vytvorená
    /// </summary>
    public DateTime Timestamp { get; set; }

    /// <summary>
    /// Log level (INFO, DEBUG, WARNING, ERROR, FATAL)
    /// </summary>
    public string Level { get; set; } = string.Empty;

    /// <summary>
    /// Log message content
    /// </summary>
    public string Message { get; set; } = string.Empty;
}