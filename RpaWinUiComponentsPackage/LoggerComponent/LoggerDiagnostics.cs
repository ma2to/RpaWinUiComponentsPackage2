using RpaWinUiComponentsPackage.LoggerComponent.Utilities;

namespace RpaWinUiComponentsPackage.LoggerComponent;

/// <summary>
/// Diagnostics helper pre LoggerComponent
/// </summary>
public static class LoggerDiagnostics
{
    /// <summary>
    /// Get diagnostic info o LoggerComponent stave
    /// </summary>
    public static string GetDiagnosticInfo(LoggerComponent logger)
    {
        var info = new List<string>
        {
            $"Current Log File: {logger.CurrentLogFile}",
            $"File Size: {logger.CurrentFileSizeMB:F2} MB",
            $"Rotation Enabled: {logger.IsRotationEnabled}",
            $"External Logger: {(logger.ExternalLogger != null ? "Connected" : "None")}",
            $"Status: {(File.Exists(logger.CurrentLogFile) ? "Active" : "Not Started")}"
        };

        return string.Join("\n", info);
    }

    /// <summary>
    /// Test logging functionality
    /// </summary>
    public static async Task<bool> TestLoggingAsync(LoggerComponent logger)
    {
        try
        {
            var testMessage = $"Test message at {DateTime.Now:HH:mm:ss.fff}";
            
            await logger.Info($"INFO test: {testMessage}");
            await logger.Debug($"DEBUG test: {testMessage}");
            await logger.Warning($"WARNING test: {testMessage}");
            await logger.Error($"ERROR test: {testMessage}");
            
            // Wait a bit for background processing
            await Task.Delay(200);
            
            // Check if file exists and contains our messages
            var logFile = logger.CurrentLogFile;
            if (!File.Exists(logFile))
                return false;
                
            var content = await File.ReadAllTextAsync(logFile);
            return content.Contains(testMessage);
        }
        catch
        {
            return false;
        }
    }
}