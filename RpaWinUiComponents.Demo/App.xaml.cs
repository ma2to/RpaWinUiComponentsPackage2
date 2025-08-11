using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml;

namespace RpaWinUiComponents.Demo;

/// <summary>
/// Demo aplikácia pre testovanie RpaWinUiComponentsPackage s package reference
/// Implementuje UI Update Strategy z newProject.md
/// </summary>
public partial class App : Application
{
    /// <summary>
    /// Logger factory pre demo aplikáciu
    /// </summary>
    public static ILoggerFactory? LoggerFactory { get; private set; }

    /// <summary>
    /// Main window
    /// </summary>
    public static Window? MainWindow { get; private set; }

    public App()
    {
        this.InitializeComponent();
        
        // Setup logging pre demo aplikáciu
        SetupLogging();
    }

    /// <summary>
    /// Invoked when the application is launched
    /// </summary>
    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        MainWindow = new MainWindow();
        MainWindow.Activate();
    }

    /// <summary>
    /// Setup Microsoft.Extensions.Logging pre testing balíka
    /// </summary>
    private void SetupLogging()
    {
        LoggerFactory = Microsoft.Extensions.Logging.LoggerFactory.Create(builder =>
        {
            builder
                .AddConsole()           // Console output
                .AddDebug()             // Debug output
                .SetMinimumLevel(LogLevel.Debug);
        });
    }
}