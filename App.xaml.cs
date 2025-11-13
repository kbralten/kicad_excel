using System;
using System.IO;
using System.Drawing;
using System.Windows;
using Hardcodet.Wpf.TaskbarNotification;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

namespace KiCadExcelBridge;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    private TaskbarIcon? _notifyIcon;
    private HttpServer? _httpServer;
    private ExcelManager? _excelManager;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // When running as a tray app with no visible windows, prevent the application
        // from shutting down automatically when there are no windows open.
        this.ShutdownMode = ShutdownMode.OnExplicitShutdown;
        // Log startup entry for diagnostics (only in DEBUG builds)
        try { AppendStartupLog($"OnStartup entered {DateTime.UtcNow:u}\r\n"); } catch { }

        // Global exception handlers to capture startup/runtime errors and log them for debugging
        this.DispatcherUnhandledException += App_DispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

        _notifyIcon = new TaskbarIcon();
        _notifyIcon.ToolTipText = "KiCad Excel Bridge";
        // Try to load the provided icon from embedded resources first, fallback to file
        try
        {
            var resourceUri = new Uri("pack://application:,,,/kicad_excel.ico", UriKind.Absolute);
            var sri = Application.GetResourceStream(resourceUri);
            if (sri?.Stream != null)
            {
                using var ms = new MemoryStream();
                sri.Stream.CopyTo(ms);
                ms.Position = 0;
                _notifyIcon.Icon = new Icon(ms);
            }
            else
            {
                var exeFolder = AppDomain.CurrentDomain.BaseDirectory;
                var iconPath = Path.Combine(exeFolder, "kicad_excel.ico");
                if (File.Exists(iconPath))
                {
                    _notifyIcon.Icon = new Icon(iconPath);
                }
            }
        }
        catch
        {
            // ignore icon loading failures and continue with default
        }
        _notifyIcon.ContextMenu = new System.Windows.Controls.ContextMenu();
        var showMenuItem = new System.Windows.Controls.MenuItem { Header = "Show" };
        showMenuItem.Click += (s, args) => ShowWindow();
        _notifyIcon.ContextMenu.Items.Add(showMenuItem);

        var exitMenuItem = new System.Windows.Controls.MenuItem { Header = "Exit" };
        exitMenuItem.Click += (s, args) => Shutdown();
        _notifyIcon.ContextMenu.Items.Add(exitMenuItem);

        // Open the main window when the user double-clicks the tray icon
        try
        {
            _notifyIcon.TrayMouseDoubleClick += (s, args) => ShowWindow();
        }
        catch
        {
            // If the event isn't available for some reason, ignore and continue.
        }

        try
        {
                MainWindow = new MainWindow();
                (MainWindow as MainWindow)!.OnFileSelectionChanged += OnFileSelectionChanged;
        }
        catch (Exception ex)
        {
            LogException("Error creating MainWindow", ex);
            // Re-throw so that dotnet run also sees non-zero exit code
            throw;
        }

        // Load saved configuration and start HTTP server immediately so API is available on launch
        try
        {
            var config = ConfigurationManager.Load();
            _excelManager = new ExcelManager(config.SourceFiles ?? new System.Collections.Generic.List<string>(), config.SheetMappings ?? new System.Collections.Generic.List<SheetMapping>());
            var serverUrl = $"http://localhost:{config.ServerPort}/kicad-api/";
            _httpServer = new HttpServer(serverUrl, new HttpHandler(_excelManager).HandleRequestAsync);
            Task.Run(() => _httpServer.Start());
            try { AppendStartupLog($"HTTP server started on {serverUrl} {DateTime.UtcNow:u}\r\n"); } catch { }
        }
        catch (Exception ex)
        {
            LogException("Error starting HTTP server on startup", ex);
        }
    }

    private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
    {
        LogException("DispatcherUnhandledException", e.Exception);
        e.Handled = false;
    }

    private void CurrentDomain_UnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            LogException("CurrentDomain_UnhandledException", ex);
        }
    }

    private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        LogException("TaskScheduler_UnobservedTaskException", e.Exception);
    }

    [System.Diagnostics.Conditional("DEBUG")]
    private static void LogException(string context, Exception ex)
    {
        try
        {
            var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KiCadExcelBridge");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            var logPath = Path.Combine(folder, "crash.log");
            var text = $"[{DateTime.UtcNow:u}] {context}: {ex}\r\n";
            File.AppendAllText(logPath, text);
            // Also write a copy to the application's folder (workspace) to make it easy to read during development
            try
            {
                var localLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "crash.log");
                File.AppendAllText(localLog, text);
            }
            catch { }
        }
        catch
        {
            // swallow logging errors
        }
    }

    [System.Diagnostics.Conditional("DEBUG")]
    private static void AppendStartupLog(string line)
    {
        try
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "startup.log");
            File.AppendAllText(path, line);
        }
        catch { }
    }

    private void OnFileSelectionChanged(List<string> filePaths, List<SheetMapping> sheetMappings)
    {
        _excelManager = new ExcelManager(filePaths, sheetMappings);
        _httpServer?.Stop();
        var config = ConfigurationManager.Load();
        var serverUrl = $"http://localhost:{config.ServerPort}/kicad-api/";
        _httpServer = new HttpServer(serverUrl, new HttpHandler(_excelManager).HandleRequestAsync);
        Task.Run(() => _httpServer.Start());
    }

    private void ShowWindow()
    {
        if (MainWindow == null)
        {
            MainWindow = new MainWindow();
        }

        // If window is minimized, restore it
        try
        {
            var wnd = MainWindow as Window;
            if (wnd != null)
            {
                if (wnd.WindowState == WindowState.Minimized)
                {
                    wnd.WindowState = WindowState.Normal;
                }

                if (!wnd.IsVisible)
                {
                    wnd.Show();
                }

                // Bring to front
                wnd.Activate();

                // Sometimes Activate() alone doesn't bring to front; toggle Topmost briefly
                var originalTopmost = wnd.Topmost;
                try
                {
                    wnd.Topmost = true;
                }
                finally
                {
                    wnd.Topmost = originalTopmost;
                }
            }
        }
        catch
        {
            // Best-effort: if anything goes wrong, attempt a simple show/activate
            try
            {
                if (MainWindow.IsVisible)
                {
                    MainWindow.Activate();
                }
                else
                {
                    MainWindow.Show();
                }
            }
            catch { }
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _httpServer?.Stop();
        _notifyIcon?.Dispose();
        base.OnExit(e);
    }
}

