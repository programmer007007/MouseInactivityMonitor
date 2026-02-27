using System;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.DependencyInjection;

namespace MouseInactivityMonitor
{
    internal static class AppConfig
    {
        public static string ESP32_IP = "192.168.0.117";
        public static int INACTIVITY_TIMEOUT_SECONDS = 300;
        public static bool FAN_INITIAL_STATE_IS_ON = false;
        public static bool SEND_INITIAL_STATE_TO_DEVICE = false;
        public static string TRIGGER_ENDPOINT = "/trigger";
        public static string LOG_FILE = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "MouseMonitor", "MouseMonitor.log");
    }

    internal static class AppLog
    {
        public static void Log(string message)
        {
            string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";

            try
            {
                var logDir = Path.GetDirectoryName(AppConfig.LOG_FILE);
                if (!string.IsNullOrEmpty(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                File.AppendAllText(AppConfig.LOG_FILE, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                try
                {
                    System.Diagnostics.EventLog.WriteEntry("Application",
                        $"MouseMonitor: {logMessage}\nLogging error: {ex.Message}",
                        System.Diagnostics.EventLogEntryType.Warning);
                }
                catch
                {
                    // If all logging fails, just continue
                }
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Check if we should run as Windows Service or Console App
            bool forceTray = args.Contains("--tray");
            bool isService = !forceTray && !(System.Diagnostics.Debugger.IsAttached || args.Contains("--console"));
            bool useTray = !isService && !args.Contains("--no-tray");

            var builder = Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddSingleton<FanController>();
                    services.AddHostedService<MouseMonitorService>();
                });

            if (isService)
            {
                builder.UseWindowsService();
            }

            using IHost host = builder.Build();

            if (useTray)
            {
                TrayApplication.Start(host.Services);
            }

            host.Run();
        }
    }

    public sealed class FanController
    {
        private readonly HttpClient httpClient = new HttpClient();
        private readonly object gate = new object();
        private bool isOn = AppConfig.FAN_INITIAL_STATE_IS_ON;

        public event Action<bool>? StateChanged;

        public bool IsOn
        {
            get
            {
                lock (gate)
                {
                    return isOn;
                }
            }
        }

        public async Task ApplyInitialStateAsync()
        {
            AppLog.Log($"Fan initial state: {(IsOn ? "ON" : "OFF")}");

            if (!AppConfig.SEND_INITIAL_STATE_TO_DEVICE)
            {
                return;
            }

            await ToggleAsync("InitialState");
        }

        public async Task TurnOnAsync(string source)
        {
            await SetStateAsync(true, source);
        }

        public async Task TurnOffAsync(string source)
        {
            await SetStateAsync(false, source);
        }

        private async Task SetStateAsync(bool newState, string source)
        {
            bool already = false;
            lock (gate)
            {
                already = isOn == newState;
            }

            if (already)
            {
                AppLog.Log($"Fan already {(newState ? "ON" : "OFF")}, no trigger sent. Source={source}");
                return;
            }

            bool success = await SendTriggerAsync(source);

            if (success)
            {
                bool changed = false;

                lock (gate)
                {
                    if (isOn != newState)
                    {
                        isOn = newState;
                        changed = true;
                    }
                }

                if (changed)
                {
                    StateChanged?.Invoke(newState);
                }
            }
        }

        public async Task ToggleAsync(string source)
        {
            bool success = await SendTriggerAsync(source);
            if (!success)
            {
                return;
            }

            bool newState;
            lock (gate)
            {
                isOn = !isOn;
                newState = isOn;
            }

            StateChanged?.Invoke(newState);
        }

        private async Task<bool> SendTriggerAsync(string source)
        {
            string apiUrl = $"http://{AppConfig.ESP32_IP}{AppConfig.TRIGGER_ENDPOINT}?source={Uri.EscapeDataString(source)}";

            try
            {
                AppLog.Log($"Sending trigger to {apiUrl}");
                HttpResponseMessage response = await httpClient.GetAsync(apiUrl);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    AppLog.Log($"SUCCESS: Trigger sent. Response: {responseContent}");
                    return true;
                }

                AppLog.Log($"FAILED: HTTP {response.StatusCode} for trigger");
                return false;
            }
            catch (Exception ex)
            {
                AppLog.Log($"ERROR sending trigger: {ex.Message}");
                return false;
            }
        }
    }

    internal static class TrayApplication
    {
        public static void Start(IServiceProvider services)
        {
            var thread = new Thread(() =>
            {
                Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                var fanController = services.GetRequiredService<FanController>();
                Application.Run(new TrayAppContext(fanController, services));
            });

            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        private sealed class TrayAppContext : ApplicationContext
        {
            private readonly FanController fanController;
            private readonly NotifyIcon notifyIcon;
            private readonly Icon onIcon;
            private readonly Icon offIcon;
            private readonly Bitmap? fanImage;

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern bool DestroyIcon(IntPtr handle);

            public TrayAppContext(FanController fanController, IServiceProvider services)
            {
                this.fanController = fanController;

                fanImage = LoadFanImage();
                onIcon = CreateFanIcon(Color.FromArgb(46, 204, 113));
                offIcon = CreateFanIcon(Color.FromArgb(231, 76, 60));

                var menu = new ContextMenuStrip();
                var fanOnItem = new ToolStripMenuItem("Fan On");
                var fanOffItem = new ToolStripMenuItem("Fan Off");
                var toggleItem = new ToolStripMenuItem("Toggle Fan");
                var exitItem = new ToolStripMenuItem("Exit");

                fanOnItem.Click += async (_, __) => await fanController.TurnOnAsync("Tray");
                fanOffItem.Click += async (_, __) => await fanController.TurnOffAsync("Tray");
                toggleItem.Click += async (_, __) => await fanController.ToggleAsync("TrayToggle");
                exitItem.Click += async (_, __) =>
                {
                    if (services.GetService<IHostApplicationLifetime>() is IHostApplicationLifetime lifetime)
                    {
                        lifetime.StopApplication();
                    }
                    ExitThread();
                };

                menu.Items.Add(fanOnItem);
                menu.Items.Add(fanOffItem);
                menu.Items.Add(toggleItem);
                menu.Items.Add(new ToolStripSeparator());
                menu.Items.Add(exitItem);

                notifyIcon = new NotifyIcon
                {
                    Icon = fanController.IsOn ? onIcon : offIcon,
                    Visible = true,
                    ContextMenuStrip = menu,
                    Text = GetTrayText(fanController.IsOn)
                };

                notifyIcon.DoubleClick += async (_, __) => await fanController.ToggleAsync("TrayDoubleClick");

                fanController.StateChanged += OnFanStateChanged;
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    fanController.StateChanged -= OnFanStateChanged;
                    notifyIcon.Visible = false;
                    notifyIcon.Dispose();
                    onIcon.Dispose();
                    offIcon.Dispose();
                    fanImage?.Dispose();
                }

                base.Dispose(disposing);
            }

            private void OnFanStateChanged(bool isOn)
            {
                notifyIcon.Icon = isOn ? onIcon : offIcon;
                notifyIcon.Text = GetTrayText(isOn);
            }

            private static string GetTrayText(bool isOn)
            {
                return $"Mouse Inactivity Monitor\nFan: {(isOn ? "ON" : "OFF")}";
            }

            private Bitmap? LoadFanImage()
            {
                try
                {
                    string imagePath = Path.Combine(AppContext.BaseDirectory, "fan.png");
                    if (!File.Exists(imagePath))
                    {
                        AppLog.Log($"Tray icon fan.png not found at {imagePath}. Falling back to circle icon.");
                        return null;
                    }

                    return new Bitmap(imagePath);
                }
                catch (Exception ex)
                {
                    AppLog.Log($"Failed to load fan.png: {ex.Message}. Falling back to circle icon.");
                    return null;
                }
            }

            private Icon CreateFanIcon(Color backgroundColor)
            {
                using var bitmap = new Bitmap(16, 16);
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    using var bgBrush = new SolidBrush(backgroundColor);
                    g.FillRectangle(bgBrush, 0, 0, 16, 16);

                    if (fanImage != null)
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        var dest = new Rectangle(1, 1, 14, 14);
                        g.DrawImage(fanImage, dest);
                    }
                    else
                    {
                        using var brush = new SolidBrush(Color.Black);
                        g.FillEllipse(brush, 3, 3, 10, 10);
                    }
                }

                IntPtr hIcon = bitmap.GetHicon();
                Icon icon = Icon.FromHandle(hIcon);
                Icon cloned = (Icon)icon.Clone();
                DestroyIcon(hIcon);
                return cloned;
            }
        }
    }

    public class MouseMonitorService : BackgroundService
    {
        private POINT lastMousePosition;
        private DateTime lastActivityTime;
        private bool signalSent = false;
        private readonly FanController fanController;

        // For keyboard monitoring
        private LASTINPUTINFO lastInputInfo;

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        public MouseMonitorService()
        {
            // Ensure log directory exists
            var logDir = Path.GetDirectoryName(AppConfig.LOG_FILE);
            if (!string.IsNullOrEmpty(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
        }

        public MouseMonitorService(FanController fanController)
            : this()
        {
            this.fanController = fanController;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
                AppLog.Log("===========================================");
                AppLog.Log("  Mouse & Keyboard Activity Monitor");
                AppLog.Log("===========================================");
                AppLog.Log($"ESP32 IP: {AppConfig.ESP32_IP}");
                AppLog.Log($"Inactivity Timeout: {AppConfig.INACTIVITY_TIMEOUT_SECONDS} seconds");
                AppLog.Log($"Log File: {AppConfig.LOG_FILE}");
                AppLog.Log($"Fan Initial State: {(AppConfig.FAN_INITIAL_STATE_IS_ON ? "ON" : "OFF")}");
                AppLog.Log($"Trigger Endpoint: {AppConfig.TRIGGER_ENDPOINT}");
                AppLog.Log("===========================================");

                // Initialize mouse position with error handling
                if (!GetCursorPos(out lastMousePosition))
                {
                    AppLog.Log("WARNING: Could not get initial cursor position. Starting with (0,0)");
                    lastMousePosition = new POINT { X = 0, Y = 0 };
                }
                lastActivityTime = DateTime.Now;

                AppLog.Log("Service started - Monitoring mouse and keyboard activity...");
                await fanController.ApplyInitialStateAsync();

                while (!stoppingToken.IsCancellationRequested)
                {
                    try
                    {
                        await CheckUserActivity();
                        await Task.Delay(100, stoppingToken);
                    }
                    catch (OperationCanceledException)
                    {
                        AppLog.Log("Service is stopping...");
                        break;
                    }
                    catch (Exception ex)
                    {
                        AppLog.Log($"ERROR in main loop: {ex.Message}");
                        AppLog.Log($"Stack trace: {ex.StackTrace}");
                        await Task.Delay(1000, stoppingToken);
                    }
                }

                AppLog.Log("Service stopped.");
            }
            catch (Exception ex)
            {
                AppLog.Log($"FATAL ERROR in ExecuteAsync: {ex.Message}");
                AppLog.Log($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        private async Task CheckUserActivity()
        {
            bool activityDetected = false;
            POINT currentPosition;

            // Check mouse movement
            if (GetCursorPos(out currentPosition))
            {
                if (currentPosition.X != lastMousePosition.X || currentPosition.Y != lastMousePosition.Y)
                {
                    lastMousePosition = currentPosition;
                    activityDetected = true;
                }
            }

            // Check keyboard activity using GetLastInputInfo
            LASTINPUTINFO currentInputInfo = new LASTINPUTINFO();
            currentInputInfo.cbSize = (uint)Marshal.SizeOf(currentInputInfo);

            if (GetLastInputInfo(ref currentInputInfo))
            {
                // If input time changed, there was keyboard or mouse activity
                if (currentInputInfo.dwTime != lastInputInfo.dwTime)
                {
                    lastInputInfo = currentInputInfo;
                    activityDetected = true;
                }
            }

            // Reset timer if any activity detected
            if (activityDetected)
            {
                lastActivityTime = DateTime.Now;
                signalSent = false;
            }

            // Check inactivity duration
            TimeSpan inactivityDuration = DateTime.Now - lastActivityTime;

            if (inactivityDuration.TotalSeconds >= AppConfig.INACTIVITY_TIMEOUT_SECONDS && !signalSent)
            {
                AppLog.Log($"INACTIVITY DETECTED: No user activity for {AppConfig.INACTIVITY_TIMEOUT_SECONDS} seconds");
                await fanController.TurnOffAsync("Inactivity");
                signalSent = true;
            }
        }
    }
}
