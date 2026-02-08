using System;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.DependencyInjection;

namespace MouseInactivityMonitor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check if we should run as Windows Service or Console App
            bool isService = !(System.Diagnostics.Debugger.IsAttached || args.Contains("--console"));

            var builder = Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddHostedService<MouseMonitorService>();
                });

            if (isService)
            {
                builder.UseWindowsService();
            }

            builder.Build().Run();
        }
    }

    public class MouseMonitorService : BackgroundService
    {
        // Configuration
        private static string ESP32_IP = "192.168.0.117";
        private static int INACTIVITY_TIMEOUT_SECONDS = 300;
        private static string LOG_FILE = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "MouseMonitor", "MouseMonitor.log");

        private POINT lastMousePosition;
        private DateTime lastActivityTime;
        private bool signalSent = false;
        private readonly HttpClient httpClient = new HttpClient();

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
            var logDir = Path.GetDirectoryName(LOG_FILE);
            if (!string.IsNullOrEmpty(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
                Log("===========================================");
                Log("  Mouse & Keyboard Activity Monitor");
                Log("===========================================");
                Log($"ESP32 IP: {ESP32_IP}");
                Log($"Inactivity Timeout: {INACTIVITY_TIMEOUT_SECONDS} seconds");
                Log($"Log File: {LOG_FILE}");
                Log("===========================================");

                // Initialize mouse position with error handling
                if (!GetCursorPos(out lastMousePosition))
                {
                    Log("WARNING: Could not get initial cursor position. Starting with (0,0)");
                    lastMousePosition = new POINT { X = 0, Y = 0 };
                }
                lastActivityTime = DateTime.Now;

                Log("Service started - Monitoring mouse and keyboard activity...");

                while (!stoppingToken.IsCancellationRequested)
                {
                    try
                    {
                        await CheckUserActivity();
                        await Task.Delay(100, stoppingToken);
                    }
                    catch (OperationCanceledException)
                    {
                        Log("Service is stopping...");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Log($"ERROR in main loop: {ex.Message}");
                        Log($"Stack trace: {ex.StackTrace}");
                        await Task.Delay(1000, stoppingToken);
                    }
                }

                Log("Service stopped.");
            }
            catch (Exception ex)
            {
                Log($"FATAL ERROR in ExecuteAsync: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
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

            if (inactivityDuration.TotalSeconds >= INACTIVITY_TIMEOUT_SECONDS && !signalSent)
            {
                Log($"INACTIVITY DETECTED: No user activity for {INACTIVITY_TIMEOUT_SECONDS} seconds");
                await SendIRTrigger();
                signalSent = true;
            }
        }

        private async Task SendIRTrigger()
        {
            try
            {
                string apiUrl = $"http://{ESP32_IP}/trigger?source=MouseMonitor";
                Log($"Sending IR trigger to {apiUrl}");

                HttpResponseMessage response = await httpClient.GetAsync(apiUrl);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    Log($"SUCCESS: IR signal sent. Response: {responseContent}");
                }
                else
                {
                    Log($"FAILED: HTTP {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR sending IR: {ex.Message}");
            }
        }

        private void Log(string message)
        {
            string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";

            try
            {
                File.AppendAllText(LOG_FILE, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // Try to write to Windows Event Log as fallback
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
}