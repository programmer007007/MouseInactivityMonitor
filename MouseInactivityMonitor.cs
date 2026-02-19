using System;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.DependencyInjection;
using AForge.Video;
using AForge.Video.DirectShow;

namespace MouseInactivityMonitor
{
    class Program
    {
        public static bool IsConsoleMode { get; private set; }

        static void Main(string[] args)
        {
            // Check if we should run as Windows Service or Console App
            IsConsoleMode = System.Diagnostics.Debugger.IsAttached || args.Contains("--console");

            var builder = Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddHostedService<MouseMonitorService>();
                });

            if (!IsConsoleMode)
            {
                builder.UseWindowsService();
            }
            else
            {
                Console.WriteLine("===========================================");
                Console.WriteLine("  Running in CONSOLE MODE");
                Console.WriteLine("===========================================");
            }

            builder.Build().Run();
        }
    }

    public class MouseMonitorService : BackgroundService
    {
        // ====================================================================
        //  CONFIGURATION - Edit these values to customize behavior
        // ====================================================================

        // --- Network Configuration ---
        private static string ESP32_IP = "192.168.0.117";  // Your ESP32 IP address
        private static int INACTIVITY_TIMEOUT_SECONDS = 300;  // 5 minutes

        // --- Red Light Detection Thresholds ---
        // Adjust these if red light is not detected correctly
        // Current settings optimized for: Small red LED indicator in dark room
        // To make detection MORE SENSITIVE (detect dimmer/smaller red lights):
        //   - DECREASE RED_MIN_VALUE (try 3 or 2)
        //   - DECREASE RED_PIXEL_PERCENTAGE (try 0.001 or lower)
        // To make detection LESS SENSITIVE (reduce false positives):
        //   - INCREASE RED_MIN_VALUE (try 10 or 20)
        //   - INCREASE RED_PIXEL_PERCENTAGE (try 0.1 or higher)

        // For very small LED indicators, we use a different approach:
        // We look for pixels where RED is significantly higher than GREEN and BLUE
        // IMPORTANT: We only search in a very specific area where the LED is located
        private static bool USE_NARROW_SEARCH = true;      // Use narrow search area for LED

        // Narrow search area: narrower width (LED is near right edge), taller height
        private static double SEARCH_X_START_FRACTION = 0.75;  // Start at 75% from left (right 1/4 of image)
        private static double SEARCH_X_END_FRACTION = 1.0;     // End at 100% (right edge)
        private static double SEARCH_Y_START_FRACTION = 0.2;   // Start at 0% (top)
        private static double SEARCH_Y_END_FRACTION = 0.3;     // End at 50% (top half)

        // Detection thresholds - relaxed for small dim LED
        private static int RED_MIN_VALUE = 40;             // LED can be dim in dark room
        private static int MAX_GREEN_VALUE = 30;           // Green must be low (true red, not white)
        private static int MAX_BLUE_VALUE = 30;            // Blue must be low (true red, not white)
        private static double RED_DOMINANCE_RATIO = 2.0;   // Red must be at least 2x higher than green+blue average
        private static int MIN_RED_PIXELS_REQUIRED = 2;    // Need at least 2 bright red pixels

        // --- Debug/Testing Configuration ---
        private static bool SAVE_TEST_IMAGES = false;  // Save images to TestImages folder
        private static bool SAVE_GRAYSCALE_IMAGES = false; // Also save grayscale versions for analysis
        private static int FRAME_STATUS_LOG_INTERVAL_SECONDS = 3600;  // Webcam status logging (1 hour)

        // --- File Paths (usually don't need to change) ---
        private static string LOG_FILE = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "MouseMonitor", "MouseMonitor.log");
        private static string TEST_IMAGES_DIR = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "MouseMonitor", "TestImages");

        // ====================================================================
        //  END OF CONFIGURATION
        // ====================================================================

        private POINT lastMousePosition;
        private DateTime lastActivityTime;
        private bool signalSent = false;
        private readonly HttpClient httpClient = new HttpClient();

        // For keyboard monitoring
        private LASTINPUTINFO lastInputInfo;

        // For webcam capture
        private VideoCaptureDevice? videoSource;
        private Bitmap? lastFrame;
        private readonly object frameLock = new object();
        private int framesCaptured = 0;
        private DateTime lastFrameTime = DateTime.MinValue;
        private DateTime lastFrameStatusLog = DateTime.MinValue;
        private bool webcamInitialized = false;

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

            // Initialize webcam
            InitializeWebcam();
        }

        private void InitializeWebcam()
        {
            try
            {
                Log("=== Webcam Initialization ===");

                // Create test images directory first
                if (SAVE_TEST_IMAGES)
                {
                    try
                    {
                        Directory.CreateDirectory(TEST_IMAGES_DIR);
                        Log($"Test images directory: {TEST_IMAGES_DIR}");
                    }
                    catch (Exception ex)
                    {
                        Log($"WARNING: Could not create test images directory: {ex.Message}");
                    }
                }

                var videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);

                Log($"Found {videoDevices.Count} video device(s)");

                if (videoDevices.Count == 0)
                {
                    Log("WARNING: No webcam devices found. Fan status detection will be disabled.");
                    Log("Possible reasons:");
                    Log("  1. No webcam is connected");
                    Log("  2. Webcam drivers are not installed");
                    Log("  3. Another application is using the webcam");
                    Log("  4. Webcam is disabled in Device Manager");
                    return;
                }

                // List all available cameras
                for (int i = 0; i < videoDevices.Count; i++)
                {
                    Log($"  Camera {i}: {videoDevices[i].Name}");
                    Log($"           Moniker: {videoDevices[i].MonikerString}");
                }

                // Use the first available camera
                Log($"Selecting camera: {videoDevices[0].Name}");
                videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);

                // Log available resolutions
                Log($"Checking video capabilities...");
                if (videoSource.VideoCapabilities != null && videoSource.VideoCapabilities.Length > 0)
                {
                    Log($"Available resolutions: {videoSource.VideoCapabilities.Length}");
                    for (int i = 0; i < Math.Min(5, videoSource.VideoCapabilities.Length); i++)
                    {
                        var cap = videoSource.VideoCapabilities[i];
                        Log($"  [{i}] {cap.FrameSize.Width}x{cap.FrameSize.Height} @ {cap.AverageFrameRate}fps");
                    }

                    // Try to use a reasonable resolution (640x480 or first available)
                    var preferredRes = videoSource.VideoCapabilities.FirstOrDefault(
                        c => c.FrameSize.Width == 640 && c.FrameSize.Height == 480);

                    if (preferredRes != null)
                    {
                        videoSource.VideoResolution = preferredRes;
                        Log($"Selected resolution: 640x480");
                    }
                    else
                    {
                        videoSource.VideoResolution = videoSource.VideoCapabilities[0];
                        Log($"Selected resolution: {videoSource.VideoCapabilities[0].FrameSize.Width}x{videoSource.VideoCapabilities[0].FrameSize.Height}");
                    }
                }
                else
                {
                    Log("WARNING: No video capabilities found, using default settings");
                }

                Log("Attaching event handler...");
                videoSource.NewFrame += VideoSource_NewFrame;

                Log("Starting video capture...");
                videoSource.Start();

                // Wait a moment and check if it actually started
                Thread.Sleep(500);

                if (videoSource.IsRunning)
                {
                    webcamInitialized = true;
                    Log($"SUCCESS: Webcam is running");
                    Log("Waiting for first frame (this may take a few seconds)...");
                }
                else
                {
                    Log("ERROR: Webcam.Start() called but IsRunning is false");
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR initializing webcam: {ex.Message}");
                Log($"Exception type: {ex.GetType().Name}");
                Log($"Stack trace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Log($"Inner exception: {ex.InnerException.Message}");
                }
            }
        }

        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            try
            {
                lock (frameLock)
                {
                    // Dispose old frame
                    lastFrame?.Dispose();
                    // Clone the frame to keep it
                    lastFrame = (Bitmap)eventArgs.Frame.Clone();
                    lastFrameTime = DateTime.Now;
                    framesCaptured++;

                    // Log first frame capture
                    if (framesCaptured == 1)
                    {
                        Log($"SUCCESS: First frame captured! Resolution: {lastFrame.Width}x{lastFrame.Height}");

                        // Save the first frame as a test
                        if (SAVE_TEST_IMAGES)
                        {
                            try
                            {
                                string testPath = Path.Combine(TEST_IMAGES_DIR, "first_frame.jpg");
                                lastFrame.Save(testPath, ImageFormat.Jpeg);
                                Log($"First frame saved to: {testPath}");
                            }
                            catch (Exception ex)
                            {
                                Log($"ERROR saving first frame: {ex.Message}");
                            }
                        }
                    }

                    // Periodic status logging
                    if ((DateTime.Now - lastFrameStatusLog).TotalSeconds >= FRAME_STATUS_LOG_INTERVAL_SECONDS)
                    {
                        Log($"Webcam Status: {framesCaptured} frames captured, Last frame: {(DateTime.Now - lastFrameTime).TotalSeconds:F1}s ago");
                        lastFrameStatusLog = DateTime.Now;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR in VideoSource_NewFrame: {ex.Message}");
            }
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

                // Give webcam a moment to start capturing
                await Task.Delay(2000, stoppingToken);

                // In console mode, offer to test camera immediately
                if (Program.IsConsoleMode)
                {
                    Log("");
                    Log("=== CAMERA TEST MODE ===");
                    Log("Press 'T' to test camera and red light detection");
                    Log("Press 'S' to check webcam status");
                    Log("Press any other key to continue monitoring");
                    Log("========================");

                    // Start a background task to listen for test key
                    _ = Task.Run(async () =>
                    {
                        while (!stoppingToken.IsCancellationRequested)
                        {
                            if (Console.KeyAvailable)
                            {
                                var key = Console.ReadKey(true);
                                if (key.KeyChar == 't' || key.KeyChar == 'T')
                                {
                                    Log(">>> MANUAL TEST TRIGGERED <<<");
                                    await TestCameraAndDetection();
                                }
                                else if (key.KeyChar == 's' || key.KeyChar == 'S')
                                {
                                    Log(">>> WEBCAM STATUS CHECK <<<");
                                    CheckWebcamStatus();
                                }
                            }
                            await Task.Delay(100);
                        }
                    });
                }

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
            finally
            {
                CleanupWebcam();
            }
        }

        private void CleanupWebcam()
        {
            try
            {
                if (videoSource != null && videoSource.IsRunning)
                {
                    videoSource.SignalToStop();
                    videoSource.WaitForStop();
                    videoSource = null;
                    Log("Webcam stopped and cleaned up");
                }

                lock (frameLock)
                {
                    lastFrame?.Dispose();
                    lastFrame = null;
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR cleaning up webcam: {ex.Message}");
            }
        }

        public override void Dispose()
        {
            CleanupWebcam();
            httpClient?.Dispose();
            base.Dispose();
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

        private Bitmap ConvertToGrayscale(Bitmap original)
        {
            // Create a new bitmap with the same dimensions
            Bitmap grayBitmap = new Bitmap(original.Width, original.Height);

            // Convert each pixel to grayscale
            for (int y = 0; y < original.Height; y++)
            {
                for (int x = 0; x < original.Width; x++)
                {
                    Color pixel = original.GetPixel(x, y);

                    // Calculate grayscale value using standard luminosity formula
                    // This preserves perceived brightness better than simple averaging
                    int gray = (int)(pixel.R * 0.299 + pixel.G * 0.587 + pixel.B * 0.114);

                    // Create grayscale color
                    Color grayColor = Color.FromArgb(gray, gray, gray);
                    grayBitmap.SetPixel(x, y, grayColor);
                }
            }

            return grayBitmap;
        }

        private bool IsFanOff()
        {
            try
            {
                if (!webcamInitialized)
                {
                    Log("WARNING: Webcam not initialized. Cannot detect fan status.");
                    return false;
                }

                Bitmap? frameToAnalyze;
                lock (frameLock)
                {
                    if (lastFrame == null)
                    {
                        Log("WARNING: No webcam frame available for analysis");
                        Log($"Frames captured so far: {framesCaptured}");
                        Log($"Last frame time: {(lastFrameTime == DateTime.MinValue ? "Never" : lastFrameTime.ToString("HH:mm:ss"))}");
                        return false; // Assume fan is not off if we can't check
                    }
                    frameToAnalyze = (Bitmap)lastFrame.Clone();
                }

                using (frameToAnalyze)
                {
                    Log($"=== Analyzing frame for red light ===");
                    Log($"Frame size: {frameToAnalyze.Width}x{frameToAnalyze.Height}");

                    // Define search area
                    int searchStartX, searchEndX, searchStartY, searchEndY;

                    if (USE_NARROW_SEARCH)
                    {
                        // Very narrow search area where LED is located
                        searchStartX = (int)(frameToAnalyze.Width * SEARCH_X_START_FRACTION);
                        searchEndX = (int)(frameToAnalyze.Width * SEARCH_X_END_FRACTION);
                        searchStartY = (int)(frameToAnalyze.Height * SEARCH_Y_START_FRACTION);
                        searchEndY = (int)(frameToAnalyze.Height * SEARCH_Y_END_FRACTION);
                        Log($"Search area: NARROW (LED location only)");
                        Log($"  X: {searchStartX} to {searchEndX} ({SEARCH_X_START_FRACTION:P0} to {SEARCH_X_END_FRACTION:P0})");
                        Log($"  Y: {searchStartY} to {searchEndY} ({SEARCH_Y_START_FRACTION:P0} to {SEARCH_Y_END_FRACTION:P0})");
                        Log($"  Area: {searchEndX - searchStartX} x {searchEndY - searchStartY} pixels");
                    }
                    else
                    {
                        // Full image
                        searchStartX = 0;
                        searchEndX = frameToAnalyze.Width;
                        searchStartY = 0;
                        searchEndY = frameToAnalyze.Height;
                        Log($"Search area: FULL image");
                    }

                    int redPixelCount = 0;
                    int searchAreaPixels = (searchEndX - searchStartX) * (searchEndY - searchStartY);
                    int maxRedValue = 0;
                    int maxGreenValue = 0;
                    int maxBlueValue = 0;
                    int avgBrightness = 0;
                    int pixelCount = 0;

                    // Sample some pixels for debugging
                    int sampleCount = 0;
                    int maxSamples = 10;

                    // Scan the search area
                    for (int y = searchStartY; y < searchEndY; y++)
                    {
                        for (int x = searchStartX; x < searchEndX; x++)
                        {
                            pixelCount++;
                            Color pixel = frameToAnalyze.GetPixel(x, y);

                            avgBrightness += pixel.R + pixel.G + pixel.B;
                            maxRedValue = Math.Max(maxRedValue, pixel.R);
                            maxGreenValue = Math.Max(maxGreenValue, pixel.G);
                            maxBlueValue = Math.Max(maxBlueValue, pixel.B);

                            // Check if pixel is a BRIGHT RED LED (not just reddish or dark)
                            // The LED is actually bright red, so we need strict criteria
                            bool isBrightRedLED = false;

                            if (pixel.R >= RED_MIN_VALUE &&          // Red must be bright
                                pixel.G <= MAX_GREEN_VALUE &&        // Green must be low
                                pixel.B <= MAX_BLUE_VALUE)           // Blue must be low
                            {
                                // Additional check: Red must dominate significantly
                                double avgOther = (pixel.G + pixel.B) / 2.0;
                                if (avgOther == 0 || pixel.R >= avgOther * RED_DOMINANCE_RATIO)
                                {
                                    isBrightRedLED = true;
                                }
                            }

                            if (isBrightRedLED)
                            {
                                redPixelCount++;

                                // Log some red pixels for debugging
                                if (sampleCount < maxSamples)
                                {
                                    Log($"Red pixel found at ({x},{y}): R={pixel.R}, G={pixel.G}, B={pixel.B}");
                                    sampleCount++;
                                }
                            }
                        }
                    }

                    avgBrightness = avgBrightness / (pixelCount * 3);

                    double redPercentage = (double)redPixelCount / searchAreaPixels * 100;

                    Log($"Image analysis:");
                    Log($"  - Average brightness: {avgBrightness}/255");
                    Log($"  - Max R/G/B values: {maxRedValue}/{maxGreenValue}/{maxBlueValue}");
                    Log($"Red light detection results:");
                    Log($"  - Search area pixels: {searchAreaPixels}");
                    Log($"  - Red-dominant pixels found: {redPixelCount}");
                    Log($"  - Red percentage: {redPercentage:F3}%");
                    Log($"  - Detection criteria:");
                    Log($"      R >= {RED_MIN_VALUE} (bright red)");
                    Log($"      G <= {MAX_GREEN_VALUE} (low green)");
                    Log($"      B <= {MAX_BLUE_VALUE} (low blue)");
                    Log($"      R >= (G+B)/2 * {RED_DOMINANCE_RATIO} (red dominance)");
                    Log($"      Minimum pixels required: {MIN_RED_PIXELS_REQUIRED}");

                    bool fanIsOff = redPixelCount >= MIN_RED_PIXELS_REQUIRED;
                    Log($"  - Conclusion: Fan is {(fanIsOff ? "OFF (red light detected)" : "ON (no red light)")}");

                    // Save the analyzed frame for debugging
                    if (SAVE_TEST_IMAGES)
                    {
                        try
                        {
                            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                            // Save grayscale version first
                            if (SAVE_GRAYSCALE_IMAGES)
                            {
                                using (Bitmap grayFrame = ConvertToGrayscale(frameToAnalyze))
                                {
                                    string grayFilename = $"grayscale_{timestamp}.jpg";
                                    string grayPath = Path.Combine(TEST_IMAGES_DIR, grayFilename);
                                    grayFrame.Save(grayPath, ImageFormat.Jpeg);
                                    Log($"Grayscale frame saved to: {grayPath}");
                                }
                            }

                            // Create a copy with search area marked
                            using (Bitmap markedFrame = (Bitmap)frameToAnalyze.Clone())
                            using (Graphics g = Graphics.FromImage(markedFrame))
                            {
                                // Draw a rectangle around the search area
                                if (USE_NARROW_SEARCH)
                                {
                                    using (Pen pen = new Pen(Color.Yellow, 3))
                                    {
                                        g.DrawRectangle(pen, searchStartX, searchStartY,
                                            searchEndX - searchStartX - 1, searchEndY - searchStartY - 1);
                                    }

                                    // Add a semi-transparent overlay outside the search area
                                    using (SolidBrush darkBrush = new SolidBrush(Color.FromArgb(100, 0, 0, 0)))
                                    {
                                        // Left area
                                        if (searchStartX > 0)
                                            g.FillRectangle(darkBrush, 0, 0, searchStartX, frameToAnalyze.Height);
                                        // Bottom area
                                        if (searchEndY < frameToAnalyze.Height)
                                            g.FillRectangle(darkBrush, 0, searchEndY, frameToAnalyze.Width, frameToAnalyze.Height - searchEndY);
                                        // Right edge (if any)
                                        if (searchEndX < frameToAnalyze.Width)
                                            g.FillRectangle(darkBrush, searchEndX, 0, frameToAnalyze.Width - searchEndX, searchEndY);
                                        // Top-left corner
                                        g.FillRectangle(darkBrush, searchStartX, 0, searchEndX - searchStartX, searchStartY);
                                    }
                                }

                                // Mark red pixels that were found
                                if (redPixelCount > 0)
                                {
                                    using (Pen redPen = new Pen(Color.Lime, 3))
                                    {
                                        for (int y = searchStartY; y < searchEndY; y++)
                                        {
                                            for (int x = searchStartX; x < searchEndX; x++)
                                            {
                                                Color pixel = frameToAnalyze.GetPixel(x, y);
                                                bool isBrightRedLED = false;

                                                if (pixel.R >= RED_MIN_VALUE &&
                                                    pixel.G <= MAX_GREEN_VALUE &&
                                                    pixel.B <= MAX_BLUE_VALUE)
                                                {
                                                    double avgOther = (pixel.G + pixel.B) / 2.0;
                                                    if (avgOther == 0 || pixel.R >= avgOther * RED_DOMINANCE_RATIO)
                                                    {
                                                        isBrightRedLED = true;
                                                    }
                                                }

                                                if (isBrightRedLED)
                                                {
                                                    // Draw a small circle around red pixels
                                                    g.DrawEllipse(redPen, x - 2, y - 2, 4, 4);
                                                }
                                            }
                                        }
                                    }
                                }

                                string filename = $"analysis_{timestamp}_{(fanIsOff ? "OFF" : "ON")}_{redPercentage:F1}pct_{redPixelCount}px.jpg";
                                string testPath = Path.Combine(TEST_IMAGES_DIR, filename);
                                markedFrame.Save(testPath, ImageFormat.Jpeg);
                                Log($"Analysis frame saved to: {testPath}");
                                Log($"  (Yellow box = search area, Green circles = detected red pixels)");
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"ERROR saving analysis frame: {ex.Message}");
                        }
                    }

                    return fanIsOff;
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR detecting fan status: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                return false; // Assume fan is not off if detection fails
            }
        }

        private void CheckWebcamStatus()
        {
            Log("========================================");
            Log("       WEBCAM STATUS");
            Log("========================================");
            Log($"Webcam initialized: {webcamInitialized}");
            Log($"Video source exists: {videoSource != null}");

            if (videoSource != null)
            {
                try
                {
                    Log($"Video source running: {videoSource.IsRunning}");
                    if (videoSource.VideoResolution != null)
                    {
                        Log($"Resolution: {videoSource.VideoResolution.FrameSize.Width}x{videoSource.VideoResolution.FrameSize.Height}");
                    }
                }
                catch (Exception ex)
                {
                    Log($"ERROR checking video source: {ex.Message}");
                }
            }

            Log($"Total frames captured: {framesCaptured}");

            if (framesCaptured > 0)
            {
                Log($"Last frame time: {lastFrameTime:HH:mm:ss}");
                Log($"Seconds since last frame: {(DateTime.Now - lastFrameTime).TotalSeconds:F1}");
            }
            else
            {
                Log("No frames captured yet!");
                Log("");
                Log("Troubleshooting:");
                Log("1. Make sure no other app is using the webcam");
                Log("2. Check Device Manager for webcam status");
                Log("3. Try unplugging and replugging the webcam");
                Log("4. Restart the application");
            }

            lock (frameLock)
            {
                Log($"Current frame available: {lastFrame != null}");
            }

            Log("========================================");
            Log("");
        }

        private async Task TestCameraAndDetection()
        {
            try
            {
                Log("========================================");
                Log("       MANUAL CAMERA TEST");
                Log("========================================");

                // Check if we have frames
                lock (frameLock)
                {
                    if (lastFrame == null)
                    {
                        Log("ERROR: No frames captured yet!");
                        Log($"Frames captured total: {framesCaptured}");
                        Log("Wait a few seconds and try again.");
                        return;
                    }

                    Log($"Total frames captured: {framesCaptured}");
                    Log($"Last frame captured: {(DateTime.Now - lastFrameTime).TotalSeconds:F1} seconds ago");
                    Log($"Frame resolution: {lastFrame.Width}x{lastFrame.Height}");
                }

                // Save current frame
                if (SAVE_TEST_IMAGES)
                {
                    lock (frameLock)
                    {
                        if (lastFrame != null)
                        {
                            string testPath = Path.Combine(TEST_IMAGES_DIR, $"manual_test_{DateTime.Now:yyyyMMdd_HHmmss}.jpg");
                            lastFrame.Save(testPath, ImageFormat.Jpeg);
                            Log($"Current frame saved to: {testPath}");
                        }
                    }
                }

                // Test red light detection
                bool fanIsOff = IsFanOff();

                Log("========================================");
                Log($"TEST RESULT: Fan is {(fanIsOff ? "OFF" : "ON")}");
                Log("========================================");
                Log("");
            }
            catch (Exception ex)
            {
                Log($"ERROR during test: {ex.Message}");
            }
        }

        private async Task SendIRTrigger()
        {
            try
            {
                // Check fan status before sending signal
                bool fanIsOff = IsFanOff();

                if (fanIsOff)
                {
                    Log("SKIP: Fan is already OFF (red light detected). Not sending IR signal.");
                    return;
                }

                Log("Fan is ON. Sending IR trigger to turn it OFF.");

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

            // Always write to console if in console mode
            if (Program.IsConsoleMode)
            {
                Console.WriteLine(logMessage);
            }

            // Also write to file
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