# Mouse Inactivity Monitor

A Windows service application that monitors mouse and keyboard activity and triggers an IR signal via ESP32 after a period of inactivity.

## Overview

This application runs as a Windows service (or console application for debugging) and continuously monitors user input. When no mouse movement or keyboard activity is detected for a specified duration, it sends an HTTP request to an ESP32 device to trigger an infrared (IR) signal.

## Features

- **Dual-Mode Operation**: Can run as a Windows service or console application
- **Activity Monitoring**: Tracks both mouse movement and keyboard input
- **Configurable Timeout**: Default 300 seconds (5 minutes) of inactivity
- **ESP32 Integration**: Sends HTTP trigger to ESP32 device for IR signal transmission
- **Comprehensive Logging**: Logs all activities to a file with fallback to Windows Event Log
- **Error Handling**: Robust exception handling and recovery

## Configuration

Default settings (modify in `MouseInactivityMonitor.cs`):

```csharp
// Network & Timing
ESP32_IP = "192.168.0.117"              // IP address of your ESP32 device
INACTIVITY_TIMEOUT_SECONDS = 300        // Timeout in seconds (5 minutes)

// Red Light Detection Thresholds
RED_THRESHOLD = 100                      // Minimum red value (0-255)
GREEN_MAX = 80                           // Maximum green value allowed
BLUE_MAX = 80                            // Maximum blue value allowed
RED_PIXEL_PERCENTAGE = 0.5               // % of pixels that must be red

// Diagnostics
SAVE_TEST_IMAGES = true                  // Save test images for verification
FRAME_STATUS_LOG_INTERVAL_SECONDS = 30   // Log webcam status every N seconds

// Paths
LOG_FILE = "%ProgramData%/MouseMonitor/MouseMonitor.log"
TEST_IMAGES_DIR = "%ProgramData%/MouseMonitor/TestImages"
```

## Requirements

- .NET 8.0 or later
- Windows operating system
- ESP32 device with HTTP endpoint at `/trigger`

## Building the Application

### Debug Build
```bash
dotnet build
```

### Release Build (Self-contained executable)
```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

The executable will be located at:
```
bin\Release\net8.0-windows\win-x64\publish\MouseInactivityMonitor.exe
```

## Running the Application

### Console Mode (for testing)
```bash
MouseInactivityMonitor.exe --console
```

### As Windows Service
1. Copy the executable to a permanent location (e.g., `C:\MouseWatcher\`)
2. Install as a Windows service using `sc create` or Task Scheduler
3. Start the service

Example using Task Scheduler (as shown in build.txt):
```bash
schtasks /run /tn "MouseMonitor"
```

## How It Works

1. **Initialization**: The service starts and initializes mouse position tracking
2. **Continuous Monitoring**: Every 100ms, it checks for:
   - Mouse cursor position changes
   - Keyboard input using Windows API (`GetLastInputInfo`)
3. **Activity Detection**: Any mouse or keyboard activity resets the inactivity timer
4. **Trigger Condition**: After the configured timeout period with no activity:
   - Sends HTTP GET request to `http://{ESP32_IP}/trigger?source=MouseMonitor`
   - Logs the event
   - Prevents duplicate triggers until activity resumes

## Log File Location

Logs are stored at:
```
%ProgramData%\MouseMonitor\MouseMonitor.log
```

Typical path: `C:\ProgramData\MouseMonitor\MouseMonitor.log`

## API Endpoint

The ESP32 device should expose an HTTP endpoint:
```
GET http://{ESP32_IP}/trigger?source=MouseMonitor
```

## Dependencies

- Microsoft.Extensions.Hosting (v10.0.2)
- Microsoft.Extensions.Hosting.WindowsServices (v10.0.2)

## Use Cases

- Automatic TV/display power off after user inactivity
- Smart home automation based on computer usage
- Energy saving by triggering devices when computer is idle
- Presence detection for home automation systems

## Diagnostic Features

The application includes comprehensive diagnostics to verify everything is working:

### Webcam Verification

1. **Startup Diagnostics**:
   - Lists all detected cameras
   - Shows selected camera and available resolutions
   - Captures and saves first frame to verify camera is working

2. **Test Images**:
   - First frame saved as: `TestImages/first_frame.jpg`
   - Analysis frames saved with timestamp: `TestImages/analysis_YYYYMMDD_HHMMSS_[ON/OFF]_XX.Xpct.jpg`
   - Check these images to verify camera position and red light visibility

3. **Periodic Status Logs**:
   - Every 30 seconds: frame count and last capture time
   - Confirms camera is continuously capturing

4. **Red Light Analysis**:
   - Sample pixel RGB values logged
   - Total red pixel count and percentage
   - Decision reasoning (why fan is considered ON or OFF)

### How to Verify It's Working

**Step 1: Check Camera Detection**
```bash
# Run in console mode
MouseInactivityMonitor.exe --console

# Look for these messages:
# "Found X video device(s)"
# "Camera 0: [Your Camera Name]"
# "SUCCESS: Webcam started successfully"
# "SUCCESS: First frame captured!"
```

**Step 2: Check Test Images**
```
Open: C:\ProgramData\MouseMonitor\TestImages\first_frame.jpg
- Verify the camera is pointed at the fan
- Verify you can see the red indicator light (if fan is off)
```

**Step 3: Test Red Light Detection**
```bash
# Wait for inactivity timeout (or temporarily set INACTIVITY_TIMEOUT_SECONDS to 10)
# Check the logs for:
# "=== Analyzing frame for red light ==="
# "Sample pixel (x,y): R=XXX, G=XXX, B=XXX"
# "Red percentage: XX.XX%"
# "Conclusion: Fan is OFF/ON"
```

**Step 4: Verify Analysis Images**
```
Open: C:\ProgramData\MouseMonitor\TestImages\analysis_*.jpg
- Check the saved images show what the camera sees
- Verify red light is visible when fan is off
- Adjust RED_THRESHOLD values if needed
```

## Troubleshooting

- **Service won't start**: Check Windows Event Log for errors
- **No cameras found**: Ensure webcam is connected and drivers installed
- **First frame never captured**: Check camera permissions, try different camera
- **Red light not detected**:
  - Check test images to verify camera angle
  - Adjust `RED_THRESHOLD`, `GREEN_MAX`, `BLUE_MAX` values
  - Lower `RED_PIXEL_PERCENTAGE` if red light is small
- **IR signal not sent**: Verify ESP32 IP address and network connectivity
- **Activity not detected**: Ensure the application has proper permissions
- **Logging issues**: Check that the log directory exists and is writable

## License

This project is provided as-is for personal use.
