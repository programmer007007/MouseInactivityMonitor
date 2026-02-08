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
ESP32_IP = "192.168.0.117"              // IP address of your ESP32 device
INACTIVITY_TIMEOUT_SECONDS = 300        // Timeout in seconds (5 minutes)
LOG_FILE = "%ProgramData%/MouseMonitor/MouseMonitor.log"
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

## Troubleshooting

- **Service won't start**: Check Windows Event Log for errors
- **IR signal not sent**: Verify ESP32 IP address and network connectivity
- **Activity not detected**: Ensure the application has proper permissions
- **Logging issues**: Check that the log directory exists and is writable

## License

This project is provided as-is for personal use.
