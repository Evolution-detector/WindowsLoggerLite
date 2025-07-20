# WindowsLoggerLite

Windows Logger Lite is a privacy-conscious monitoring tool designed specifically for Windows systems. It runs silently as a background service and is intended for system administrators, technical support teams, and power users who need localized, high-precision logs of hardware status and application usage for system analysis, troubleshooting, and performance review.

<br>

🔧Key Features

- Silent Background Monitoring:Runs silently in the background with elevated privileges and no user interface, ensuring zero disruption to computer users.
- Comprehensive Logging:Periodically logs dynamic hardware data, tracks application start/stop events in real time, and includes commonly used static system information in reports.
- Secure Local Persistence:All data is stored locally in password-protected, encrypted Excel files. A caching mechanism ensures that logs are not lost even in the event of a sudden shutdown.
- Zero-Configuration Self-Deployment:On first run, the program automatically sets itself up by creating a scheduled task for auto-start on boot.
- Intelligent Environment Awareness:Detects removable drives and dynamically loads dependencies (e.g., LibreHardwareMonitor) as needed, providing a more intelligent and adaptive experience.
- Multi-language Support:Automatically detects the system language and localizes all output content, including reports and message prompts.

<br>

⚠️Important Notes

To simplify deployment for administrators, the software is designed to automatically create a scheduled task on its first run to enable auto-start on boot.
This behavior might be flagged as suspicious by antivirus software on some Windows versions, which could result in the task creation failing or the executable being quarantined.
> 💡Recommendation:
Before running the software, add WindowsLoggerLite.exe to the exclusion list in Windows Defender or any other security software you're using.

This project is fully open-source, and its security is guaranteed by the contributions and reviews of developers across the community. Feel free to use it with confidence!
【SourceCode】https://github.com/Evolution-detector/WindowsLoggerLite/

<br>

🚀 Usage

After launching Windows Logger Lite, you also need to start LibreHardwareMonitor and configure it with the following "Options" settings:
- Enable Start minimized
- Enable Minimize To Tray
- Enable Minimize On Close
- Enable Run On Windows Startup
These options ensure seamless background monitoring and integration with Windows Logger Lite.

<br>

📁 Log File Storage

Daily log files are stored in D:\SystemLog\ by default. If drive D: is not available, logs will be saved in C:\SystemLog\.
All logs are automatically encrypted each day.The default password is: WindowsLogger

Recommended tools for viewing the logs:
- ONLYOFFICE (Free & open-source office suite)
- Microsoft Excel

<br>

📥 Related Downloads

LibreHardwareMonitor

Free and open-source hardware sensor monitoring tool

https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases

ONLYOFFICE Desktop Editors

Free and open-source office suite with Excel support

https://www.onlyoffice.com/download-desktop.aspx

<br>

💡 Why This Project Exists

There are many tools out there with similar functionality, but most of them have critical drawbacks. For example:

- Not truly invisible:
This tool was originally designed to run on public computers performing specific tasks — such as smart information kiosks in exhibition halls — where usage needs to be monitored without interfering with regular operation. The goal was to make the tool completely silent and unnoticeable to the end user.

- Poor multilingual support:
In my opinion, proper language support is a basic requirement that lowers the barrier to entry for all users. Many existing tools fail in this area.

- Data that's too detailed or noisy:
I wanted to help technical support staff resolve issues quickly, so I intentionally excluded some of the less relevant or overly technical data points based on my own experience.

- Missing key features:
For instance, encrypted log files aren’t meant to protect highly sensitive information — the goal is simply to prevent accidental modification or corruption, especially on public or shared machines.

Because of these gaps, I decided to build a tool that fits these specific needs myself.

