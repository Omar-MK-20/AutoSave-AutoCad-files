# AutoCAD Autosave Script

This script automatically saves open AutoCAD files every 15 minutes and manages backup storage by deleting old backups. It's designed to work with Windows and AutoCAD.

## Features

- Automatically saves open AutoCAD files every 15 minutes
- Creates timestamped backups of all open files
- Automatically deletes backups older than 1 day to save space
- Detailed logging for troubleshooting
- Works with files containing special characters in their names

## Prerequisites

- Windows operating system
- Python 3.6 or higher
- AutoCAD installed and running
- Administrative privileges (for Task Scheduler setup)

## Installation

1. Clone or download this repository to your local machine.

2. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

3. The script will create a backup directory at `%USERPROFILE%\AutoCAD_Backups` when it first runs.

## Usage

### Manual Run

To run the script manually:

1. Make sure AutoCAD is running with your files open
2. Open a command prompt
3. Navigate to the script directory
4. Run:
   ```bash
   python autocad_autosave.py
   ```

### Automatic Setup

To set up the script to run automatically every 15 minutes:

1. Open Task Scheduler (search for it in Windows)
2. Click "Create Basic Task"
3. Set the following options:
   - Name: "AutoCAD Autosave"
   - Description: "Automatically saves open AutoCAD files every 15 minutes"
   - Trigger: "Daily"
   - Start time: Set to when you want it to begin
   - Check "Repeat task every: 15 minutes"
   - Set "for a duration of: 1 day"
   - Action: "Start a program"
   - Program/script: Browse to `run_autosave.bat`
   - Start in: The directory containing the script

4. After creating the task, right-click it and select "Properties"
5. Go to the "Settings" tab and check:
   - "Run task as soon as possible after a scheduled start is missed"
   - "If the task is already running, then the following rule applies: Do not start a new instance"

## How It Works

1. The script connects to the running AutoCAD instance
2. It identifies all open AutoCAD files
3. For each open file:
   - Creates a timestamped backup in the backup directory
   - Logs the backup operation
4. Cleans up old backups (older than 1 day)
5. Logs all operations to `autocad_autosave.log`

## Troubleshooting

If the script isn't working as expected:

1. Check the log file (`autocad_autosave.log`) for error messages
2. Ensure AutoCAD is running when the script executes
3. Verify you have write permissions in your user directory
4. Make sure the files aren't locked by another process

Common issues:
- "No open AutoCAD files found" - Make sure AutoCAD is running with files open
- "File does not exist" - The script couldn't find the file at the expected location
- "Error connecting to AutoCAD" - AutoCAD might not be running or the COM interface is not available

## Customization

You can modify the following settings in the script:

- `BACKUP_DIR`: Change the backup directory location
- `DELETE_AFTER_DAYS`: Change how long backups are kept
- `SAVE_INTERVAL`: Change how often the script runs (in Task Scheduler)

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 