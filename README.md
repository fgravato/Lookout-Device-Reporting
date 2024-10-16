# Android and iOS Device and Threat Reporting Tool

This script generates an Excel report of Android and iOS devices and associated threats using the Lookout API. It now includes progress indicators to show the status of various operations.

## Prerequisites

- Python 3.6 or higher
- pip (Python package installer)

## Installation

### Windows

1. Install Python from the [official website](https://www.python.org/downloads/windows/).
2. During installation, make sure to check the box that says "Add Python to PATH".
3. Open Command Prompt and run the following commands:

```
pip install requests python-dotenv openpyxl tqdm
```

### macOS

1. Install Homebrew if you haven't already:

```
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

2. Install Python using Homebrew:

```
brew install python
```

3. Open Terminal and run the following commands:

```
pip3 install requests python-dotenv openpyxl tqdm
```

## Configuration

1. Create a `.env` file in the same directory as the script with the following content:

```
REACT_APP_APPLICATION_KEY=your_application_key_here
```

Replace `your_application_key_here` with your actual Lookout API application key.

## Running the Script

### Windows

1. Open Command Prompt
2. Navigate to the directory containing the script:

```
cd path\to\script\directory
```

3. Run the script:

```
python app.py
```

### macOS

1. Open Terminal
2. Navigate to the directory containing the script:

```
cd path/to/script/directory
```

3. Run the script:

```
python3 app.py
```

## Output

The script will generate an Excel file named `device_and_threat_report.xlsx` in the same directory. This file contains two sheets:

1. "Device and Threat Report": Detailed information about each device (both Android and iOS) and its associated threats.
2. "Threat Aging": Statistics on unresolved threat aging.

You can open this Excel file to view, sort, filter, and analyze the data.

## Features

- Supports both Android and iOS devices
- Handles large numbers of devices (more than 1000) by implementing pagination
- Stores device information in a local SQLite database for faster subsequent runs
- Generates a comprehensive Excel report with device details and threat information
- Displays progress bars and status updates during execution

## Troubleshooting

If you encounter any issues:

1. Ensure that Python and pip are correctly installed and added to your system's PATH.
2. Verify that all required libraries are installed by running:
   - Windows: `pip list`
   - macOS: `pip3 list`
3. Check that the `.env` file is in the same directory as the script and contains the correct API key.
4. If you get a "ModuleNotFoundError", try reinstalling the required packages:
   - Windows: `pip install requests python-dotenv openpyxl tqdm`
   - macOS: `pip3 install requests python-dotenv openpyxl tqdm`
5. If the script seems to hang or take a very long time, check the progress bars and status messages. For organizations with thousands of devices, the process may take several minutes to complete.

For any other issues, please check the error message and consult the Python or library documentation.

## Note on Data Volume

This script is designed to handle large numbers of devices and threats. However, for very large organizations with tens of thousands of devices, the initial run of the script may take a considerable amount of time. Subsequent runs will be faster as device information is cached in the local database.

The progress bars and status messages will help you track the script's progress during long-running operations.

If you consistently have issues with timeouts or incomplete data, you may need to implement additional pagination or chunking strategies. Please contact the developer for assistance in such cases.
