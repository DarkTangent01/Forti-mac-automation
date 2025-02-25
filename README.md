# FortiGate MAC Address Automation

This project automates the process of adding MAC address entries and grouping them on a FortiGate firewall via SSH. It reads MAC addresses and corresponding names from an Excel file and executes the necessary configuration commands on the FortiGate device.

## Features

- **Excel Integration:** Reads MAC addresses and names from an Excel file.
- **SSH Automation:** Connects to FortiGate over SSH using Paramiko.
- **Batch Processing:** Commands are executed in batches for efficient processing.
- **Progress Indicator:** Uses tqdm for a visual progress bar during command execution.
- **Environment Variables:** Sensitive credentials and settings are loaded from a `.env` file using `python-dotenv`.

## Prerequisites

- Python 3.6 or higher
- FortiGate device accessible via SSH
- Excel file (`.xlsx`) with MAC addresses and names

## Installation

### Clone the repository:
```bash
git clone https://github.com/yourusername/fortigate-mac-automation.git
cd fortigate-mac-automation
```

### Create a virtual environment (optional but recommended):
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

### Install required Python packages:
```bash
pip install -r requirements.txt
```

If you don't have a `requirements.txt`, create one with the following content:
```txt
paramiko
openpyxl
tqdm
python-dotenv
```

## Setup

### Excel File:
Ensure you have an Excel file named `mac_addresses.xlsx` in the project directory. The file should have at least two columns with a header row:

| Name    | MAC Address        |
|---------|-------------------|
| Device1 | AA:BB:CC:DD:EE:FF |
| Device2 | 11:22:33:44:55:66 |

### Environment Variables:
Create a `.env` file in the project root and add the following lines (update the values as needed):
```env
HOST=<fortigate-ip>
USERNAME=<username>
PASSWORD=<password>
VDOM_NAME=<vdom name if you have any>
MAC_GROUP_NAME=<Group name as per your requirement>
MAX_GROUP_SIZE=256
```

## Usage

Run the script using Python:
```bash
python your_script.py
```

The script will:
- Load the configuration from the `.env` file.
- Read the MAC addresses from the `mac_addresses.xlsx` file.
- Connect to the FortiGate firewall using the provided SSH credentials.
- Execute the necessary commands to add MAC addresses and configure address groups in batches.
- Display a progress bar while processing commands.

## Troubleshooting

### Excel File Errors:
If you see an error reading the Excel file, ensure the file exists and the sheet name in the script (`Sheet1`) matches the actual sheet name.

### SSH Connection Issues:
Verify that the FortiGate device is reachable and that the credentials in the `.env` file are correct.

### Dependencies:
Ensure all required Python packages are installed. Run `pip install -r requirements.txt` to reinstall if necessary.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Feel free to submit a pull request or open an issue.

Happy automating! 🚀

