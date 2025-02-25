import paramiko
import openpyxl
import time
import math
from tqdm import tqdm  # For progress bar
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

# FortiGate SSH Credentials (loaded from .env)
HOST = os.getenv("HOST")
USERNAME = os.getenv("USERNAME")
PASSWORD = os.getenv("PASSWORD")
VDOM_NAME = os.getenv("VDOM_NAME", "")  # Default empty if not set
MAC_GROUP_NAME = os.getenv("MAC_GROUP_NAME", "Default-Group")
MAX_GROUP_SIZE = int(os.getenv("MAX_GROUP_SIZE", 256))  # Convert to int

# Excel file path
EXCEL_FILE = "mac_addresses.xlsx"
SHEET_NAME = "Sheet1"

def read_mac_addresses_from_excel(file_path, sheet_name):
    """Reads MAC addresses and names from an Excel file"""
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        mac_entries = []

        for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):  # Skips header row
            name, mac = row
            if name and mac and isinstance(mac, str):
                mac_entries.append({"name": name.strip(), "mac": mac.strip()})

        return mac_entries

    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return []


def execute_ssh_commands(commands, batch_size=50):
    """Executes SSH commands on the FortiGate in batches with cleaner output"""
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(HOST, username=USERNAME, password=PASSWORD, timeout=30)

        ssh_shell = client.invoke_shell()
        time.sleep(1)

        ssh_shell.recv(5000)  # Clear initial output

        if VDOM_NAME:
            ssh_shell.send(f"config vdom\nedit {VDOM_NAME}\n")
            time.sleep(1)

        ssh_shell.send("config global\n")
        time.sleep(1)

        print(f"\n🚀 Adding {len(commands)} entries to FortiGate...\n")

        # Process commands in batches
        for i in tqdm(range(0, len(commands), batch_size), desc="Processing", unit="batch"):
            batch = commands[i:i + batch_size]
            for command in batch:
                ssh_shell.send(command + "\n")
                time.sleep(1.2)

        ssh_shell.send("end\nexit\n")
        client.close()

        print("\n✅ MAC addresses added successfully.")

    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    mac_entries = read_mac_addresses_from_excel(EXCEL_FILE, SHEET_NAME)

    if not mac_entries:
        print("⚠️ No MAC addresses found. Check the Excel file format.")
    else:
        print(f"📂 Found {len(mac_entries)} MAC addresses in the Excel file.")

        # Prepare commands
        commands = []
        for entry in mac_entries:
            commands.append(f"""
config firewall address
edit "{entry['name']}"
set type mac
set macaddr {entry['mac']}
next
end
""")

        # Create Address Groups in Batches
        num_groups = math.ceil(len(mac_entries) / MAX_GROUP_SIZE)
        for i in range(num_groups):
            start_idx = i * MAX_GROUP_SIZE
            end_idx = min(start_idx + MAX_GROUP_SIZE, len(mac_entries))
            group_name = f"{MAC_GROUP_NAME}_{i+1}"
            members = " ".join([f'"{entry["name"]}"' for entry in mac_entries[start_idx:end_idx]])

            commands.append(f"""
config firewall addrgrp
edit "{group_name}"
set member {members}
next
end
""")

        # Execute SSH commands in batches
        execute_ssh_commands(commands, batch_size=50)
