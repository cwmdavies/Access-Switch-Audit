**Network Automation Utilities**
A modular Python package for auditing Cisco switch access ports and exporting results to Excel.  
Main script: audit_access_ports.py  
Modules: Secure credential management, SSH jump server support, Netmiko helpers, and configuration.  

**Features**
Cisco Access-Port Audit: Scans multiple devices, analyzes all physical ports, and flags stale ports.  
Excel Output: Generates a filters-only workbook with a summary sheet first.  
Progress Bar: Event-driven, shows real-time progress.  
Jump Server Support: Maintains a persistent SSH jump connection for device access.  
Credential Management: Securely retrieves credentials from Windows Credential Manager or prompts interactively.  
Modular Design: All core logic is split into reusable modules.  

**Directory Structure**
Modules/
├── audit_access_ports.py      # Main script (at root)
└── Modules/
    ├── __init__.py
    ├── config.py
    ├── credentials.py
    ├── jump_manager.py
    ├── netmiko_utils.py 

**Requirements**
Python 3.8+
Netmiko
Paramiko
pandas
openpyxl

pip install netmiko paramiko pandas openpyxl pywin32

**Usage**
1. Prepare Device List
Create a text file (e.g., devices.txt) with one device IP or hostname per line.

2. Run the Audit
From the parent directory of Modules, run:
python -m main.py --devices devices.txt --output access_port_audit.xlsx

Options:  
--devices: Path to device list file (required)  
--output: Output Excel file name (default: access_port_audit.xlsx)  
--workers: Max concurrent device sessions (default: 10)  
--stale-days: Days threshold for stale port analysis (default: 30, set 0 to disable)  
--debug: Enable verbose logging  

3. Credentials
The script will attempt to retrieve credentials from Windows Credential Manager (MyApp/ADM).
If not found, you will be prompted for username and password.

**Module Overview**
audit_access_ports.py: Main entry point. Handles argument parsing, concurrency, Excel output, and orchestrates the audit.  
credentials.py: get_secret_with_fallback() retrieves credentials securely.  
jump_manager.py: JumpManager class manages SSH jump server connections for Netmiko.  
config.py: Stores static configuration variables (e.g., JUMP_HOST).  
netmiko_utils.py: connect_to_device() wraps Netmiko connections, supporting jump server if needed.  

**Example**
python .audit_access_ports.py --devices devices.txt --output results.xlsx --workers 5 --stale-days 60

**License**
MIT License  

**Author**
Christopher Davies
Email: chris.davies@weavermanor.co.uk