# SecureCRT-get-data
Execute multiple commands on multiple Cisco devices and capture the output. Sort of network automation :)
Exports to plain-text and CSV files (useful for filtering out / reviewing the data)

## Installation
Just copy the `GetDataFromDevices.vbs` somewhere. Then edit `GetDataFromDevices.vbs` to update the constants to your needs:

```
	Const safewordUsername = "alexander_ivashkin" ' Username to use for login
	Const staticUsername = "cisco" ' Login to use for locally managed devices
	Const staticPassword = "cisco" ' Password to use locally managed devices
	Const DEVICE_FILE_PATH = "devices.txt"
	Const OUTPUT_FILE_PATH = "output.txt"
	Const COMMANDS_FILE_PATH = "commands.txt"
	Const PASSWORDS_FILE_PATH = "passwords.txt"
	Const RESULTS_FILE_PATH = "devices_processed.csv"
	Const CSV_FILE_PATH = "output.csv"
	Const waitingTimeout = 30 ' How long to wait for a response from devices (in seconds)
```

## Usage

Run SecureCRT. Login onto your jumpbox.
Fill the input files with data (devices, commands and passwords).
From the SecureCRT, open "Script - Run..." menu.
Choose the `GetDataFromDevices.vbs`
Sit back, enjoy, have a cuppa and let the robots do the hard work for you!
