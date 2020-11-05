# Lenovo PC warranty checker
This is an unofficial powershell script for requesting warranty information for PCs from Lenovo.
It's the same data you find from the [Lenovo warranty check site](https://pcsupport.lenovo.com/warrantylookup).

## Features
- [X] Re-run unfinished job where it left off
- [X] Parameter-less run
- [X] Drag'n drop run
- [X] Locally saved Client-ID (Note: Check out [Client-ID encryption](#Client-ID-Encryption))

## Columns added
It will keep the existing columns and add/replace these columns in the outputted file.
- `Name`
- `Model`
- `Manufacturer`
- `Warranty End`

Example:
```csv
Serial Number, Name,                         Model,      Manufacturer, Warranty End
001DUMMY,      THINKPAD-L480-TYPE-20LS-20LT, 20LTS5SP00, Lenovo,       24/10/2021 00:00:00
```

## Usage
> Note: The `get-lenovo-warranty.ps1` and `get-lenovo-warranty.exe` is referred to as `.ps1` and `.exe`, respectively.

### Download
Either download the [get-lenovo-warranty.ps1](/get-lenovo-warranty.ps1) script from this repository.
Or download the latest windows executable from the [Releases](https://github.com/vtfk/lenovo-pc-warranty-check/releases/latest) page.

### First time setup
1. Double-click the `.exe` or run the `.ps1` script with Powershell.
2. Enter your Lenovo Client-ID, this is provided by your Lenovo Representative.
3. Follow one of the following steps

### Run with drag'n drop
> Note: Only supported by the `.exe` file.
1. Drag your `.xlsx` or `.csv` file onto the `.exe` file.
2. Now wait for it to finish
3. When done, it will output the file to `{original-file-name}-updated.xlsx` or `{original-file-name}-updated.csv`, depending on the input.

### Run without parameters
1. Double-click the `.exe` or run the `.ps1` script with Powershell.
2. Open the Excel or CSV file.
    - If your computer supports it, an open file dialog will be shown. Select the `.xlsx` file or `.csv`, containing at least a `Serial Number` column.  
      (If a `Model` column is present it will only import serials where `Model = Lenovo` or is blank).

    - If not it will ask for a file path in the terminal. Either drop a file on the terminal or manually enter the path.
3. Now wait for it to finish
4. When done, it will output the file to `{original-file-name}-updated.xlsx` or `{original-file-name}-updated.csv`, depending on the input.

## Supported parameters
| Parameter name | Value/Type | Description |
|-|-|-|
| `FilePath` | `/path/to/the/serial-file.xlsx` | Path to the file to be imported (either `.csv` or `.xlsx`) |
| `SerialsPerRequest` | `100` | How many serials to put in each request |
| `MaxAttempts` | `3` | How many failed requests to try before a 30s sleep |
| `ConfigPath` | `.\lenovo-serial-config.json` | The config file contaning the encrypted Client-ID and API url |
| `TempFilePath` | `.\lenovo-serials.tmp.csv` | The `.csv` file containing the unfinished job, it uses this file if the job is aborted or crashes |
| `InvalidSerialsPath` | `.\lenovo-serials.invalid.csv` | The `.csv` file containing the serials that is either invalid or failed |

## Client-ID encryption
The `Client-ID` is saved encrypted in the config file and can only be unencrypted by the same computer.

**BUT!** This is only supported on Windows, any other OS (MacOS / any Linux distro) will only save it encoded (ie. NOT secured).

## LICENSE
[MIT](LICENSE)