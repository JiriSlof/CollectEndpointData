# CollectEndpointData.ps1

![PowerShell](https://img.shields.io/badge/PowerShell-7.0%2B-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [Parameters](#parameters)
  - [Examples](#examples)
- [Output](#output)
- [Troubleshooting](#troubleshooting)
- [License](#license)
- [Acknowledgements](#acknowledgements)


## Overview

**CollectEndpointData.ps1** is a PowerShell script designed to gather detailed system and licensing information from endpoint PCs within an organization. It efficiently collects data such as the hostname, serial number, operating system details, installed Microsoft Office products, licensed Windows and Office products, and OEM product key information. The collected data is then exported into structured CSV and JSON files, facilitating easy reporting and analysis.

## Features

- **Comprehensive Data Collection**: Retrieves critical system and licensing information, including:
  - Hostname
  - Serial Number
  - Operating System Information
  - Installed Microsoft Office Products
  - Licensed Windows and Office Products
  - OEM Product Key Information

- **Structured Reporting**: Exports data into organized CSV and JSON formats, ensuring compatibility with various reporting tools and platforms.

- **Customizable Output Directory**: Allows users to specify a custom directory for saving reports, enhancing flexibility and organization.

- **Built-in Help System**: Provides detailed usage instructions and parameter descriptions to guide users.

## Prerequisites

Before running the script, ensure that the following prerequisites are met:

- **PowerShell Version**: PowerShell **3.0** or higher is required to run this script. However, **PowerShell 5.1** is **strongly recommended** for optimal performance, enhanced security features, and improved cmdlet support. You can check your PowerShell version with:

```powershell
$PSVersionTable.PSVersion
```

- **Administrative Privileges**: The script must be executed with administrative rights to access all required system and licensing information.

- **Execution Policy**: Ensure that your system's execution policy allows running scripts. You can check the current policy with:

```powershell
Get-ExecutionPolicy
```

If necessary, set the execution policy to `RemoteSigned` or `Unrestricted` for the current user:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Note**: Always adhere to your organization's security guidelines when modifying execution policies.


## Installation

 1. **Download the Script**:

    - Clone the repository or download the `CollectEndpointData.ps1` script directly.

      ```bash
      git clone https://github.com/JiriSlof/CollectEndpointData.git
      ```

    - Alternatively, download the script from your preferred source and place it in a desired directory, e.g., `C:\Scripts`.

 2. **Verify Script Integrity**:

    - Ensure that the script has not been tampered with and is safe to execute, especially if downloaded from an external source.

## Usage

### Parameters

 The script accepts the following parameters:

 | Parameter     | Type    | Description                                                   | Required | Default Value  |
 |---------------|---------|---------------------------------------------------------------|----------|----------------|
 | `-ReportPath` | `String` | Specifies the destination directory for the reports.         | No       | `C:\Reports`   |
 | `-Usage`      | `Switch` | Displays the usage information and exits the script.        | No       | N/A            |

### Examples

 1. **Running the Script with Default Settings**:

    Collects endpoint information and saves reports to the default path `C:\Reports`.

    ```powershell
    .\CollectEndpointData.ps1
    ```

 2. **Specifying a Custom Report Path**:

    Collects endpoint information and saves reports to a custom directory, e.g., `D:\CustomReports`.

    ```powershell
    .\CollectEndpointData.ps1 -ReportPath "D:\CustomReports"
    ```

 3. **Displaying Usage Information**:

    Displays detailed usage instructions and exits the script.

    ```powershell
    .\CollectEndpointData.ps1 -Usage
    ```

## Output

 Upon execution, the script generates structured reports containing the collected data. The reports are organized into a directory named after the hostname of the endpoint within the specified `ReportPath`.

### Report Structure

 ```
 C:\Reports\Hostname\
 ├── HostInfo.csv
 ├── OperatingSystem.csv
 ├── LicensedProducts.csv
 ├── InstalledOffice.csv
 ├── OEMProductKey.csv
 └── EndpointData_Hostname.json
 ```

 - **HostInfo.csv**: Contains the hostname and serial number.
 - **OperatingSystem.csv**: Details about the operating system, including version, architecture, and build number.
 - **LicensedProducts.csv**: Information on licensed Windows and Office products, including license status descriptions.
 - **InstalledOffice.csv**: Lists installed Microsoft Office products with version and installation date.
 - **OEMProductKey.csv**: OEM product key information retrieved from the system.
 - **EndpointData_Hostname.json**: Comprehensive JSON report encompassing all collected data.

## Troubleshooting

 If you encounter issues while running the script, consider the following solutions:

 1. **Execution Policy Errors**:

    - **Symptom**: Errors indicating that scripts cannot be run.
    - **Solution**: Adjust the execution policy as described in the [Prerequisites](#prerequisites) section.

 2. **Permission Denied**:

    - **Symptom**: Access denied errors when attempting to retrieve certain system information.
    - **Solution**: Run PowerShell with administrative privileges:
      - Right-click on the PowerShell icon.
      - Select **"Run as Administrator"**.

 3. **Missing Parameters**:

    - **Symptom**: Errors about unrecognized parameters like `-Help`.
    - **Solution**: Use the `-Usage` flag to access help information instead.

 4. **PowerShell Version Compatibility**:

    - **Symptom**: Script fails to execute due to incompatible PowerShell version.
    - **Solution**: Update PowerShell to version 5.1 or higher.

 5. **Script Execution Issues**:

    - **Symptom**: Unexpected behavior or incomplete data collection.
    - **Solution**:
      - Ensure all prerequisites are met.
      - Verify that the report path is accessible and has sufficient permissions.
      - Check for network issues if running in a domain environment.

## License

 This project is licensed under the [MIT License](LICENSE).

## Acknowledgements

 - Inspired by best practices in PowerShell scripting and system administration.
 - Thanks to the PowerShell community for continuous support and knowledge sharing.
