<#
.SYNOPSIS
    Collects detailed system and licensing information from an endpoint PC.

.DESCRIPTION
    This script gathers the following information:
    - Hostname
    - Serial Number
    - Operating System Information
    - Installed Microsoft Office Products
    - Licensed Products (Windows and Office) from the License Server
    - OEM Product Key Information from SoftwareLicensingService and Win32_BIOS

    The data is exported to structured CSV and JSON files for reporting purposes.

.PARAMETER ReportPath
    Specifies the destination directory for the reports. If not provided, defaults to "C:\Reports".

.PARAMETER Usage
    Displays the usage information and exits.

.NOTES
    Publisher:    PREDNY SLM s.r.o.
    Version:      24.11.0
    Date:         2024-11-22
    License:      MIT License
    Requires:     PowerShell 3.0 or higher (recommended: 5.1)
    Tested On:    Windows 11 Pro 23H2
    
    - Run this script with administrative privileges to ensure access to all required data.
    - The script creates a report directory at the specified destination.
    - To view help, run the script with the -Usage flag.

.EXAMPLE
    .\CollectEndpointData.ps1 -ReportPath "D:\CustomReports"

    Collects endpoint information and saves reports to "D:\CustomReports".

.EXAMPLE
    .\CollectEndpointData.ps1

    Collects endpoint information and saves reports to the default path "C:\Reports".

.EXAMPLE
    .\CollectEndpointData.ps1 -Usage

    Displays usage information.
#>

[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $false,
        Position = 0,
        HelpMessage = "Specify the output directory for the reports."
    )]
    [string]$ReportPath = "C:\Reports",

    [Parameter(
        Mandatory = $false,
        HelpMessage = "Displays the usage information."
    )]
    [switch]$Usage
)

# ========================
# Handle -Usage Flag
# ========================

if ($Usage) {
    Get-Help -Name $MyInvocation.MyCommand.Path -Full
    exit
}

# ========================
# Helper Functions
# ========================

# Function to translate LicenseStatus codes to human-readable descriptions
function Get-LicenseStatusDescription {
    param (
        [int]$Status
    )
    switch ($Status) {
        0 {"Unlicensed"}
        1 {"Licensed"}
        2 {"Initial Grace Period"}
        3 {"Additional Grace Period"}
        4 {"Non-Genuine Grace Period"}
        5 {"Notification"}
        6 {"Extended Grace Period"}
        Default {"Unknown Status"}
    }
}

# Function to retrieve the Hostname
function Get-Hostname {
    return $env:COMPUTERNAME
}

# Function to retrieve the Serial Number
function Get-SerialNumber {
    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS
        if ($bios -and $bios.SerialNumber) {
            return $bios.SerialNumber
        } else {
            return "N/A"
        }
    } catch {
        Write-Warning "Unable to retrieve Serial Number: $_"
        return "N/A"
    }
}

# Function to retrieve Operating System Information
function Get-OperatingSystemInfo {
    try {
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Caption, Version, OSArchitecture, BuildNumber
        return $osInfo
    } catch {
        Write-Warning "Unable to retrieve Operating System Information: $_"
        return [PSCustomObject]@{
            Caption        = "N/A"
            Version        = "N/A"
            OSArchitecture = "N/A"
            BuildNumber    = "N/A"
        }
    }
}

# Function to retrieve Installed Office Information from the Registry
function Get-InstalledOffice {
    try {
        $officePaths = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        $installedOffice = Get-ItemProperty -Path $officePaths -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*Microsoft Office*" } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
        return $installedOffice
    } catch {
        Write-Warning "Unable to retrieve Installed Office Information: $_"
        return @()
    }
}

# Function to retrieve Licensed Products from the License Server
function Get-LicensedProducts {
    try {
        $licensedProducts = Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object {
            ($_.Name -like "*Windows*" -or $_.Name -like "*Office*") -and
            $_.LicenseStatus -eq 1
        } | Select-Object Name, Description, LicenseStatus, PartialProductKey, LicenseFamily, ProductKeyChannel, IsKeyManagementServiceMachine
        return $licensedProducts
    } catch {
        Write-Warning "Unable to retrieve Licensed Products Information: $_"
        return @()
    }
}

# Function to retrieve OEM Product Key Information from SoftwareLicensingService and Win32_BIOS
function Get-OEMProductKeyInfo {
    param (
        [string]$Hostname,
        [string]$SerialNumber
    )
    try {
        Write-Verbose "Attempting to retrieve OEM Product Key from SoftwareLicensingService..."

        # Primary Retrieval: SoftwareLicensingService
        $oemInfo = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop | Select-Object OA3xOriginalProductKey

        if ($oemInfo.OA3xOriginalProductKey) {
            Write-Verbose "OEM Product Key found in SoftwareLicensingService."
            [PSCustomObject]@{
                Hostname                         = $Hostname
                SerialNumber                     = $SerialNumber
                OA3xOriginalProductKey           = $oemInfo.OA3xOriginalProductKey
                OA3xOriginalProductKeyDescription = "[4.0] Professional OEM:DM"
                OA3xOriginalProductKeyResult     = "OEM Product Key retrieved successfully from SoftwareLicensingService."
            }
        } else {
            Write-Verbose "OEM Product Key not found in SoftwareLicensingService. Attempting alternative method..."

            # Secondary Retrieval: Win32_BIOS
            $biosInfo = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop | Select-Object OA3xOriginalProductKey

            if ($biosInfo.OA3xOriginalProductKey) {
                Write-Verbose "OEM Product Key found in Win32_BIOS."
                [PSCustomObject]@{
                    Hostname                         = $Hostname
                    SerialNumber                     = $SerialNumber
                    OA3xOriginalProductKey           = $biosInfo.OA3xOriginalProductKey
                    OA3xOriginalProductKeyDescription = "N/A"
                    OA3xOriginalProductKeyResult     = "OEM Product Key retrieved successfully from Win32_BIOS."
                }
            } else {
                Write-Verbose "OEM Product Key not found in Win32_BIOS."
                [PSCustomObject]@{
                    Hostname                         = $Hostname
                    SerialNumber                     = $SerialNumber
                    OA3xOriginalProductKey           = "Not Found"
                    OA3xOriginalProductKeyDescription = "N/A"
                    OA3xOriginalProductKeyResult     = "OEM Product Key not found in both SoftwareLicensingService and Win32_BIOS."
                }
            }
        }
    } catch {
        Write-Error "Error retrieving OEM Product Key: $_"
        return [PSCustomObject]@{
            Hostname                         = $Hostname
            SerialNumber                     = $SerialNumber
            OA3xOriginalProductKey           = "Error"
            OA3xOriginalProductKeyDescription = "Error"
            OA3xOriginalProductKeyResult     = "Error retrieving OEM Product Key."
        }
    }
}

# Function to collect all endpoint data
function Collect-EndpointData {
    $hostname = Get-Hostname
    $serialNumber = Get-SerialNumber

    $data = [PSCustomObject]@{
        Hostname         = $hostname
        SerialNumber     = $serialNumber
        OperatingSystem  = Get-OperatingSystemInfo
        InstalledOffice  = Get-InstalledOffice
        LicensedProducts = Get-LicensedProducts
        OEMProductKey    = Get-OEMProductKeyInfo -Hostname $hostname -SerialNumber $serialNumber
    }
    return $data
}

# Function to export data to CSV and JSON
function Export-EndpointData {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Data
    )

    # Define the host-specific report directory
    $hostReportPath = Join-Path -Path $ReportPath -ChildPath $Data.Hostname
    if (!(Test-Path -Path $hostReportPath)) {
        try {
            New-Item -ItemType Directory -Path $hostReportPath -Force | Out-Null
        } catch {
            Write-Error "Failed to create report directory at ${hostReportPath}: $_"
            return
        }
    }

    # Export Host and Serial Number
    $hostInfo = [PSCustomObject]@{
        Hostname     = $Data.Hostname
        SerialNumber = $Data.SerialNumber
    }
    $hostInfo | Export-Csv -Path (Join-Path -Path $hostReportPath -ChildPath "HostInfo.csv") -NoTypeInformation

    # Export Operating System Information with Hostname and SerialNumber
    $osInfo = [PSCustomObject]@{
        Hostname        = $Data.Hostname
        SerialNumber    = $Data.SerialNumber
        OperatingSystem = $Data.OperatingSystem.Caption
        OSVersion       = $Data.OperatingSystem.Version
        OSArchitecture  = $Data.OperatingSystem.OSArchitecture
        OSBuildNumber   = $Data.OperatingSystem.BuildNumber
    }
    $osInfo | Export-Csv -Path (Join-Path -Path $hostReportPath -ChildPath "OperatingSystem.csv") -NoTypeInformation

    # Export Licensed Products with Hostname and SerialNumber
    if ($Data.LicensedProducts) {
        $licensedProducts = $Data.LicensedProducts | Select-Object `
            @{Name="Hostname"; Expression={$Data.Hostname}},
            @{Name="SerialNumber"; Expression={$Data.SerialNumber}},
            Name,
            Description,
            @{Name="LicenseStatus"; Expression={$_.LicenseStatus}},
            @{Name="LicenseStatusDescription"; Expression={Get-LicenseStatusDescription $_.LicenseStatus}},
            PartialProductKey,
            LicenseFamily,
            ProductKeyChannel,
            IsKeyManagementServiceMachine
        $licensedProducts | Export-Csv -Path (Join-Path -Path $hostReportPath -ChildPath "LicensedProducts.csv") -NoTypeInformation
    } else {
        Write-Warning "No Licensed Products found."
    }

    # Export Installed Office Information with Hostname and SerialNumber
    if ($Data.InstalledOffice) {
        $installedOffice = $Data.InstalledOffice | Select-Object `
            @{Name="Hostname"; Expression={$Data.Hostname}},
            @{Name="SerialNumber"; Expression={$Data.SerialNumber}},
            DisplayName,
            DisplayVersion,
            Publisher,
            InstallDate
        $installedOffice | Export-Csv -Path (Join-Path -Path $hostReportPath -ChildPath "InstalledOffice.csv") -NoTypeInformation
    } else {
        Write-Warning "No Installed Office products found."
    }

    # Export OEM Product Key Information with Hostname, SerialNumber, and Description
    if ($Data.OEMProductKey) {
        $oemProductKey = $Data.OEMProductKey | Select-Object `
            Hostname,
            SerialNumber,
            OA3xOriginalProductKey,
            OA3xOriginalProductKeyDescription,
            OA3xOriginalProductKeyResult
        $oemProductKey | Export-Csv -Path (Join-Path -Path $hostReportPath -ChildPath "OEMProductKey.csv") -NoTypeInformation
    } else {
        Write-Warning "No OEM Product Key Information found."
    }

    # Export Comprehensive JSON Report with LicenseStatusDescription
    $jsonData = [PSCustomObject]@{
        Hostname         = $Data.Hostname
        SerialNumber     = $Data.SerialNumber
        OperatingSystem  = $Data.OperatingSystem
        InstalledOffice  = $Data.InstalledOffice
        LicensedProducts = $Data.LicensedProducts | Select-Object *, @{Name="LicenseStatusDescription"; Expression={Get-LicenseStatusDescription $_.LicenseStatus}}
        OEMProductKey    = $Data.OEMProductKey
    }
    $jsonPath = Join-Path -Path $hostReportPath -ChildPath "EndpointData_$($Data.Hostname).json"
    $jsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8

    Write-Output "Data collection complete. Reports saved to $hostReportPath"
}

# ========================
# Main Execution
# ========================

# Collect data
$endpointData = Collect-EndpointData

# Export data
Export-EndpointData -Data $endpointData
