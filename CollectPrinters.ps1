# Collect all the printers with IDs
# and allows to export to CSV.
# Created just for tests by Francesco V. Buccoli

param(
    [switch]$Export,
    [string]$csvPath = "printer_info.csv"
)

function Get-PrinterInfo {
    $printers = Get-WmiObject Win32_Printer

    $printerInfo = @()
    foreach ($printer in $printers) {
        $deviceId = $printer.DeviceID
        $vendorId = if ($deviceId -match 'VID_([0-9A-F]{4})') { $matches[1] } else { "N/A" }
        $productId = if ($deviceId -match 'PID_([0-9A-F]{4})') { $matches[1] } else { "N/A" }

        $printerDetails = [PSCustomObject]@{
            PrinterName = $printer.Name
            DeviceID = $deviceId
            VendorID = $vendorId
            ProductID = $productId
            ParentID = $printer.SystemName
        }

        $printerInfo += $printerDetails
    }

    $printerInfo
}

$printerData = Get-PrinterInfo

if ($Export) {
    $printerData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Printer information exported to $csvPath"
} else {
    $printerData | Format-Table
}
