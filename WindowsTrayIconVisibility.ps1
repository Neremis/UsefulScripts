function ConvertTo-Rot13AndInsertZero {

    param ([string]$inputString)
    
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($inputString)
    $outputBytes = New-Object byte[] ($bytes.Length * 2)
    
    for ($i = 0; $i -lt $bytes.Length; $i++) {

        $byte = $bytes[$i]

        if ($byte -ge 65 -and $byte -le 90) {

            $outputBytes[$i * 2] = (($byte - 65 + 13) % 26) + 65

        } elseif ($byte -ge 97 -and $byte -le 122) {

            $outputBytes[$i * 2] = (($byte - 97 + 13) % 26) + 97

        } else {

            $outputBytes[$i * 2] = $byte

        }

        $outputBytes[$i * 2 + 1] = 0

    }
    
    return $outputBytes

}

function Set-TrayIconVisibility() {

    param (
        [Parameter(Mandatory=$true)]
        [string]$ApplicationName,
        [Parameter(Mandatory=$true)]
        [Int16]$Visibility
    )

    $visibilityWasSet = $false
    $headerSize = 20
    $chunkSize = 1640
    $visibilityPosition = 528

    $regTrayNotifyPath = "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\TrayNotify"
    $regTrayNotifyPSPath = (Get-Item -Path $regTrayNotifyPath).PSPath
    $iconStreamBytes = (Get-ItemProperty -Path $regTrayNotifyPSPath).IconStreams
    $iconStreamHex = [System.BitConverter]::ToString($iconStreamBytes) -replace '-', ''

    $applicationNameBytes = ConvertTo-Rot13AndInsertZero -inputString $ApplicationName
    $applicationNameHex = [System.BitConverter]::ToString($applicationNameBytes) -replace '-', ''

    if ($iconStreamHex.Contains($applicationNameHex)) {

        $chunkItems = @{}

        for ($x = 0; $x -lt [math]::Ceiling(($iconStreamBytes.Count - $headerSize) / $chunkSize); $x++) {

            $startingByte = $headerSize + ($x * $chunkSize)
            $chunkItem = New-Object byte[] $chunkSize
            [Array]::Copy($iconStreamBytes, $startingByte, $chunkItem, 0, $chunkSize)
            $chunkItems[$startingByte.ToString()] = $chunkItem

        }

        foreach($currChunkKey in $chunkItems.Keys) {

            $currEntryHex = [System.BitConverter]::ToString($chunkItems[$currChunkKey]) -replace '-', ''

            if ($currEntryHex.Contains($applicationNameHex)) {

                $iconStreamBytes[([Convert]::ToInt32($currChunkKey) + $visibilityPosition)] = $Visibility
                Set-ItemProperty -Path $regTrayNotifyPSPath -Name "IconStreams" -Value $iconStreamBytes
                $visibilityWasSet = $true
                break

            }

        }

    }

    return $visibilityWasSet

}

<#

    This is a heavily modified and optimized script, based on the original script from the Xtreme Deployment Consulting Team,
    to adjust the visibility of individual icons in the Windows tray. More information and the original script can be
    found here: https://tmintner.wordpress.com/2011/07/08/windows-7-notification-area-automation-falling-back-down-the-binary-registry-rabbit-hole/

    Visibility: 0 = Notifications only, 1 = Always hide, 2 = Always show
    Don't forget to stop explorer.exe first.

    Example: Set-TrayIconVisibility -ApplicationName "ms-teams.exe" -Visibility 2

#>
