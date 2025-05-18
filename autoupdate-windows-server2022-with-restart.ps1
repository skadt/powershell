Function Install-WindowsUpdatesNow {
 
 Write-Output "`n" (Get-Date -Format "yyyy-MM-dd HH:mm:ss") "Включаем обновление других продуктов Microsoft"
$ServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" -Verbose
$ServiceManager.ClientApplicationID = "My App"
$ServiceManager.AddService2( "7971f918-a847-4430-9279-4a52d1efe18d",7,"")


# RESETS ALL VARIABLES FOR FUNCTION
    $UpdateSession = $null
    $UpdateSearcher = $null
    $SearchResults = $null
    $cnt = $null
    $UpdateList = @()
    $UpdateSelection = $null
    $UpdatesToDownload = $null
    $UTitle = $null
    $Update = $null
    $Downloader = $null
    $UpdatesToInstall = $null
    $Installer = $null
    $InstallationResult = $null

    # EXECUTES SYSTEM SCAN FOR RELAVENT PATCHES
    Write-Host "Beginning system scan. This may take a while..."
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $SearchResults = $UpdateSearcher.Search("IsInstalled=0")

    # CREATES OUTPUT OBJECT FOR EACH REQUIRED PATCH
    If ($SearchResults.Updates.Count -eq 0) {
        ""
        Write-Host -ForegroundColor Yellow "System is up-to-date. Break..."

        Start-Sleep -Seconds 3
        break
    }


$cnt = 1
    $SearchResults.Updates | Foreach {$UpdateList += New-Object -TypeName psobject -Property @{
            SelNumber = $cnt
            Title = $_.Title
            Description = "$cnt) $($_.Title)"
        }
        $cnt++
    }

    # RETURNS TEXT DATA
    ""
    $UpdateList.Description
    ""

    # FILTERS TO ONLY SELECTED UPDATES
    $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl

        Foreach ($KB in $SearchResults.Updates) {
            $UpdatesToDownload.Add($KB) | Out-Null}


    # EXECUTES DOWNLOAD OF ALL SELECTED UPDATES
    $Downloader = $UpdateSession.CreateUpdateDownloader()
    $Downloader.Updates = $UpdatesToDownload
    ""
    Write-Host "Beginning download of patches. This may take a while..."
    $Downloader.Download() | Out-Null

    # FILTERS TO ONLY DOWNLOADED UPDATES
    $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    Foreach ($KB in $SearchResults.Updates) {
        If ($KB.IsDownloaded -eq $true) {
            $UpdatesToInstall.Add($KB) | Out-Null
        } Else {
            Write-Host -ForegroundColor Red "$($KB.Title) failed to download. Skipping..."
        }
    }



    # EXECUTES INSTALLATION OF ALL DOWNLOADED UPDATES
    $Installer = $UpdateSession.CreateUpdateInstaller()
    $Installer.Updates = $UpdatesToInstall
    Write-Host "Beginning installation of patches. This may take awhile..."
    $InstallationResult = $Installer.Install()

    # REUTRNS STATUS OF ALL UPDATES ATTEMPTED TO BE INSTALLED
    $cnt = 0
    Foreach ($Result in $Installer.Updates) {
        Switch ($InstallationResult.GetUpdateResult($cnt).ResultCode) {
            0 {Write-Host "$($Result.Title): Not Started" ; break}
            1 {Write-Host "$($Result.Title): In Progress" ; break}
            2 {Write-Host "$($Result.Title): Succeeded" ; break}
            3 {Write-Host "$($Result.Title): Succeeded with errors" ; break}
            4 {Write-Host "$($Result.Title): Failed" ; break}
            5 {Write-Host "$($Result.Title): Process stopped before completing" ; break}
        }
        $cnt++
    }

Start-Sleep -Seconds 60

Restart-Computer -Force -Confirm:$false 

}

# SIG # Begin signature block
# MIII8gYJKoZIhvcNAQcCoIII4zCCCN8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUf8wV38ucLxhn5f1u37t1Q8+S
# VM6gggZPMIIGSzCCBTOgAwIBAgITLAAAABzli0wowWnTYwAAAAAAHDANBgkqhkiG
# 9w0BAQsFADBVMRMwEQYKCZImiZPyLGQBGRYDbG9jMRIwEAYKCZImiZPyLGQBGRYC
# aGwxEzARBgoJkiaJk/IsZAEZFgNhZG0xFTATBgNVBAMTDGFkbS1EQzItQ0EwMTAe
# Fw0yNTA1MTQxNDA0MzBaFw0yNjA1MTQxNDA0MzBaMIHOMRMwEQYKCZImiZPyLGQB
# GRYDbG9jMRIwEAYKCZImiZPyLGQBGRYCaGwxEzARBgoJkiaJk/IsZAEZFgNhZG0x
# EzARBgNVBAsMCl9IYXBweUxvb2sxRjBEBgNVBAsMPV/QkNC00LzQuNC90LjRgdGC
# 0YDQsNGC0LjQstC90YvQtSDRg9GH0LXRgtC90YvQtSDQt9Cw0L/QuNGB0LgxMTAv
# BgNVBAMMKNCa0L7QstCw0LvQtdCyINCh0LXRgNCz0LXQuSAo0LDQtNC80LjQvSkw
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC2y4DMI1Ctv2VIF4hryrfb
# wdT6ELiU7q4Ap919Zt4DqUjSJj/HI1XOgYMZiUSwSvcQrTcMqFO97QlY+qSuSZTj
# Plk+JofXv5W3ZZbuU/fb0gojCqMKH/HnVz8A2UMd2JJpVCIWSsa/CTZfYf5TUYcL
# fkekXMIs0rc1cboNelAbBL00hx6Pa50Qsp/6PUQL0QaYzETSe+9IbYh2QIkZVNe3
# 1yDhK4FJGQT3bbu8cty3DM4wgJtKAMFREBO+/fQI4iuFId0TJau7ohh3x9/1F0yb
# k+YAbgPxSnriczbZj38eeTg0Zpq0shU2c0i8cjHHwMuYeVjZBPfXwO6IWEJ+EFqh
# AgMBAAGjggKYMIIClDAlBgkrBgEEAYI3FAIEGB4WAEMAbwBkAGUAUwBpAGcAbgBp
# AG4AZzATBgNVHSUEDDAKBggrBgEFBQcDAzAOBgNVHQ8BAf8EBAMCB4AwHQYDVR0O
# BBYEFBUH0gjHbvjES8JvdKfUpJhtcsFBMB8GA1UdIwQYMBaAFHTXdBHMIeijds3+
# 3jGs7lntQNflMIHBBgNVHR8EgbkwgbYwgbOggbCgga2GgapsZGFwOi8vL0NOPWFk
# bS1EQzItQ0EwMSxDTj1kYzIsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZp
# Y2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9aGwsREM9bG9jP2Nl
# cnRpZmljYXRlUmV2b2NhdGlvbkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0
# cmlidXRpb25Qb2ludDCBuQYIKwYBBQUHAQEEgawwgakwgaYGCCsGAQUFBzAChoGZ
# bGRhcDovLy9DTj1hZG0tREMyLUNBMDEsQ049QUlBLENOPVB1YmxpYyUyMEtleSUy
# MFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9aGwsREM9
# bG9jP2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9u
# QXV0aG9yaXR5MDUGA1UdEQQuMCygKgYKKwYBBAGCNxQCA6AcDBpzLmtvdmFsZXYu
# YWRtaW5AYWRtLmhsLmxvYzBPBgkrBgEEAYI3GQIEQjBAoD4GCisGAQQBgjcZAgGg
# MAQuUy0xLTUtMjEtMzk4MDA3NzYzMy0zNDYxMDUxOTcwLTI0MTQ0NzE2ODYtMTE2
# NzANBgkqhkiG9w0BAQsFAAOCAQEAg/+Oa95LIWgbux7dLe6PosSYgXWh6wSKL1Rs
# R8NzFTkj5IOjUnuDVB2Sos59STTmIGAPnH4W7jUZSQ3AgKPqp9ouHPFIlyJIl1IQ
# YlejwVfPIn1z6Wf0GaKGGSkqJOpEPecYyBuhGg47NRMr8hOXcNfiuTWGwQN6L+45
# TVhBOE3pMBQv7w1kyLxttT2HhAFec5uWNKyQXl12PLwbMpwC9fwpgChFSK/JCU6e
# pJSYec+9ut2w/klOakkg4MoanwH1I+4WIQSwL/3Wmf53FXuJcG0apE37V2nc7LLT
# wzhl35tdlvFh2vY5IFDHK8yuHkxWxh41XI5n/cglS+fTGFFyaDGCAg0wggIJAgEB
# MGwwVTETMBEGCgmSJomT8ixkARkWA2xvYzESMBAGCgmSJomT8ixkARkWAmhsMRMw
# EQYKCZImiZPyLGQBGRYDYWRtMRUwEwYDVQQDEwxhZG0tREMyLUNBMDECEywAAAAc
# 5YtMKMFp02MAAAAAABwwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKA
# AKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFEvh0GSez6ltqdiwKOgPWPdj
# aQQbMA0GCSqGSIb3DQEBAQUABIIBAAQd7deotxiXx4Q2iYIiZbl5BJenJWAsTPQe
# 8+TYPsDTc1QmEFs4hQ5DhcEPrFRVkWyy4W8eMTexRtXKQlbqiGex9+7MvFgD8UgM
# wjylm87qZgcvQoPKycWSoysGSxxG3Na8Lhj+HdWlSfOsjknG6F2U2ES6XtcWS7PU
# uynCOkLCjX9WJqwbJ90LEEgXZJKG/BZNW6ErRc7jXzu6kb4b6mUZJjjyMT5njV7W
# TeLNHwvnZgaD1cEMpySNEbXpmMXd4Zm9MaqbrQ1ijB/gtF4dprfgxRP5xGSe7beR
# Yeb/iH0SW53H5/D8XTcNgY9eKnwpRRxwedZ3I9NnxpDx5vnpdOw=
# SIG # End signature block
