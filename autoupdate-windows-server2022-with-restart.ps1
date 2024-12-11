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
    $SearchResults = $UpdateSearcher.Search("IsInstalled=0 and Type='Software'")

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
