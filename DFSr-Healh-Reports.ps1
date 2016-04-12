#-----------------------------------------------------------------------
$To = 'Email@SomeWhere.com'
$From = 'Email@SomeWhere.com'
$Subject = "DFS Health Reports - $((get-date).ToString('MM/dd/yyyy'))"
$MailServer = 'mail.SomeWhere.com'
#-----------------------------------------------------------------------
$File_Path = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$Reports_Path = "$File_Path\Health_Reports"
$TempFile = "$File_Path\DFSReport.txt"
$SummaryFile = "$File_Path\Summary\DFS_Summary_Report_$(get-date -Format MMddyyyy).csv"
If (Test-Path $TempFile) {Remove-Item $TempFile -Force}
If (-not(Test-Path -path "$File_Path\Health_Reports")) {New-Item "$File_Path\Health_Reports" -Type Directory | Out-Null}
If (-not(Test-Path -path "$File_Path\Summary")) {New-Item "$File_Path\Summary" -Type Directory | Out-Null}
$Domains = 'Domain1','Domain2'
Get-ChildItem -Path "$File_Path\Summary" | Where {$_.PSisContainer -eq $false -and $_.LastWriteTime -lt (Get-date).AddDays(-5)} | Remove-Item -Force
Get-ChildItem -path $Reports_Path | Remove-Item -Force
Foreach ($Domain in $Domains){
$RepGroupNames = dfsradmin rg list /Domain:$Domain /attr:rgname /CSV | Where {-not($_ -like 'RgName' -or $_ -like "")} | Foreach {
    If ($_ -eq 'Domain System Volume'){
            $HTML_Path = """$Reports_Path\$($($Domain.split("."))[0])-SYSVOL"""
        }#End of If
        Else{
            $HTML_Path = """$Reports_Path\$_"""
            }#End of Else
        $RepGroup = """$_"""
        while (@(Get-Job -State Running).Count -ge 75) {Start-Sleep -Seconds 2}
            Start-Job -Name "DFSrHealth-$RepGroup" -ScriptBlock {param($Domain,$RepGroup,$HTML_Path)
            &cmd /c "DfsrAdmin.exe Health New /Domain:$Domain /RgName:$RepGroup /RepName:$HTML_Path /FsCount:false"
        } -ArgumentList $Domain,$RepGroup,$HTML_Path
}#End of Foreach RepGroup
}#End of Foreach Domain
Get-Job | Wait-Job -Timeout 5400
Get-ChildItem -path $Reports_Path *.xml | Remove-Item -Force
$HTMLFiles = Get-ChildItem -path $Reports_Path *.html
$OutPut = @()
$TotalWarningCount = 0
$TotalErrorCount = 0
$ErrorBody = "The following RepGroups have Errors:`n"
$WarningBody = "The following RepGroups have Warnings:`n"
Foreach ($HTMLFile in $HTMLFiles){
        Get-Process -Name iexplore -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        sleep -Milliseconds 75
        $ie = new-object -com "InternetExplorer.Application"
        sleep -Milliseconds 75
        $ie.navigate($HTMLFile.FullName)
        sleep -Milliseconds 75
        $ie.Document.body.outerText | Out-File $TempFile
        sleep -Milliseconds 75
        $ie.Quit()
        (Select-String -Pattern "\w" -Path $TempFile) | ForEach-Object { $_.line.Trim(" ") -Replace '\s+',' ' } | Set-Content -Path $TempFile
        $Lines = Get-Content $TempFile
        Foreach ($Line in $Lines){
            switch -WildCard ($Line){
                "ERRORS (*server*with*error*)"{
                    If ($Line -match "[1-9]\d*"){
                        $ErrorCount = $Matches[0]
                        $TotalErrorCount ++
                        $FoundError = $True
                    }#End of If
                }#End of ERRORS
                "WARNINGS (*server*with*warning*)*"{
                    If ($Line -match "[1-9]\d*"){
                        $WarningCount = $Matches[0]
                        $TotalWarningCount ++
                        $FoundWarning = $True 
                    }#End of If
                }#End of WARNINGS
                "Description:*"{
                $FoundDescription += "$Line"
                }#End of Description
                "Last occurred:*"{
                $FoundOccurred += "$Line"
                }#End of Last occurred
            }# End of Switch
        
        }# End of Foreach Line
        If ($FoundError -or $FoundWarning){
            If ($FoundError){
                $ErrorBody += "`t$($($HTMLFile.Name) -replace '.html','')`n"
            }
            Elseif ($FoundWarning){
                $WarningBody += "`t$($($HTMLFile.Name) -replace '.html','')`n"
            }
            
            If (($FoundDescription -split "Description:").Count -gt 1){
                $Count = 1
                Foreach ($Desc in ($FoundDescription -split "Description:" | Where {$_ -notlike ""})){
                    $SplitFoundOccurred = $FoundOccurred -split "Last occurred:"
                        $Data = New-Object psobject
                        $Data | Add-Member -MemberType "noteproperty" -Name 'Replication Group' -Value $($($HTMLFile.Name) -replace '.html','')
                        $Data | Add-Member -MemberType "noteproperty" -Name 'Errors' -Value $ErrorCount
                        $Data | Add-Member -MemberType "noteproperty" -Name 'Warnings' -Value $WarningCount
                        $Data | Add-Member -MemberType "noteproperty" -Name 'Description' -Value $Desc
                        $Data | Add-Member -MemberType "noteproperty" -Name 'Last Occurred' -Value $SplitFoundOccurred[$Count]
                        $Data | Add-Member -MemberType "noteproperty" -Name 'Path' -Value $($($HTMLFile.FullName).Replace('C:\',"\\$(get-content env:computername)\$($($HTMLFile.FullName).substring(0,1))$\"))
                    $Count++
                    $OutPut += $Data
                }
            }
            Else{$Data = New-Object psobject
                 $Data | Add-Member -MemberType "noteproperty" -Name 'Replication Group' -Value $($($HTMLFile.Name) -replace '.html','')
                 $Data | Add-Member -MemberType "noteproperty" -Name 'Errors' -Value $ErrorCount
                 $Data | Add-Member -MemberType "noteproperty" -Name 'Warnings' -Value $WarningCount
                 $Data | Add-Member -MemberType "noteproperty" -Name 'Description' -Value $FoundDescription
                 $Data | Add-Member -MemberType "noteproperty" -Name 'Last Occurred' -Value $FoundOccurred
                 $Data | Add-Member -MemberType "noteproperty" -Name 'Path' -Value $($($HTMLFile.FullName).Replace('C:\',"\\$(get-content env:computername)\$($($HTMLFile.FullName).substring(0,1))$\"))
                 $OutPut += $Data
            }
            #
            $WarningCount = 0
            $ErrorCount = 0
            $FoundError = $False
            $FoundWarning = $False
            $FoundDescription = $Null
            $FoundOccurred = $Null
        }
        Else{
            $WarningCount = 0
            $ErrorCount = 0
            $FoundError = $False
            $FoundWarning = $False
            $FoundDescription = $Null
            $FoundOccurred = $Null
        }
}#End of foreach HTMLFile
$OutPut | Export-Csv -Path "$SummaryFile" -NoTypeInformation
If ((Get-Date).DayOfWeek -eq 'Monday' -or (Get-Date).DayOfWeek -eq 'Friday'){
    $EmailBody = "Total RepGroups with errors = $TotalErrorCount`nTotal RepGroups with warnings = $TotalWarningCount`n`nMore detail can be found here:`n$($SummaryFile.Replace('C:\',"\\$(get-content env:computername)\$($SummaryFile.substring(0,1))$\"))`n`n$ErrorBody`n$WarningBody`n`n"
    Send-MailMessage -to $To -from $From -Subject $Subject -Body $EmailBody -SmtpServer $MailServer
}
Get-Job | Remove-Job -Force