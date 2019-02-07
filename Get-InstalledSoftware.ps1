function Get-InstalledSoftware(){

    $UninstallRegKeys=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")

    [string[]]$ComputerName = (get-adcomputer -Filter *).Name
    
    $installedSoftware = @() 

    $installedSoftwarePerPC = @()
    $i = 1

    foreach($Computer in $ComputerName) {
        Write-Progress -Activity "Connecting to Ad Computers" -Status 'Connecting' -PercentComplete ($i/$ComputerName.Count*100) -CurrentOperation $Computer;
        $i++
        $pcApps = @()            
        Write-Verbose "Trying to connect to pc: $Computer"            
        if(Test-Connection -ComputerName $Computer -Count 1 -ErrorAction SilentlyContinue) { # -ea 0
            Write-Verbose "Working on pc $Computer"           
            foreach($UninstallRegKey in $UninstallRegKeys) {            
                try {
                    $Applications = Get-ItemProperty -Path $UninstallRegKey           
                } catch {            
                    Write-Verbose "Failed to read $UninstallRegKey"            
                    Continue            
                }           
            $j = 1
            foreach ($App in $Applications) {       
               $AppDisplayName  = $($App.DisplayName)            
               $AppVersion   = $($App.DisplayVersion)            
               $AppPublisher  = $($App.Publisher)            
               $AppInstalledDate = $($App.InstallDate)            
               $AppUninstall  = $($App.UninstallString)
               Write-Progress -Activity "Collecting installed Apps" -Status 'Collecting App info' -PercentComplete ($j/$Applications.Count*100) -CurrentOperation $AppDisplayName;  
               $j++          
               if($UninstallRegKey -match "Wow6432Node") {            
                    $Softwarearchitecture = "x86"            
               } else {            
                    $Softwarearchitecture = "x64"            
               }            
               if(!$AppDisplayName) { continue }        
                   $OutputObj = New-Object -TypeName PSobject
                   $OutputObj | Add-Member -MemberType NoteProperty -Name AppName -Value $AppDisplayName            
                   $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()                
                   $OutputObj | Add-Member -MemberType NoteProperty -Name AppVersion -Value $AppVersion            
                   $OutputObj | Add-Member -MemberType NoteProperty -Name AppVendor -Value $AppPublisher            
                   $OutputObj | Add-Member -MemberType NoteProperty -Name InstalledDate -Value $AppInstalledDate            
                   $OutputObj | Add-Member -MemberType NoteProperty -Name UninstallKey -Value $AppUninstall          
                   $OutputObj | Add-Member -MemberType NoteProperty -Name SoftwareArchitecture -Value $Softwarearchitecture            
                   $installedSoftware += $OutputObj
                   #$pcApps += $OutputObj                          
               }            
            }             
        }
        #$installedSoftwarePerPC += $pcApps            
    }
    
    return $installedSoftware 
}
$dir = (pwd).Path
$timestamp = Get-Date -Format o | foreach {$_ -replace ":", "."}

#Get-InstalledSoftware -verbose | Sort-Object -Property AppName -Unique | Format-Table -AutoSize

$soft = Get-InstalledSoftware -verbose

#$soft | Sort-Object -Property AppName -Unique | select AppName, @{label='count';Expression={($soft.AppName -eq $_.AppName).Count}}, @{label='ComputerName';expression={($soft.AppName -eq $_.AppName | select $_.ComputerName) | Out-String}}, @{label='AppVersion';expression={($soft.AppName -eq $_.AppName | select $_.AppVersion)}} | Export-Csv "$dir\installedSoftware_$timestamp.csv" -Delimiter ',' -NoTypeInformation

#Get-InstalledSoftware | select AppName | sort AppName | Get-Unique -OnType | sort AppName | measure AppName | select AppName, AppVendor, InstalledDate, ComputerName | Export-Csv "$dir\installedSoftware_$timestamp.csv" -Delimiter ',' -NoTypeInformation

$FinaleArray = @()
$installedSoftwarePerPC | foreach {
    $AppName = $_.AppName
    $AppVersion = $_.AppVersion
    if($FinaleArray.AppName -eq $_.AppName -and $FinaleArray.AppVersion -eq $_.AppVersion){
        $temp = $FinaleArray | where { $AppName -eq $_.AppName -and $AppVersion -eq $_.AppVersion }
        $index = $FinaleArray.indexof($temp)
        $FinaleArray[$index].Count = $FinaleArray[$index].Count+1
        $FinaleArray[$index].ComputerName = $FinaleArray[$index].ComputerName + "`n" + $_.ComputerName
    }else{
        $OutputObj = New-Object -TypeName PSobject
        $OutputObj | Add-Member -MemberType NoteProperty -Name AppName -Value $_.AppName
        $OutPutObj | Add-Member -MemberType NoteProperty -Name Count -Value 1            
        $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $_.ComputerName                
        $OutputObj | Add-Member -MemberType NoteProperty -Name AppVersion -Value $_.AppVersion            
        $OutputObj | Add-Member -MemberType NoteProperty -Name AppVendor -Value $_.AppVendor            
        $OutputObj | Add-Member -MemberType NoteProperty -Name InstalledDate -Value $_.InstalledDate            
        $OutputObj | Add-Member -MemberType NoteProperty -Name UninstallKey -Value $_.UninstallKey         
        $OutputObj | Add-Member -MemberType NoteProperty -Name SoftwareArchitecture -Value $_.SoftwareArchitecture
        $FinaleArray += $OutputObj    
    }
}
$FinaleArray | Export-Csv "$dir\installedSoftware_$timestamp.csv" -Delimiter ',' -NoTypeInformation #| Format-Table -AutoSize 
