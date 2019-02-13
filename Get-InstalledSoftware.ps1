function Get-InstalledSoftware(){

    $UninstallRegKeys=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")

    [string[]]$ComputerName = (get-adcomputer -Filter *).Name
    
    $installedSoftware = @() 

    $installedSoftwarePerPC = @()
    $i = 1

    foreach($Computer in $ComputerName) {
        Write-Progress -Activity "Connecting to AD Computers" -Status 'Connecting' -PercentComplete ($i/$ComputerName.Count*100) -CurrentOperation $Computer;
        $i++
        $pcApps = @()            
        LogWrite "Trying to connect to pc: $Computer"            
        if(Test-Connection -ComputerName $Computer -Count 1 -ErrorAction SilentlyContinue) { # -ea 0
            LogWrite "Working on pc $Computer"           
            foreach($UninstallRegKey in $UninstallRegKeys) {            
                try {
                    $Applications = Get-ItemProperty -Path $UninstallRegKey           
                } catch {            
                    LogWrite "Failed to read $UninstallRegKey"            
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
               }            
            }             
        }          
    }
    
    return $installedSoftware 
}

Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

$dir = (pwd).Path
$timestamp = Get-Date -Format o | foreach {$_ -replace ":", "."}

$Logfile = "$dir\installedSoftware_$timestamp.log"

$soft = Get-InstalledSoftware

LogWrite "Managing the data and exporting it to make a csv file"

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
LogWrite "Writing the CSV file to $dir\installedSoftware_$timestamp.csv"
$FinaleArray | Export-Csv "$dir\installedSoftware_$timestamp.csv" -Delimiter ',' -NoTypeInformation #| Format-Table -AutoSize 
LogWrite "$Error[0]"
