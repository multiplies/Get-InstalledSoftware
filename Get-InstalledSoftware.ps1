# this need to be runned on a DC or pc with RSAT tools installed and DC module installed

# functie to get installed software from domain pc's
function Get-InstalledSoftware(){
    
    # regkeys 32 and 64bit of installed software
    $UninstallRegKeys=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")

    # getting all computers in a domain
    [string[]]$ComputerName = (get-adcomputer -Filter *).Name
    
    # creating empty arrays
    $installedSoftware = @()
    $installedSoftwarePerPC = @()

    # making int var for progress bar
    $i = 1

    # looping through all pc's in the domain
    foreach($Computer in $ComputerName) {
        # making progress bar and updating it after each pc 
        Write-Progress -Activity "Connecting to AD Computers" -Status 'Connecting' -PercentComplete ($i/$ComputerName.Count*100) -CurrentOperation $Computer;
        # increassing i counter
        $i++
        # creating empty array
        $pcApps = @()
        # writing o log file            
        LogWrite "Trying to connect to pc: $Computer"
        # testing if pc is online            
        if(Test-Connection -ComputerName $Computer -Count 1 -ErrorAction SilentlyContinue) { # -ea 0
            # writing to log file working on the pc
            LogWrite "Working on pc $Computer"
            # looping through the regkeys           
            foreach($UninstallRegKey in $UninstallRegKeys) {
                # trying to get all installed apps from the remote pc            
                try {
                    $Applications = Get-ItemProperty -Path $UninstallRegKey 
                # catch if there is a error and write it to logfile          
                } catch {            
                    LogWrite "Failed to read $UninstallRegKey"            
                    Continue            
                } 
            # making int var for progress bar          
            $j = 1
            #looping through all apps
            foreach ($App in $Applications) {
                # putting all info in seperated vars       
               $AppDisplayName  = $($App.DisplayName)            
               $AppVersion   = $($App.DisplayVersion)            
               $AppPublisher  = $($App.Publisher)            
               $AppInstalledDate = $($App.InstallDate)            
               $AppUninstall  = $($App.UninstallString)
               # making new progresbar en updating it after each run
               Write-Progress -Activity "Collecting installed Apps" -Status 'Collecting App info' -PercentComplete ($j/$Applications.Count*100) -CurrentOperation $AppDisplayName; 
               # increasing the counter 
               $j++ 
               # checking if Softwarearchitecture is 32 or 64 bit         
               if($UninstallRegKey -match "Wow6432Node") {            
                    $Softwarearchitecture = "x86"            
               } else {            
                    $Softwarearchitecture = "x64"            
               }
               # if there is a ppDisplyName put all info in a PSObj and add it to the installedSoftware array          
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
    # return the installedSoftware
    return $installedSoftware 
}

# functie to write to a log file
Function LogWrite
{
   # reading string param
   Param ([string]$logstring)

   # writing content to logfile
   Add-content $Logfile -value $logstring
}

# creating the default path from where the script is being runt
$dir = (pwd).Path
# making a timestamp
$timestamp = Get-Date -Format o | foreach {$_ -replace ":", "."}

#creating the logfile
$Logfile = "$dir\installedSoftware_$timestamp.log"

#getting all installed software from the function Get-InstalledSoftware
$soft = Get-InstalledSoftware

# writing to log file
LogWrite "Managing the data and exporting it to make a csv file"

# creating the finalearray here will everthing be stored to output the info
$FinaleArray = @()
# looping through all the installed software
$soft | foreach {
    # creating 2 vars AppName and AppVersion
    $AppName = $_.AppName
    $AppVersion = $_.AppVersion
    # checking if there is already a PSObj with the same AppName and AppVersion
    if($FinaleArray.AppName -eq $_.AppName -and $FinaleArray.AppVersion -eq $_.AppVersion){
        # if there is already a PSObj with the same AppName and AppVersion than ad 1 to the count and add the pc name to the list
        # making a temp var to get later the index of the PSObj in the array
        $temp = $FinaleArray | where { $AppName -eq $_.AppName -and $AppVersion -eq $_.AppVersion }
        # finding the index of the PSObj
        $index = $FinaleArray.indexof($temp)
        # adding 1 to the count and the pc name to the PSObj
        $FinaleArray[$index].Count = $FinaleArray[$index].Count+1
        $FinaleArray[$index].ComputerName = $FinaleArray[$index].ComputerName + " " + $_.ComputerName
    }else{
        # if there is not already a PSObj with the same AppName and AppVersion than make the PSObj and add it to the array
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

# writing info to the log file
LogWrite "Writing the XML file to $dir\installedSoftware_$timestamp.csv"
# creating the XML file
$FinaleArray | Export-Clixml "$dir\installedSoftware_$timestamp.xml" -NoClobber #-Delimiter ',' -NoTypeInformation #| Format-Table -AutoSize 
# making the gridview to see all installed software
$FinaleArray | Out-GridView

# LogWrite "$Error[0]"
