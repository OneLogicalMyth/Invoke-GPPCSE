Function Get-GPPData {
    param([System.IO.FileInfo]$GPPFile)

    [xml]$GPPXML = Get-Content $GPPFile.FullName
    $PotentialPasswords = @()

    switch ($GPPFile.Name) {

        'Groups.xml' {
            if($GPPXML.Groups.User)
            {
                foreach($User in $GPPXML.Groups.User)
                {
                    $Pass = '' | Select-Object Username, NewName, cPassword, Changed
                    $Pass.Username = $User.Properties.userName
                    $Pass.NewName = $User.Properties.newName
                    $Pass.cPassword = $User.Properties.cpassword
                    $Pass.Changed = [datetime]$User.changed
                    $PotentialPasswords += $Pass
                    Remove-Variable -Name Pass
                }
                Remove-Variable -Name User
            }
        }
        
        'Services.xml' {  
        $Cpassword += , $Xml | Select-Xml "/NTServices/NTService/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $UserName += , $Xml | Select-Xml "/NTServices/NTService/Properties/@accountName" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $Changed += , $Xml | Select-Xml "/NTServices/NTService/@changed" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        }
        
        'Scheduledtasks.xml' {
        $Cpassword += , $Xml | Select-Xml "/ScheduledTasks/Task/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $UserName += , $Xml | Select-Xml "/ScheduledTasks/Task/Properties/@runAs" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $Changed += , $Xml | Select-Xml "/ScheduledTasks/Task/@changed" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        }
        
        'DataSources.xml' { 
        $Cpassword += , $Xml | Select-Xml "/DataSources/DataSource/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $UserName += , $Xml | Select-Xml "/DataSources/DataSource/Properties/@username" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $Changed += , $Xml | Select-Xml "/DataSources/DataSource/@changed" | Select-Object -Expand Node | ForEach-Object {$_.Value}                          
        }
                    
        'Printers.xml' { 
        $Cpassword += , $Xml | Select-Xml "/Printers/SharedPrinter/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $UserName += , $Xml | Select-Xml "/Printers/SharedPrinter/Properties/@username" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $Changed += , $Xml | Select-Xml "/Printers/SharedPrinter/@changed" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        }
  
        'Drives.xml' { 
        $Cpassword += , $Xml | Select-Xml "/Drives/Drive/Properties/@cpassword" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $UserName += , $Xml | Select-Xml "/Drives/Drive/Properties/@username" | Select-Object -Expand Node | ForEach-Object {$_.Value}
        $Changed += , $Xml | Select-Xml "/Drives/Drive/@changed" | Select-Object -Expand Node | ForEach-Object {$_.Value} 
        }
    }

    return $PotentialPasswords

}
