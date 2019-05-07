Function Get-GPPData {
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [System.IO.FileInfo] $GPPFile
        )

    [xml]$GPPXML = Get-Content $GPPFile.FullName

    switch ($GPPFile.Name) {

        'Groups.xml' {
            if($GPPXML.Groups.User)
            {
                foreach($User in $GPPXML.Groups.User)
                {
                    $Pass = '' | Select-Object Type, Username, NewName, cPassword, Changed
                    $Pass.Type = 'Groups'
                    $Pass.Username = $User.Properties.userName
                    $Pass.NewName = $User.Properties.newName
                    $Pass.cPassword = $User.Properties.cpassword
                    $Pass.Changed = [datetime]$User.changed
                    $Pass
                    Remove-Variable -Name Pass
                }
                Remove-Variable -Name User
            }
        }
        
        'Services.xml' {
            foreach($Item in ($GPPXML | Select-Xml "/NTServices/NTService/Properties" | Select -ExpandProperty Node))
            {
                    $Pass = '' | Select-Object Type, Username, NewName, cPassword, Changed
                    $Pass.Type = 'Services'
                    $Pass.Username = $Item.accountName
                    $Pass.NewName = 'n/a'
                    $Pass.cPassword = $Item.cpassword
                    $Pass.Changed = [datetime]$Item.ParentNode.changed
                    $Pass
                    Remove-Variable -Name Pass
            }
        }
        
        'Scheduledtasks.xml' {
            foreach($Item in ($GPPXML | Select-Xml "/ScheduledTasks/Task/Properties" | Select -ExpandProperty Node))
            {
                    $Pass = '' | Select-Object Type, Username, NewName, cPassword, Changed
                    $Pass.Type = 'Scheduledtasks'
                    $Pass.Username = $Item.runAs
                    $Pass.NewName = 'n/a'
                    $Pass.cPassword = $Item.cpassword
                    $Pass.Changed = [datetime]$Item.ParentNode.changed
                    $Pass
                    Remove-Variable -Name Pass
            }
        }
        
        'DataSources.xml' {
            foreach($Item in ($GPPXML | Select-Xml "/DataSources/DataSource/Properties" | Select -ExpandProperty Node))
            {
                    $Pass = '' | Select-Object Type, Username, NewName, cPassword, Changed
                    $Pass.Type = 'DataSources'
                    $Pass.Username = $Item.username
                    $Pass.NewName = 'n/a'
                    $Pass.cPassword = $Item.cpassword
                    $Pass.Changed = [datetime]$Item.ParentNode.changed
                    $Pass
                    Remove-Variable -Name Pass
            }                      
        }
  
        'Drives.xml' {
            foreach($Item in ($GPPXML | Select-Xml "/Drives/Drive/Properties" | Select -ExpandProperty Node))
            {
                    $Pass = '' | Select-Object Type, Username, NewName, cPassword, Changed
                    $Pass.Type = 'Drives'
                    $Pass.Username = $Item.userName
                    $Pass.NewName = 'n/a'
                    $Pass.cPassword = $Item.cpassword
                    $Pass.Changed = [datetime]$Item.ParentNode.changed
                    $Pass
                    Remove-Variable -Name Pass
            }
        }
    }

}
