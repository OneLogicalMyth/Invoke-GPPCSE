Function Get-GPPData {
    [CmdletBinding()]
    param(
        [System.IO.FileInfo] $GPPFile
        )

    [xml]$GPPXML = Get-Content $GPPFile.FullName

    switch ($GPPFile.Name) {

        'Groups.xml' {
            if($GPPXML.Groups.User)
            {
                foreach($User in $GPPXML.Groups.User)
                {
                    [pscustomobject]@{
                        Type = 'Groups'
                        Username = $User.Properties.userName
                        NewName = $User.Properties.newName
                        cPassword = $User.Properties.cpassword
                        Changed = [datetime]$User.changed
                    }

                }
                Remove-Variable -Name User
            }
        }
        
        'Services.xml' {
            foreach($Item in ($GPPXML | Select-Xml "/NTServices/NTService/Properties" | Select -ExpandProperty Node))
            {
                    [pscustomobject]@{
                        Type = 'Services'
                        Username = $Item.accountName
                        NewName = 'n/a'
                        cPassword = $Item.cpassword
                        Changed = [datetime]$Item.ParentNode.changed
                    }
            }
        }
        
        'Scheduledtasks.xml' {
            foreach($Item in ($GPPXML | Select-Xml "/ScheduledTasks/Task/Properties" | Select -ExpandProperty Node))
            {
                    [pscustomobject]@{
                        Type = 'Scheduledtasks'
                        Username = $Item.runAs
                        NewName = 'n/a'
                        cPassword = $Item.cpassword
                        Changed = [datetime]$Item.ParentNode.changed
                    }
            }
        }
        
        'DataSources.xml' {
            foreach($Item in ($GPPXML | Select-Xml "/DataSources/DataSource/Properties" | Select -ExpandProperty Node))
            {
                    [pscustomobject]@{
                        Type = 'DataSources'
                        Username = $Item.username
                        NewName = 'n/a'
                        cPassword = $Item.cpassword
                        Changed = [datetime]$Item.ParentNode.changed
                    }
            }                      
        }
  
        'Drives.xml' {
            foreach($Item in ($GPPXML | Select-Xml "/Drives/Drive/Properties" | Select -ExpandProperty Node))
            {
                    [pscustomobject]@{
                        Type = 'Drives'
                        Username = $Item.userName
                        NewName = 'n/a'
                        cPassword = $Item.cpassword
                        Changed = [datetime]$Item.ParentNode.changed
                    }
            }
        }
    }

}
