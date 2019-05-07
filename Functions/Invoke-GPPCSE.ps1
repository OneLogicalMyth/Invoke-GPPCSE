Function Invoke-GPPCSE {
    [CmdletBinding()]
    param(
        [string] $DomainController = $env:USERDNSDOMAIN,
        [string] $Username = $null,
        [string] $Password = $null
        )


    begin
    {
        # authenticate        
        if($DomainController -and ([string]::IsNullOrEmpty($Username) -OR [string]::IsNullOrEmpty($Password))){
            # assume a remote domain that the current logon credentials have access to
            $DirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainController")
            $DomainType = 'RemoteTrusted'
        }elseif([string]::IsNullOrEmpty($Username) -OR [string]::IsNullOrEmpty($Password)){
            # bind to the current domain
            $DirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry
            $DomainType = 'Local'
        }else{
            # connect to a remote domain with credentials
            $DirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainController",$Username,$Password)
            $DomainType = 'ExternalDomain'
        }

        # build LDAP filter
        $CSEGUIDs = Get-CSEGUIDs
        $LDAPFilter = '(|'
        foreach($GUID in $CSEGUIDs.Keys){ $LDAPFilter += "(gPCMachineExtensionNames=*$GUID*)" }
        $LDAPFilter += ')'

        # Build searcher
        $Searcher             = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot  = $DirectoryEntry
        $Searcher.Filter      = $LDAPFilter

        # map network share to domain SYSVOL
        if($DomainType -eq 'ExternalDomain')
        {
            # passwords with double quotes in them get escaped
            $null = Invoke-Expression ('net use \\' + $DomainController + '\SYSVOL /USER:' + $Username + ' "' + ($Password -Replace '"','\"') + '"')
        }else{
            $null = Invoke-Expression ('net use \\' + $DomainController + '\SYSVOL')
        }

        # construct SYSVOL string for policies
        $DomainName = $DirectoryEntry.Properties['distinguishedName'] -replace ',DC=','.' -replace 'DC='
        $PolicyPath = "\\$DomainController\sysvol\$DomainName\Policies"

        # Look for GPOs that have a GPP CSE
        $GPOs = $Searcher.FindAll()

        # if no GPOs found then error and exit no need to process an empty list :)
        if($GPOs -eq $null)
        {
            throw 'No GPOs containing GPP found'
            return
        }

    }

    process
    {

        foreach($GPO in $GPOs)
        {
            # first workout if any files exist and if it has any passwords to save more network calls
            $GPOPath = Join-Path $PolicyPath $GPO.Properties['name'][0]
            $Files = $CSEGUIDs.Values | foreach{
                # some GPP preferences apply to both user and computer
                foreach($CSEFile in $_)
                {
                    Write-Verbose "Searching file $(Join-Path $GPOPath $CSEFile)"
                    Get-Item (Join-Path $GPOPath $CSEFile) -ErrorAction SilentlyContinue
                }
            }

            if(-not $Files)
            {
                # no files found, skip this GPO
                Write-Verbose "$($GPO.Properties.name[0]) has no GPP files"
                return
            }
            
            Function Get-GPOFlag {
            param([int]$flag)

                switch($flag)
                {
                    3 { 'All Settings Disabled' }
                    2 { 'Computer configuration settings disabled' }
                    0 { 'Enabled' }
                    1 { 'User configuration settings disabled' }
                }

            }


            $Passwords = $Files | foreach { Get-GPPData -GPPFile $_ }

            if($Passwords)
            {
                foreach($Pass in $Passwords)
                {
                    
                    $Result = '' | Select-Object GPPType, GPPUsername, GPPNewUsername, GPPPassword, GPOPath, GPOName, GPOGUID, GPOLinkedOUs, GPOAssociatedOUs, GPOStatus, RelatedActiveComputers
                    $Result.GPPType = $Pass.Type
                    $Result.GPPUsername = $Pass.Username
                    $Result.GPPNewUsername = $Pass.NewName
                    $Result.GPPPassword = $(Get-DecryptedCpassword $Pass.cPassword)
                    $Result.GPOPath = $GPOPath
                    $Result.GPOName = $GPO.Properties['displayname'][0]
                    $Result.GPOGUID = $GPO.Properties['name'][0]
                    $Result.GPOLinkedOUs = $(Get-GPOOUs -GUID $GPO.Properties['name'] -ReturnDNOnly -LinkedOnly -DirectoryEntry $DirectoryEntry)
                    $Result.GPOAssociatedOUs = $(Get-GPOOUs -GUID $GPO.Properties['name'] -ReturnDNOnly -DirectoryEntry $DirectoryEntry)
                    $Result.GPOStatus = Get-GPOFlag $GPO.Properties['flags'][0]
                    $Result.RelatedActiveComputers = $(Get-GPOOUs -GUID $GPO.Properties['name'] -DirectoryEntry $DirectoryEntry | %{ Get-ADComputers -DirectoryEntry $_ })
                    $Result
                }
            }



        }

    }

}