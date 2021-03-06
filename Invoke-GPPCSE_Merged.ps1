function Get-ADComputers {
    [cmdletbinding()]
    param(
        [int] $ActiveDays = 60,
        [string] $DomainController = $env:USERDNSDOMAIN,
        [string] $Username,
        [string] $Password,
        [System.DirectoryServices.DirectoryEntry] $DirectoryEntry = $null
        )

    $ActiveDaysFileTime = (Get-Date).AddDays(-$ActiveDays).ToFileTimeUtc()
    $EnabledComputersFilter = '(&(objectCategory=Computer)(lastLogonTimestamp>='
    $EnabledComputersFilter += $ActiveDaysFileTime
    $EnabledComputersFilter += ')(operatingsystem=*)(dNSHostName=*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))'
    
    if([string]::IsNullOrEmpty($Username) -OR [string]::IsNullOrEmpty($Password)){
        if(-not $DirectoryEntry)
        {
            $DirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry
        }
    }else{
        $DirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainController",$Username,$Password)
    }
    
    $PropertyList = @(
	    'dNSHostName'
        'operatingsystem'
        'lastLogonTimestamp'
        'whenCreated'
        'whenChanged'
        'ServicePrincipalName'
    )

    $Searcher             = New-Object System.DirectoryServices.DirectorySearcher
    $Searcher.SearchRoot  = $DirectoryEntry
    $Searcher.PageSize    = 500
    $Searcher.Filter      = $EnabledComputersFilter
    $Searcher.SearchScope = "Subtree"

    foreach($Property IN $PropertyList){
        $Searcher.PropertiesToLoad.Add($Property) | Out-Null
    }

    $Results = $Searcher.FindAll()

    Function Get-PropertyValue {
        param($Properties,$ItemName)
        if(($Properties.PropertyNames -contains $ItemName)){
            return $Properties.$ItemName[0]
        }else{
            return $null
        }
    }

    $Results | foreach{

        $Out = '' | Select-Object DNSHostName, OperatingSystem, LastLogonTimeStamp, WhenCreated, WhenChanged, ServicePrincipalName, ComputerType   
        $Out.DNSHostName            = Get-PropertyValue $_.Properties dnshostname
        $Out.OperatingSystem        = Get-PropertyValue $_.Properties operatingsystem
        $Out.LastLogonTimeStamp     = [datetime]::FromFileTimeUtc((Get-PropertyValue $_.Properties lastlogontimestamp))
        $Out.WhenCreated            = Get-PropertyValue $_.Properties whencreated
        $Out.WhenChanged            = Get-PropertyValue $_.Properties whenchanged
        $Out.ServicePrincipalName   = Get-PropertyValue $_.Properties serviceprincipalname
        $Out.ComputerType           = Test-ComputerType -SPN $Out.ServicePrincipalName -OS $Out.OperatingSystem
        $Out

    }
}

Function Test-ComputerType {
    [cmdletbinding()]
    Param(
        [string] $SPN,
        [string] $OS
        )

	# Check if cluster virtual name
	if(($SPN) -like '*MSClusterVirtualServer*')
	{
		return 'Cluster'
	}

	# Check if OS has the word server in it
	if($OS -like "*server*")
	{
		return 'Server'
	}

	# Check if its a client device
	$patterns = @('Windows 8*','Windows 7*','Windows XP*','Windows Embedded*','Windows 2000 Professional*','Windows Vista*','Windows 10*')
	foreach($pattern in $patterns) { if($OS -like $pattern) { return 'ClientDevice'; } }

	# no other matches found
	return 'Unknown'

}
Function Get-CSEGUIDs {

    @{
    '{5794DAFD-BE60-433f-88A2-1A31939AC01F}'=@('\User\Preferences\Drives\Drives.xml')
    '{17D89FEC-5C44-4972-B12D-241CAEF74509}'=@('\Machine\Preferences\Groups\Groups.xml','\User\Preferences\Groups\Groups.xml') # user preference can set a admin password ?!??!
    '{B087BE9D-ED37-454f-AF9C-04291E351182}'=@('\Machine\Preferences\Registry\Registry.xml')
    '{CAB54552-DEEA-4691-817E-ED4A4D1AFC72}'=@('\Machine\Preferences\Scheduledtasks\Scheduledtasks.xml','\User\Preferences\Scheduledtasks\Scheduledtasks.xml')
    '{728EE579-943C-4519-9EF7-AB56765798ED}'=@('\Machine\Preferences\DataSources\DataSources.xml','\User\Preferences\DataSources\DataSources.xml')
    '{91FBB303-0CD5-4055-BF42-E512A681B325}'=@('\Machine\Preferences\Services\Services.xml')
    }

}
# taken from https://github.com/PowerShellMafia/PowerSploit/blob/master/Exfiltration/Get-GPPPassword.ps1
function Get-DecryptedCpassword {
    [CmdletBinding()]
    Param (
        [string] $Cpassword 
    )

    try {
        #Append appropriate padding based on string length  
        $Mod = ($Cpassword.length % 4)
            
        switch ($Mod) {
        '1' {$Cpassword = $Cpassword.Substring(0,$Cpassword.Length -1)}
        '2' {$Cpassword += ('=' * (4 - $Mod))}
        '3' {$Cpassword += ('=' * (4 - $Mod))}
        }

        $Base64Decoded = [Convert]::FromBase64String($Cpassword)
            
        #Create a new AES .NET Crypto Object
        try
        {
            $AesObject = New-Object System.Security.Cryptography.AesCryptoServiceProvider -ErrorAction Stop
        }
        catch
        {
            # Added error handling to stop null being returned
            Write-Warning 'Unable to decrypt cPassword is .Net 3.5 installed?'
            return $Cpassword
        }
        [Byte[]] $AesKey = @(0x4e,0x99,0x06,0xe8,0xfc,0xb6,0x6c,0xc9,0xfa,0xf4,0x93,0x10,0x62,0x0f,0xfe,0xe8,
                                0xf4,0x96,0xe8,0x06,0xcc,0x05,0x79,0x90,0x20,0x9b,0x09,0xa4,0x33,0xb6,0x6c,0x1b)
            
        #Set IV to all nulls to prevent dynamic generation of IV value
        $AesIV = New-Object Byte[]($AesObject.IV.Length) 
        $AesObject.IV = $AesIV
        $AesObject.Key = $AesKey
        $DecryptorObject = $AesObject.CreateDecryptor() 
        [Byte[]] $OutBlock = $DecryptorObject.TransformFinalBlock($Base64Decoded, 0, $Base64Decoded.length)
            
        return [System.Text.UnicodeEncoding]::Unicode.GetString($OutBlock)
    } 
        
    catch {Write-Error $Error[0]}
}  

Function Get-GPLinks {
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [System.DirectoryServices.ResultPropertyCollection] $Properties
        )

    $Links = $Properties['gplink'] -split '\]\[' -replace '\[|\]'

    Foreach($Link in $Links)
    {
        $link = $link -split ';'
        $Out = '' | Select-Object PolicyGUID, LinkEnforced, LinkEnabled
        $Out.PolicyGUID = ($link[0] -replace ',cn=.*?$' -replace '^LDAP://cn=')
        $Out.LinkEnforced = ([int]$link[1] -ge 2)
        $Out.LinkEnabled = (0,2 -ccontains [int]$link[1])
        $Out
    }

}

Function Get-GPOOUs {
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [string] $GUID,
        [switch] $ReturnDNOnly,
        [string] $DomainController,
        [string] $Username,
        [string] $Password,
        [System.DirectoryServices.DirectoryEntry] $DirectoryEntry = $null,
        [switch] $LinkedOnly
        )

    begin
    {

        if(-not $DirectoryEntry)
        {
            # authenticate
            if($DomainController -and ([string]::IsNullOrEmpty($Username) -OR [string]::IsNullOrEmpty($Password))){
                # assume a remote domain that the current logon credentials have access to
                $DirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainController")
            }elseif([string]::IsNullOrEmpty($Username) -OR [string]::IsNullOrEmpty($Password)){
                # bind to the current domain
                $DirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry
            }else{
                # connect to a remote domain with credentials
                $DirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainController",$Username,$Password)
            }
        }

        # Build searcher
        $Searcher             = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot  = $DirectoryEntry
        $Searcher.Filter      = "(gPLink=*$GUID*)"

        $Results = $Searcher.FindAll()
        $Output = @()
    }

    process
    {

        foreach($Result in $Results)
        {
            # is the link enabled
            $GPLinks = Get-GPLinks -Properties $Result.Properties
            $GPLink = $GPLinks | Where-Object { $_.PolicyGUID -eq $GUID -and $_.LinkEnabled }

            if($LinkedOnly)
            {
                # Send DN to output
                if($ReturnDNOnly)
                {
                    # return string array of DNs
                    $Output += $Result.Properties['distinguishedname']
                }else{
                    # return a directory entry so its easier to process with other commands
                    $Output += $Result.GetDirectoryEntry()
                }
                
                # skip the rest of this loop
                return
            }


            if($GPLink)
            {
                if($GPLink.LinkEnforced)
                {
                    # link is enforced get all sub OUs
                    $Search = [adsisearcher]"objectclass=organizationalUnit"
                }else{
                    # gPOptions - 1 - Block Inheritance 
                    $Search = [adsisearcher]"(&(objectclass=organizationalUnit)(!(gPOptions=1)))"
                }
                $Search.SearchRoot = $Result.GetDirectoryEntry()
                $SubOUs = $Search.FindAll()
                foreach($SubOU in $SubOUs)
                {
                    # Send DN to output
                    if($ReturnDNOnly)
                    {
                        # return string array of DNs
                        $Output += $SubOU.Properties['distinguishedname']
                    }else{
                        # return a directory entry so its easier to process with other commands
                        $Output += $SubOU.GetDirectoryEntry()
                    }
                }
            }

        }
    }

    end
    {
        if($ReturnDNOnly)
        {
            $Output | Sort-Object -Unique
        }else{
            $Output | Sort-Object -Property distinguishedName -Unique
        }
    }

}
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
Function New-GPPCSEReport {
param($Data)

@"
<!DOCTYPE HTML>
<html>
<head>
	<style>
	body {
			font-family: "Helvetica Neue", Arial, sans-serif;
		}
	div {
			border: 2px solid;
			border-radius: 5px;
			width: 80%;
			margin:0 auto;
			padding: 10px;
		}
	h1 {
			width: 80%;
			margin:0 auto;
			padding: 10px;
		}
	table {
	  font-size: 14px;
	  border-collapse: collapse;
	}

	td, th {
	  padding: 10px;
	  text-align: left;
	  margin: 0;
	}

	tbody tr:nth-child(2n){
	  background-color: #eee;
	}

	th {
	  position: sticky;
	  top: 0;
	  background-color: #333;
	  color: white;
	}
</style>
</head>
<body>
	<h1>Invoke-GPPCSE Report</h1>
"@
foreach($Item in $Data)
{
@"
	<div>
		<h3>GPO Information</h3>
		<ul>
			<li>GPP Type: <b>$($Item.GPPType)</b></li>
			<li>GPP Username: <b>$($Item.GPPUsername)</b></li>
			<li>GPP New Username: <b>$($Item.GPPNewUsername)</b></li>
			<li>GPP Password: <b>$($Item.GPPPassword)</b></li>
			<li>GPO Path: <b>$($Item.GPOPath)</b></li>
			<li>GPO Name: <b>$($Item.GPOName)</b></li>
			<li>GPO GUID: <b>$($Item.GPOGUID)</b></li>
			<li>GPO Status: <b>$($Item.GPOStatus)</b></li>
			<li>GPO Linked OUs:
				<ul>
                    $($Item.GPOLinkedOUs | %{ "<li><b>$($_)</b></li>" })
				</ul>
			 </li>
			<li>GPO Associated OUs:
				<ul>
					$($Item.GPOAssociatedOUs | %{ "<li><b>$($_)</b></li>" })
				</ul>
			 </li>
		</ul>
      <h3>Related Active Computers</h3>
       $(if($Item.RelatedActiveComputers){ $Item.RelatedActiveComputers | ConvertTo-Html -Fragment }else{ '<i>None</i>' })

	</div>
<p>&nbsp;</p>
"@

}
@"
</body>
</html>
"@

}
