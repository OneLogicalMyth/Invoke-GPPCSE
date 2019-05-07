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