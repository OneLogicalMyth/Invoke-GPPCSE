
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