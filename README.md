# Invoke-GPPCSE
Obtains a list of GPOs based on known Client Side Extensions (CSE) that normally contain passwords.

# Running
First import the module.

```
Import-Module .\GPPCSE.psd1
```

Running against the local domain.

```
Invoke-GPPCSE
```

Running against a trusted domain. This will use your current credentials to authenticate against the remote domain.

```
Invoke-GPPCSE -DomainController 192.168.1.1
```

Running against a remote domain.

```
Invoke-GPPCSE -DomainController 192.168.1.1 -Username LowPriv -Password Password1
```

You can create a HTML report of the results by doing the following:

```
$Results = Invoke-GPPCSE
New-GPPCSEReport -Data $Results | Out-File Report.html
```
