# Compare-ADObject

Compare Active Directory Objects located in mounted source and destination AD databases. You can use for compare to mounted AD snapshot database and current AD database then observe the single and multi attribute based changes.

# Usage

First you can mount an AD snapshot database.

```powershell
# Mount to latest snapshot run below command.
ntdsutil snapshot "list all" "mount 1" quit quit
# Changed $SNAP_201812121230_VOLUMEC$ with your mounted snapshot directory.
dsamain.exe -dbpath 'C:\$SNAP_201812121230_VOLUMEC$\Windows\NTDS\ntds.dit' -ldapport 33389
```

Compare mounted AD snapshot database and live AD database.

```powershell
# Download ps1 file to local folder.
. .\Compare-ADObject.ps1

Set-ExecutionPolicy -Bypass

# Get changed user and computer objects.
$ChangedObjects = Compare-ADObject -DestinationLDAPPort 33389 -Output Html
# Filter only user objects.
$ChangedUsers = $ChangedObjects | ? User -eq $true
# Filter only moved user objects.
$MovedUsers = $ChangedObjects | ? User -eq $true | ? Moved -eq $true | ? Deleted -ne $true

# Get changed groups objects.
$ChangedObjects = Compare-ADObject -DestinationLDAPPort 33389 -ObjectClass group -Output Html
$ChangedGroups = $ChangedObjects | Select -ExpandProperty Identity | Get-ADGroup
```