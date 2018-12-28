Function Compare-ADObject {
<#
.SYNOPSIS
    Compare AD Object an mounted Active Directory snapshot and another snapshot or current database.
.DESCRIPTION
    Uses for compare AD Objects and report changed attributes with old and new values.
.PARAMETER DestinationServer
    Server name of the database where destination objects are located.
    Default is localhost.
.PARAMETER DestinationLDAPPort
    Port of the destination database is mounted. Used for difference objects.
    Must be in the range 1025 to 65535.
.PARAMETER DestinationServer
    Server name of the database where source objects are located.
    Default is localhost.
.PARAMETER SourceLDAPPort
    Port of the source database is mounted. Used for reference object.
    If not specified 389 used.
    Must be in the range 1025 to 65535.
    Checks to make sure another database is not already using the port.
.PARAMETER FilterDate
    Uses for LDAP whenChanged filter criteria.
    If not specified prompts the user for which snapshot to mount.
.PARAMETER Properties
    One or more Active Directory attribute names delimited by commas.
    Supports single-value and multi-value attributes.
    Default wildcard.
.PARAMETER ExcludeProperty
    One or more Active Directory attribute names delimited by commas.
    Supports single-value and multi-value attributes.
    Default common attributes changed frequently.
.PARAMETER SearchBase
    Specifies an Active Directory path to search.
    Default value is the default naming context of the LDAP.
.PARAMETER ObjectClass
    Uses for LDAP ObjectClass filter criteria.
    Default uses User.
.PARAMETER MatchingProperty
    Uses for locate the object to destination source.
    Default uses sAMAccountName.
.PARAMETER IncludeDeletedObjects
    Uses for retrieves deleted objects and the deactivated forward and backward links.
    Default set to false. Not include deleted objects and links.
.PARAMETER Output
    Uses for user friendly viewing the changed objects.
.OUTPUTS
    PowerShell custom object for viewing changed properties:
        Identity          : Matching property value.
        SourceObject      : DistinguishedName of the object located in the source database (like current database)
        DestinationObject : DistinguishedName of the object located in the destination database (like snapshot database)
        ChangedProperties : Powershell custom object for viewing changed attributes
            Name     : Chagned property name
            OldValue : Destination object property value
            NewValue : Source object property value
        WhenChanged       : Source object whenChanged value
        Renamed           : Is CN attribute changed.
        Moved             : Compare source and destionation object location.
        Deleted           : Source object Deleted value
        Action            : Describes the changed reason: Changed, NotFound
.EXAMPLE
    Compare-ADObject -DestinationLDAPPort 33389
.EXAMPLE
    Compare-ADObject -DestinationLDAPPort 33389 -ObjectClass "User"
.EXAMPLE
    Compare-ADObject -DestinationLDAPPort 33389 -FilterDate (Get-Date).AddDays(-2) -ObjectClass "User"
.EXAMPLE
    Compare-ADObject -DestinationLDAPPort 33389 -SourceLDAPPort 33390 -FilterDate (Get-Date).AddDays(-2) -Property "DisplayName", "MemberOf"
.LINK
    http://www.savaskartal.com
    https://github.com/sskl
#>
Param(
    [parameter(Mandatory=$false)]
    [String[]]
    $DestinationServer = "localhost",
    [Parameter(Mandatory=$true)]
    [ValidateScript({
        # Must specify an LDAP port mounted
        $x = $null
        Try {$x = Get-ADRootDSE -Server localhost:$_}
        Catch {$null}
        If ($x) {$true} Else {$false}
    })]
    [ValidateRange(1025,65535)]
    [int]
    $DestinationLDAPPort,
    [parameter(Mandatory=$false)]
    [String[]]
    $SourceServer = "localhost",
    [Parameter(Mandatory=$false)]
    [ValidateScript({
        # Must specify an LDAP port mounted
        $x = $null
        Try {$x = Get-ADRootDSE -Server localhost:$_}
        Catch {$null}
        If ($x) {$true} Else {$false}
    })]
    [int]
    $SourceLDAPPort = 389,
    [parameter(Mandatory=$false)]
    [String[]]
    $Properties = "*",
    [parameter(Mandatory=$false)]
    [String[]]
    $ExcludeProperty = @("lastLogoff","lastLogon","lastLogonTimestamp","logonCount","badPasswordTime","pwdLastSet","Modified","modifyTimeStamp","uSNChanged","msDS-AuthenticatedAtDC","msDS-FailedInteractiveLogonCount","msDS-FailedInteractiveLogonCountAtLastSuccessfulLogon","msDS-FailedInteractiveLogonTime","msDS-LastSuccessfulInteractiveLogonTime","DirXML-Associations","ACL","sDRightsEffective","dSCorePropagationData","msExchUMDtmfMap"),
    [Parameter(Mandatory=$false)]
    [datetime]
    $FilterDate,
    [Parameter(Mandatory=$false)]
    [string]
    $SearchBase = (Get-ADRootDSE).defaultNamingContext,
    [Parameter(Mandatory=$false)]
    [string]
    $ObjectClass = "User",
    [Parameter(Mandatory=$false)]
    [ValidateSet("sAMAccountName", "objectGuid", "DistinguishedName")]
    [string]
    $MatchingProperty = "sAMAccountName",
    [Parameter(Mandatory=$false)]
    [switch]
    $IncludeDeletedObjects = $false,
    [Parameter(Mandatory=$false)]
    [ValidateSet("Html")]
    [string]
    $Output
)

    If (! $PSBoundParameters.ContainsKey('FilterDate')) {
        $MountedSnapshots = Get-WmiObject Win32_ShadowCopy | Where-Object ExposedName -ne $null | Select-Object Id, Installdate, OriginatingMachine, ExposedName
        If ($null -eq $MountedSnapshots) {
            Write-Error "Not Found mounted snapshots, please remount snapshot."
            Return
        }

        $Choice = $MountedSnapshots | Out-GridView -Title 'Select mounted snapshot for FilterDate parameter' -OutputMode Single
        If ($null -eq $Choice) {
            # What if the user hits the Cancel button in the OGV?
            Return
        } Else {
            $Date = ($Choice.ExposedName -split "_")[1].Trim()
            $FilterDate = [DateTime]::ParseExact($Date, "yyyyMMddHHmm", $null)
        }
    }

    Write-Verbose "FilterDate = $FilterDate"
    Write-Verbose "SearchBase = $SearchBase"
    Write-Verbose "ObjectClass = $ObjectClass"
    Write-Verbose "MatchingProperty = $MatchingProperty"

    # Set $PropertiesGet and $PropertyNames variable to *.
    $PropertiesGet = $PropertyNames = '*'
    # If $Properties parameter set, make sure all required parameters adding to $PropertiesGet variable.
    If ($Properties -ne '*') {
        $PropertiesGet = $Properties + @("sAMAccountName", "objectGUID", "DistinguishedName", "Deleted", "objectClass")
        $PropertyNames = $Properties
    }

    try {
        $SourceObjects = Get-ADObject -Filter 'ObjectClass -eq $ObjectClass -and whenChanged -gt $FilterDate' -SearchBase $SearchBase -Properties $PropertiesGet -Server "$($SourceServer):$($SourceLDAPPort)" -IncludeDeletedObjects:$IncludeDeletedObjects -ErrorAction Stop |
                         Select-Object -Property * -ExcludeProperty $ExcludeProperty
    } catch {
        $SourceObjects = $null
    }

    If ($null -eq $SourceObjects) {
        Write-Warning "Changed source objects not found filter by ObjectClass attribute $ObjectClass and whenChanged attribute $Date."
        return
    }

    $Result = @()

    $SourceObjects | ForEach-Object {

        $SourceObject = $PSItem
        $MatchingPropertyValue = $SourceObject.$MatchingProperty

        $DestinationObject = Get-ADObject -Filter 'ObjectClass -eq $ObjectClass -and $MatchingProperty -eq $MatchingPropertyValue' -Properties $PropertiesGet -Server "$($DestinationServer):$($DestinationLDAPPort)" -IncludeDeletedObjects:$IncludeDeletedObjects -ErrorAction Stop |
                             Select-Object -Property * -ExcludeProperty $ExcludeProperty

        $DestinationObjectOU = $DestinationObject.DistinguishedName -replace '^.+?,(CN|OU.+)', '$1'
        $SourceObjectOU = $SourceObject.DistinguishedName -replace '^.+?,(CN|OU.+)', '$1'

        $Object = New-Object -TypeName PSCustomObject -Property @{
            Identity          = $SourceObject.$MatchingProperty
            SourceObject      = $SourceObject.DistinguishedName
            DestinationObject = $DestinationObject.DistinguishedName
            ChangedProperties = @()
            WhenChanged       = $SourceObject.whenChanged
            User              = ($SourceObject.objectClass -eq 'user')
            Renamed           = ($SourceObject.CN -ne $DestinationObject.CN)
            Moved             = ($SourceObjectOU -ne $DestinationObjectOU)
            Deleted           = ($SourceObject.Deleted -eq $true)
            Action            = ''
        }

        If ($null -eq $DestinationObject) {
            Write-Verbose "$($SourceObject.$MatchingProperty) not found on destination."
            $Object.Action = 'NotFound'
        } Else {

            # If the compared properties not specified then the source and destination object's property names are combined.
            # Because of if object properties setting to null then Get-ADObject cmd-let output not return that properties.
            If ($Properties -eq '*') {
                $PropertyNames = $SourceObject.PropertyNames + $DestinationObject.PropertyNames | Select-Object -Unique
            }

            foreach ($Property in $PropertyNames) {

                # Skip whenChanged property.
                If ($Property -eq 'whenChanged') {
                    continue
                }

                $IsPropertyChanged = $false

                try {
                    # If the property value returns null, then Select-Object -ExpandProperty errors.
                    # Select-Object does not seem to obey when you specify -ErrorAction SilentlyContinue.
                    $SourceValue = $null = $SourceObject | Select-Object -ExpandProperty $Property -ErrorAction Stop
                } catch {
                    $SourceValue = $null
                }

                try {
                    # If the property value returns null, then Select-Object -ExpandProperty errors.
                    # Select-Object does not seem to obey when you specify -ErrorAction SilentlyContinue.
                    $DestinationValue = $null = $DestinationObject | Select-Object -ExpandProperty $Property -ErrorAction Stop
                } catch {
                    $DestinationValue = $null
                }

                # Compare source and destination value two ways.
                # Because of;
                #     $SourceValue = "abc"
                #     $DestinationValue =  "abc", "def"
                #         $SourceValue -eq $DestinationValue => return abc
                #         $DestinationValue -eq $SourceValue => return false
                # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comparison_operators?view=powershell-6#equality-operators
                If (($SourceValue -eq $DestinationValue) -and ($DestinationValue -eq $SourceValue)) {
                    Write-Verbose "$($SourceObject.$MatchingProperty) $Property property has same value. (SourceValue = $SourceValue, DestinationValue = $DestinationValue)"
                } Else {
                    # Is this a multi-value attribute?
                    $PropertyDefinition = $SourceObject | Get-Member | Where-Object { $_.Name -eq $Property } | Select-Object -ExpandProperty Definition
                    If ($PropertyDefinition -match "ADPropertyValueCollection|ActiveDirectorySecurity|HashSet|byte\[\]") {

                        # If either of the values are NULL, then Compare-Object will error.
                        If (($null -eq $SourceValue) -or ($null -eq $DestinationValue)) {
                            $IsPropertyChanged = $true
                        } ElseIf (Compare-Object -ReferenceObject $SourceValue -DifferenceObject $DestinationValue) {
                            $IsPropertyChanged = $true
                        }

                    } Else {
                        $IsPropertyChanged = $true
                    }

                    If ($IsPropertyChanged) {
                        Write-Verbose "$($SourceObject.$MatchingProperty) $Property property changed. (SourceValue = $SourceValue, DestinationValue = $DestinationValue)"
    
                        $ChangedProperty = New-Object -TypeName PSCustomObject -Property @{
                                                Name     = $Property
                                                OldValue = $DestinationValue
                                                NewValue = $SourceValue
                                            }
                        $Object.ChangedProperties += $ChangedProperty
                        $Object.Action = 'Changed'
                    }
                }
            }
        }

        If ($Object.Action) {
            $Result += $Object
        }

    }

    # Save output as html format.
    If ($Result -and $Output) {
        $FormattedOutput = $Result |
            Select-Object Identity, SourceObject, DestinationObject, WhenChanged,
                    @{
                        Name='ChangedProperties';
                        Expression={ $PSItem.ChangedProperties |
                                    Select-Object Name,
                                        @{Name='NewValue';Expression={ If ($PSItem.NewValue.GetType().IsArray) { $PSItem | Select-Object -ExpandProperty NewValue } Else { $PSItem.NewValue } } },
                                        @{Name='OldValue';Expression={ If ($PSItem.OldValue.GetType().IsArray) { $PSItem | Select-Object -ExpandProperty OldValue } Else { $PSItem.OldValue } } } |
                                    ConvertTo-Html -As Table -Fragment |
                                    Out-String }
                    },
                    Action, User, Renamed, Moved, Deleted

        If ($Output -eq "Html") {
            $CSS =
@"
    <style type="text/css">
        body {font-family: Segoe UI; font-size: 11px;}
        h1, h5, th { text-align: center; }
        table { margin: auto; border: thin ridge grey; }
        th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
        td { padding: 2px 10px; color: #000; }
        tr { background: #dae5f4; }
        table table { border: 0px; width: 100%; }
        table table tr { background: #ecf2f9; }
        table table th { background: #1166ff; }
    </style>
"@

            Add-Type -AssemblyName System.Web
            [System.Web.HttpUtility]::HtmlDecode(($FormattedOutput | ConvertTo-Html -Head $CSS)) | Out-File ".\Output_$(Get-Date -UFormat %Y%m%d%H%M).html" -Encoding utf8
        }
    }
    return $Result

}