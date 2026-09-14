<#
.SYNOPSIS
    Removes only the six Classic Outlook font values managed by this package.

.DESCRIPTION
    Disable or remove the Intune Remediation assignment before running this
    rollback script, otherwise the remediation can reapply the settings.

    This returns the values to an unmanaged state. It does not reconstruct
    each user's previous custom font selections.
#>

[CmdletBinding()]
param()

$BasePath = 'SOFTWARE\Microsoft\Office\16.0\Common\MailSettings'
$Names = @(
    'ComposeFontSimple',
    'ComposeFontComplex',
    'ReplyFontSimple',
    'ReplyFontComplex',
    'TextFontSimple',
    'TextFontComplex'
)

$UserSids = @(
    (Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue).PSChildName |
        Where-Object {
            $_ -match '^S-1-(5-21|12-1)-\d+-\d+-\d+-\d+$'
        }
)

foreach ($Sid in $UserSids) {
    $Path = "Registry::HKEY_USERS\$Sid\$BasePath"

    foreach ($Name in $Names) {
        Remove-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    }
}

Write-Output 'Rollback completed. Managed Outlook Classic font values removed.'
exit 0
