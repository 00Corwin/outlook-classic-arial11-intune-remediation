<#
.SYNOPSIS
    Captures the six Classic Outlook MailSettings REG_BINARY font values.

.DESCRIPTION
    Run this as the signed-in reference user after configuring Classic Outlook
    to the required font. It makes no registry changes.
#>

[CmdletBinding()]
param()

$Path = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\MailSettings'
$Names = @(
    'ComposeFontSimple',
    'ComposeFontComplex',
    'ReplyFontSimple',
    'ReplyFontComplex',
    'TextFontSimple',
    'TextFontComplex'
)

foreach ($Name in $Names) {
    try {
        $Bytes = [byte[]](Get-ItemPropertyValue -Path $Path -Name $Name -ErrorAction Stop)
        $Hex = ($Bytes | ForEach-Object { $_.ToString('x2') }) -join ','
        Write-Output "$Name = $Hex"
    }
    catch {
        Write-Output "$Name = <NOT PRESENT>"
    }
}
