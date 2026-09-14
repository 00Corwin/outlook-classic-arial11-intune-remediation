# Outlook Classic Arial 11 - REMEDIATION
# Run in Intune as SYSTEM, 64-bit PowerShell.

# SECTION 1 - Approved configuration
$BasePath = 'SOFTWARE\Microsoft\Office\16.0\Common\MailSettings'

$Expected = [ordered]@{

    ComposeFontSimple = '3c,00,00,00,1f,00,00,f8,00,00,00,40,dc,00,00,00,00,00,00,00,00,00,00,00,00,22,41,72,69,61,6c,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'

    ComposeFontComplex = '3c,68,74,6d,6c,3e,0d,0a,0d,0a,3c,68,65,61,64,3e,0d,0a,3c,73,74,79,6c,65,3e,0d,0a,0d,0a,20,2f,2a,20,53,74,79,6c,65,20,44,65,66,69,6e,69,74,69,6f,6e,73,20,2a,2f,0d,0a,20,73,70,61,6e,2e,50,65,72,73,6f,6e,61,6c,43,6f,6d,70,6f,73,65,53,74,79,6c,65,0d,0a,09,7b,6d,73,6f,2d,73,74,79,6c,65,2d,6e,61,6d,65,3a,22,50,65,72,73,6f,6e,61,6c,20,43,6f,6d,70,6f,73,65,20,53,74,79,6c,65,22,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,74,79,70,65,3a,70,65,72,73,6f,6e,61,6c,2d,63,6f,6d,70,6f,73,65,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,6e,6f,73,68,6f,77,3a,79,65,73,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,75,6e,68,69,64,65,3a,6e,6f,3b,0d,0a,09,6d,73,6f,2d,61,6e,73,69,2d,66,6f,6e,74,2d,73,69,7a,65,3a,31,31,2e,30,70,74,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,73,69,7a,65,3a,31,32,2e,30,70,74,3b,0d,0a,09,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,41,72,69,61,6c,22,2c,73,61,6e,73,2d,73,65,72,69,66,3b,0d,0a,09,6d,73,6f,2d,61,73,63,69,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,41,72,69,61,6c,3b,0d,0a,09,6d,73,6f,2d,68,61,6e,73,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,41,72,69,61,6c,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,54,69,6d,65,73,20,4e,65,77,20,52,6f,6d,61,6e,22,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,62,69,64,69,3b,0d,0a,09,63,6f,6c,6f,72,3a,77,69,6e,64,6f,77,74,65,78,74,3b,0d,0a,09,66,6f,6e,74,2d,77,65,69,67,68,74,3a,6e,6f,72,6d,61,6c,3b,0d,0a,09,66,6f,6e,74,2d,73,74,79,6c,65,3a,6e,6f,72,6d,61,6c,3b,7d,0d,0a,2d,2d,3e,0d,0a,3c,2f,73,74,79,6c,65,3e,0d,0a,3c,2f,68,65,61,64,3e,0d,0a,0d,0a,3c,2f,68,74,6d,6c,3e,0d,0a'

    ReplyFontSimple = '3c,00,00,00,1f,00,00,f8,00,00,00,40,dc,00,00,00,00,00,00,00,00,00,00,00,00,22,41,72,69,61,6c,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'

    ReplyFontComplex = '3c,68,74,6d,6c,3e,0d,0a,0d,0a,3c,68,65,61,64,3e,0d,0a,3c,73,74,79,6c,65,3e,0d,0a,0d,0a,20,2f,2a,20,53,74,79,6c,65,20,44,65,66,69,6e,69,74,69,6f,6e,73,20,2a,2f,0d,0a,20,73,70,61,6e,2e,50,65,72,73,6f,6e,61,6c,52,65,70,6c,79,53,74,79,6c,65,31,0d,0a,09,7b,6d,73,6f,2d,73,74,79,6c,65,2d,6e,61,6d,65,3a,22,50,65,72,73,6f,6e,61,6c,20,52,65,70,6c,79,20,53,74,79,6c,65,31,22,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,74,79,70,65,3a,70,65,72,73,6f,6e,61,6c,2d,72,65,70,6c,79,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,6e,6f,73,68,6f,77,3a,79,65,73,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,75,6e,68,69,64,65,3a,6e,6f,3b,0d,0a,09,6d,73,6f,2d,61,6e,73,69,2d,66,6f,6e,74,2d,73,69,7a,65,3a,31,31,2e,30,70,74,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,73,69,7a,65,3a,31,32,2e,30,70,74,3b,0d,0a,09,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,41,72,69,61,6c,22,2c,73,61,6e,73,2d,73,65,72,69,66,3b,0d,0a,09,6d,73,6f,2d,61,73,63,69,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,41,72,69,61,6c,3b,0d,0a,09,6d,73,6f,2d,61,73,63,69,69,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,6c,61,74,69,6e,3b,0d,0a,09,6d,73,6f,2d,66,61,72,65,61,73,74,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,41,72,69,61,6c,3b,0d,0a,09,6d,73,6f,2d,66,61,72,65,61,73,74,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,6c,61,74,69,6e,3b,0d,0a,09,6d,73,6f,2d,68,61,6e,73,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,41,72,69,61,6c,3b,0d,0a,09,6d,73,6f,2d,68,61,6e,73,69,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,6c,61,74,69,6e,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,54,69,6d,65,73,20,4e,65,77,20,52,6f,6d,61,6e,22,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,62,69,64,69,3b,0d,0a,09,63,6f,6c,6f,72,3a,77,69,6e,64,6f,77,74,65,78,74,3b,7d,0d,0a,2d,2d,3e,0d,0a,3c,2f,73,74,79,6c,65,3e,0d,0a,3c,2f,68,65,61,64,3e,0d,0a,0d,0a,3c,2f,68,74,6d,6c,3e,0d,0a'

    TextFontSimple = '3c,00,00,00,1f,00,00,f8,00,00,00,40,dc,00,00,00,00,00,00,00,00,00,00,00,00,22,41,72,69,61,6c,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00'

    TextFontComplex = '3c,68,74,6d,6c,3e,0d,0a,0d,0a,3c,68,65,61,64,3e,0d,0a,3c,73,74,79,6c,65,3e,0d,0a,0d,0a,20,2f,2a,20,53,74,79,6c,65,20,44,65,66,69,6e,69,74,69,6f,6e,73,20,2a,2f,0d,0a,20,70,2e,4d,73,6f,50,6c,61,69,6e,54,65,78,74,2c,20,6c,69,2e,4d,73,6f,50,6c,61,69,6e,54,65,78,74,2c,20,64,69,76,2e,4d,73,6f,50,6c,61,69,6e,54,65,78,74,0d,0a,09,7b,6d,73,6f,2d,73,74,79,6c,65,2d,6e,6f,73,68,6f,77,3a,79,65,73,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,70,72,69,6f,72,69,74,79,3a,39,39,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,6c,69,6e,6b,3a,22,50,6c,61,69,6e,20,54,65,78,74,20,43,68,61,72,22,3b,0d,0a,09,6d,61,72,67,69,6e,3a,30,63,6d,3b,0d,0a,09,6d,73,6f,2d,70,61,67,69,6e,61,74,69,6f,6e,3a,77,69,64,6f,77,2d,6f,72,70,68,61,6e,3b,0d,0a,09,66,6f,6e,74,2d,73,69,7a,65,3a,31,31,2e,30,70,74,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,73,69,7a,65,3a,31,30,2e,35,70,74,3b,0d,0a,09,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,41,72,69,61,6c,22,2c,73,61,6e,73,2d,73,65,72,69,66,3b,0d,0a,09,6d,73,6f,2d,66,61,72,65,61,73,74,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,54,69,6d,65,73,20,4e,65,77,20,52,6f,6d,61,6e,22,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,54,69,6d,65,73,20,4e,65,77,20,52,6f,6d,61,6e,22,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,62,69,64,69,3b,0d,0a,09,6d,73,6f,2d,66,6f,6e,74,2d,6b,65,72,6e,69,6e,67,3a,31,2e,30,70,74,3b,0d,0a,09,6d,73,6f,2d,6c,69,67,61,74,75,72,65,73,3a,73,74,61,6e,64,61,72,64,63,6f,6e,74,65,78,74,75,61,6c,3b,0d,0a,09,6d,73,6f,2d,66,61,72,65,61,73,74,2d,6c,61,6e,67,75,61,67,65,3a,45,4e,2d,55,53,3b,7d,0d,0a,2d,2d,3e,0d,0a,3c,2f,73,74,79,6c,65,3e,0d,0a,3c,2f,68,65,61,64,3e,0d,0a,0d,0a,3c,2f,68,74,6d,6c,3e,0d,0a'
}

# SECTION 2 - Helper functions
function Convert-HexToBytes {
    param([Parameter(Mandatory)][string]$Hex)

    $CleanHex = $Hex -replace '\s', ''

    [byte[]](($CleanHex -split ',') |
        Where-Object { $_ -ne '' } |
        ForEach-Object {
            [Convert]::ToByte($_,16)
        })
}

function Test-BytesEqual {
    param(
        [byte[]]$Actual,
        [byte[]]$Expected
    )

    if ($null -eq $Actual -or $null -eq $Expected) {
        return $false
    }

    if ($Actual.Length -ne $Expected.Length) {
        return $false
    }

    return (
        [Convert]::ToBase64String($Actual) -eq
        [Convert]::ToBase64String($Expected)
    )
}

# SECTION 3 - Find eligible loaded user hives
$UserSids = @(
    (Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue).PSChildName |
        Where-Object {
            $_ -match '^S-1-(5-21|12-1)-\d+-\d+-\d+-\d+$'
        }
)

if ($UserSids.Count -eq 0) {
    Write-Output 'Skipped: no eligible user registry hive is currently loaded.'
    exit 0
}

# SECTION 4 - Apply and verify
$Failures = [System.Collections.Generic.List[string]]::new()

foreach ($Sid in $UserSids) {

    $Path = "Registry::HKEY_USERS\$Sid\$BasePath"

    try {
        if (-not (Test-Path $Path)) {
            New-Item `
                -Path $Path `
                -Force `
                -ErrorAction Stop | Out-Null
        }
    }
    catch {
        $Failures.Add(
            "SID=${Sid}; Reason=Unable to create/access MailSettings: $($_.Exception.Message)"
        )
        continue
    }

    foreach ($Name in $Expected.Keys) {

        try {
            $Wanted = Convert-HexToBytes -Hex $Expected[$Name]

            New-ItemProperty `
                -Path $Path `
                -Name $Name `
                -PropertyType Binary `
                -Value $Wanted `
                -Force `
                -ErrorAction Stop | Out-Null

            $Actual = [byte[]](Get-ItemPropertyValue `
                -Path $Path `
                -Name $Name `
                -ErrorAction Stop)

            if (-not (Test-BytesEqual -Actual $Actual -Expected $Wanted)) {
                $Failures.Add(
                    "SID=${Sid}; Value=${Name}; Reason=Verification failed after write"
                )
            }
        }
        catch {
            $Failures.Add(
                "SID=${Sid}; Value=${Name}; Reason=$($_.Exception.Message)"
            )
        }
    }
}

# SECTION 5 - Intune result
if ($Failures.Count -gt 0) {
    Write-Output ("Remediation failed: " + ($Failures -join ' | '))
    exit 1
}

Write-Output 'Remediated: Outlook Classic new mail, replies/forwards and plain text configured for Arial 11.'
exit 0
