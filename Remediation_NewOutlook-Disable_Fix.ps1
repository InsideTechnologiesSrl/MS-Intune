#Get SID of current interactive users
$CurrentLoggedOnUser = (Get-CimInstance win32_computersystem).UserName
if (-not ([string]::IsNullOrEmpty($CurrentLoggedOnUser))) {
    $AdObj = New-Object System.Security.Principal.NTAccount($CurrentLoggedOnUser)
    $strSID = $AdObj.Translate([System.Security.Principal.SecurityIdentifier])
    $userSID = $strSID.Value
} else {
    $userSID = $null
}

$regpath="Software\Policies\Microsoft\office\16.0\outlook\preferences\"
$name="NewOutlookMigrationUserSetting"
$value=0

New-PSDrive -PSProvider Registry -Name "HKU" -Root HKEY_USERS | Out-Null
$regkey = "HKU:\$userSID\$regpath"

# Create the registry key if it doesn't exist
If (!(Test-Path $regkey))
{
New-Item -Path $regkey -ErrorAction stop
}

# Create the registry value if it doesn't exist
if (!(Get-ItemProperty -Path $regkey -Name $name -ErrorAction SilentlyContinue))
{
New-ItemProperty -Path $regkey -Name $name -Value $value -PropertyType DWORD -ErrorAction stop
write-output "remediation complete"
Remove-PSDrive -Name "HKU" | Out-Null
exit 0
}

# Update the registry value if it exists
set-ItemProperty -Path $regkey -Name $name -Value $value -ErrorAction stop
write-output "remediation complete"
Remove-PSDrive -Name "HKU" | Out-Null
exit 0