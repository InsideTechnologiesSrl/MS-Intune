#Get SID of current interactive users
$CurrentLoggedOnUser = (Get-CimInstance win32_computersystem).UserName
if (-not ([string]::IsNullOrEmpty($CurrentLoggedOnUser))) {
    $AdObj = New-Object System.Security.Principal.NTAccount($CurrentLoggedOnUser)
    $strSID = $AdObj.Translate([System.Security.Principal.SecurityIdentifier])
    $UserSID = $strSID.Value
} else {
    $UserSID = $null
}

$regpath="Software\Policies\Microsoft\office\16.0\outlook\preferences\"
$name="NewOutlookMigrationUserSetting"
$value=0

New-PSDrive -PSProvider Registry -Name "HKU" -Root HKEY_USERS | Out-Null
$regkey = "HKU:\$UserSID\$regpath"


If (!(Test-Path $regkey))
{
Write-Output 'RegKey not available - remediate'
Remove-PSDrive -Name "HKU" | Out-Null
Exit 1
}

$check=(Get-ItemProperty -path $regkey -name $name -ErrorAction SilentlyContinue).$name
if ($check -eq $value){
write-output 'setting ok - no remediation required'
Remove-PSDrive -Name "HKU" | Out-Null
Exit 0
}

else {
write-output 'value not ok, no value or could not read - go and remediate'
Remove-PSDrive -Name "HKU" | Out-Null
Exit 1
}