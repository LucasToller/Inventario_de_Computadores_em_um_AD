param(
  [Parameter(Mandatory=$true)][string]$Sam
)
Import-Module ActiveDirectory -ErrorAction Stop
$u = Get-ADUser -Identity $Sam -Properties DisplayName -ErrorAction SilentlyContinue
if ($u) { $u.DisplayName }