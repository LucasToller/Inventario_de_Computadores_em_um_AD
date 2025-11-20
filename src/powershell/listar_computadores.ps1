# Lista computadores no AD dentro de uma Unidade Organizacional (OU)
param(
  [Parameter(Mandatory=$true)][string]$SearchBase
)
Import-Module ActiveDirectory -ErrorAction Stop
Get-ADComputer -SearchBase $SearchBase -Filter * | Select-Object -ExpandProperty Name