$EmailToKeep = [byte[]](03,00,00,00)
$EmailToKeepString = $EmailToKeep -join ''
$EmailToKeepRegistry = '00036649'
$Outlook16Profiles = 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\'

$EmailToKeepInSync = $true
$Profiles = Get-ChildItem -Path Registry::$Outlook16Profiles -Recurse | Where-Object { $_.Property -eq $EmailToKeepRegistry }
$Profiles | ForEach-Object { 
  if (((Get-ItemProperty -Path Registry::$_ -Name $EmailToKeepRegistry).$EmailToKeepRegistry -join '') -ne $EmailToKeepString) {
    $EmailToKeepInSync = $false
  }
}
Return $EmailToKeepInSync