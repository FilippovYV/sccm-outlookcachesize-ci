$EmailToKeep = [byte[]](03,00,00,00)

$EmailToKeepRegistry = '00036649'
$Outlook16Profiles = 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\'

$Profiles = Get-ChildItem -Path Registry::$Outlook16Profiles -Recurse | Where-Object { $_.Property -eq $EmailToKeepRegistry }
$Profiles | ForEach-Object {
  Set-ItemProperty -Path Registry::$_ -Name $EmailToKeepRegistry -Value $EmailToKeep
}
