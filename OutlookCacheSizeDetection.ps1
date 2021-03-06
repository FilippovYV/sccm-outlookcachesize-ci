# Size of Outlook Cache 
$EmailToKeepMonths=3

#Construct binary value and then converting it to string for comparison
$EmailToKeepString = [byte[]]($EmailToKeepMonths,00,00,00) -join ''
#Registry name
$EmailToKeepRegistry = '00036649'
#Registry key for profiles store. Outlook 2016/2019/365
$Outlook16Profiles = 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\'

$EmailToKeepInSync = $true
Get-ChildItem -Path Registry::$Outlook16Profiles -Recurse | Where-Object { $_.Property -eq $EmailToKeepRegistry } | ForEach-Object { 
  if (((Get-ItemProperty -Path Registry::$_ -Name $EmailToKeepRegistry).$EmailToKeepRegistry -join '') -ne $EmailToKeepString) {
    $EmailToKeepInSync = $false
  }
}
Return $EmailToKeepInSync