# Size of Outlook Cache 
$EmailToKeepMonths=3

#Construct binary value
$EmailToKeepValue = [byte[]]($EmailToKeepMonths,00,00,00)
#Regsitry name
$EmailToKeepRegistry = '00036649'
#Registry key for profiles store. Outlook 2016/2019/365
$Outlook16Profiles = 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\'

$Profiles = Get-ChildItem -Path Registry::$Outlook16Profiles -Recurse | Where-Object { $_.Property -eq $EmailToKeepRegistry }
$Profiles | ForEach-Object {
  Set-ItemProperty -Path Registry::$_ -Name $EmailToKeepRegistry -Value $EmailToKeepValue
}
