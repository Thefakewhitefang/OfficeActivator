cls

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())

If (($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) -eq $false){
 Write-Host "Please run with administrator privileges"
 Write-Host "Exiting..."
 Start-Sleep 2
 exit
 }
 Else{}

$ARCH = (wmic os get OSArchitecture)[2]

$KMS = Read-Host -Prompt "Set the KMS server(Leave blank for default)"
If ($KMS -ieq ""){
 $KMS="s8.uk.to"
 }

If ($ARCH = "64-bit"){
  cls
  Write-Host "64-bit system detected"
  Write-Host " "
  pause
  cd "$env:ProgramFiles\Microsoft Office\Office16"
 }

Else{
  cls
  Write-Host "32-bit system detected"
  Write-Host " "
  pause
  cd "$env:ProgramFiles(x86)\Microsoft Office\Office16"
 }

$DIR = "C:\Program Files\Microsoft Office\root\Licenses16\"
$List = Get-ChildItem -Path "C:\Program Files\Microsoft Office\root\Licenses16\*" -Include ProPlus2019VL_KMS_Client_AE-ppd*.xrm-ms | select-object -expandproperty Name

$CharArray = $List.Split(" ")

Foreach ($i in $CharArray){
 $Item = $DIR + $i
 cscript ospp.vbs /inslic:$Item
}

cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6F7TH
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:$KMS
cscript ospp.vbs /act
Start-Sleep 3

cd $env:USERPROFILE\Desktop
cls
Write-Host "Exiting..."
Start-Sleep 3
exit