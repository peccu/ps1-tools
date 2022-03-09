# based on https://community.spiceworks.com/topic/2292318-select-audio-device-with-powershell
If (! (Get-Module -Name "AudioDeviceCmdlets" -ListAvailable)) {
  # but needs Admin
  Install-Module -Name AudioDeviceCmdlets -Force -Verbose  
  get-module -Name "AudioDeviceCmdlets" -ListAvailable | Sort-Object Version | select -last 1 | Import-Module -Verbose
}

if ( Get-AudioDevice -Recording | where name -like "* SPEAK *"){
  echo "JabraSPEAK -> JabraLink"
  Get-AudioDevice -List | where name -like "* Link *" | Set-AudioDevice -Verbose
}else{
  echo "JabraLink -> JabraSPEAK"
  Get-AudioDevice -List | where name -like "* SPEAK *" | Set-AudioDevice -Verbose
}
