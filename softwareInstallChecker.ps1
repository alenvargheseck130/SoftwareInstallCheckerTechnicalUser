####################################################################
#BUSINESS USER SCRIPT STARTS HERE
####################################################################

Write-Host "Script Executed"
#Destiantion folder to download installer for all the software
$username = [Environment]::UserName
$destination = "C:\Users\$username\Desktop\"
Install-Module PSWindowsUpdate
$driverName = Read-host "Please enter the disk name of the USB(in a single alphabet): "
$userInput =  $driverName.ToUpper()
$businessInstallerDestination = $userInput+":\InstallerScript\SoftwareInstallCheckerBusinessUser";
$technicalInstallerDestination = $userInput+":\InstallerScript\SoftwareInstallCheckerTechnicalUser";

#FOR TEAMS
$exeName = "teams.exe"
$software = "Teams Machine-Wide Installer";
$installed = $null -ne (Get-WmiObject -Class Win32_product | Where-Object { $_.Name -eq $software }) 
[Net.ServicePointManager]::SecurityProtocol +='tls12'
#Condition to check whether or not the machine has the software installed or not
If(-Not $installed) {
	$source = "https://go.microsoft.com/fwlink/p/?LinkID=2187327&clcid=0x1009&culture=en-ca&country=CA"
	Write-Host "'$software' is NOT installed.";
	#Download software
	Invoke-WebRequest -URI $source -OutFile $destination$exeName
	#install software
	Start-Process $destination$exeName
	Start-Sleep -Seconds 10
	
} 
else {
	Write-Host "'$software' is already installed"
	
}

#FOR GOOGLE CHROME
#sort to find chrome:- Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName | Format-Table -AutoSize
$exeName = "googleChrome.exe"
$software = "Google Chrome";
$installed =  $null -ne (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName | Where-Object { $_.DisplayName -eq "Google Chrome" })  

#Condition to check whether or not the machine has the software installed or not
If(-Not $installed) {
	$source = "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B7E39B4A3-F9C3-D714-548A-A3D317C937C5%7D%26lang%3Den%26browser%3D4%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26ap%3Dx64-stable-statsdef_1%26installdataindex%3Dempty/update2/installers/ChromeSetup.exe"
	Write-Host "'$software' is NOT installed.";
	#Download software
	Invoke-WebRequest -URI $source -OutFile $destination$exeName
	#install software
	Start-Process $destination$exeName
	Start-Sleep -Seconds 10
} 
else {
	Write-Host "'$software' is already installed"
	
}


#FOR GOOGLE DRIVE
#sort to find google Drive:- Get-PSDrive -PSProvider FileSystem | Select-Object Root
 #get-wmiobject -class win32_logicaldisk | Select-Object DeviceID 
 $installerExeName = "\GoogleDriveSetup.exe"
 $installerPath = $businessInstallerDestination+$installerExeName
 $software = "Google Drive"
 #method to verify if the google driver exists in the local machine disk
 $installed = $null -ne (Get-WmiObject -Class win32_logicaldisk  | Select-Object DeviceID | Where-Object { $_.DeviceID -eq "G:" }) 
 
 #Condition to check whether or not the machine has the software installed or not
 If(-Not $installed) {
	Write-Host "'$software' is NOT installed.";
	#install software
	Start-Process $installerPath
	Start-Sleep -Seconds 10
} 
 else {
	 Write-Host "'$software' is already installed"
}


#FOR OFFICE APPS
#to sort = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$exeName = "microsoftOffice.exe"
$software = "Microsoft Office"
$uninstallKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$O365Exists = $null -ne ($uninstallKeys | Where-Object { $_.GetValue("DisplayName") -eq "Office 16 Click-to-Run Extensibility Component" })
if ($O365Exists) {
Write-Host "'$software' is already installed"

}
else {
	$source = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365ProPlusRetail&platform=Def&language=en-us&TaxRegion=pr&correlationId=224dcf7a-8131-486f-9832-9e4dc2336cdf&token=049fdd00-51a9-42b8-9328-f2159c510d3a&version=O16GA&source=O15OLSO365&Br=2"
	Write-Host "'$software' is NOT installed.";
	#Download software
	Invoke-WebRequest -URI $source -OutFile $destination$exeName
	#install software
	Start-Process $destination$exeName
	Start-Sleep -Seconds 10
}

#FOR WINDOWS ACTIVATION

#method to check if windows is installed - 
#(Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%'" | 
#Where-Object { $_.PartialProductKey } | select Description, LicenseStatus | select #LicenseStatus | Where-Object {$_.LicenseStatus -eq "1"}) -ne $null
$software = "Windows"
$windowsActivated = $null -ne (Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%'" | Where-Object { $_.PartialProductKey } | Select-Object Description, LicenseStatus | Select-Object LicenseStatus | Where-Object {$_.LicenseStatus -eq "1"})
if($windowsActivated){
	Write-Host "'$software' is already activated"
}
else{
	slmgr.vbs /ipk HKGWV-79N29-F89Y8-VQT7H-XD72F
}
	<# alternative #1
	$computer = gc env:computername
	$key = "HKGWV-79N29-F89Y8-VQT7H-XD72F"
	$service = get-wmiObject -query "select * from SoftwareLicensingService" -computername $computer
	$service.InstallProductKey($key)
	$service.RefreshLicenseStatus()
	Write-Host "Windows is now activated on your system...."
	#>
####################################################################
#BUSINESS USER SCRIPT ENDS HERE
####################################################################

##
##
##
##
##

####################################################################
#TECHNICAL USER SCRIPT STARTS HERE
####################################################################

#FOR NOTEPAD++ 
$exeName = "notepad++.exe"
$software = "Notepad++ (64-bit x64)";
$installed = $null -ne (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq $software }) 

If(-Not $installed) {
	$source = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.4.8/npp.8.4.8.Installer.x64.exe"
	Write-Host "'$software' is NOT installed.";
	#Download software
	Invoke-WebRequest -URI $source -OutFile $destination$exeName
	#install software
	Start-Process $destination$exeName
	Start-Sleep -Seconds 10
	
} 
 else {
	Write-Host "'$software' is already installed."
	
}

#FOR VIPACCESS
$exeName = "vipAccess.exe"
$software = "VIP Access"
$installed = $null -ne (Get-WmiObject -Class Win32_product  | Select-Object Name | Where-Object { $_.Name -eq "VIP Access" })

If(-Not $installed) {
	$source = "https://storage.googleapis.com/sedvip-prd-idcenter-downloads/VIPAccessSetup.exe"
	Write-Host "'$software' is NOT installed.";
	#Download software
	Invoke-WebRequest -URI $source -OutFile $destination$exeName
	#install software
	Start-Process $destination$exeName
	Start-Sleep -Seconds 10
	
} 
 else {
	Write-Host "'$software' is already installed."
	
}

#FOR CISCO ANYCONNECT SECURE MOBILITY CLIENT
$installerExeName = "\anyconnect-win-4.10.00093-core-vpn-predeploy-k9.msi"
$installerPath = $technicalInstallerDestination+$installerExeName
$software = "Cisco AnyConnect Secure Mobility Client"
 #method to verify if the google driver exists in the local machine disk
$installed = $null -ne (Get-WmiObject -Class Win32_product  | Select-Object Name | Where-Object { $_.Name -eq "Cisco AnyConnect Secure Mobility Client" })

If(-Not $installed) {
	$source = "https://doc-0g-84-docs.googleusercontent.com/docs/securesc/pm9mojsjoaki32du80l2q7v7ethi201i/5buknag1pagmana83t55e7jsfnnssdng/1673966625000/13035833597927717980/13035833597927717980/1-oCu4ZbpNKZsaU3bZgThrgIGj9-CclqA?e=download&ax=ALjR8sxUsl7DHpLz-06QMx11i1vxP9BjPagUGCO-ZGCQ4JFKILzUhA08wsOhyjUgngUI2ldNe48fjVOLh0HmC_PIXUHjug_kI1bgzN2K0gZ_YV_yWtQ60Y1Z40AVvmFt0e9AC84x45HqEZOU3ZILnhvZ07vTVa55zDaSyiiz1qOghV5aFQPCs_IiQ7L1awnll5pmEwOuJM-EW2Tm9yP19l4p-RtNlR6f6xI5GhsH4e-zdqtc7PoGkn0ZrmVcu4fO-CYnopJ1NJULwuSMQ4CL-dZMPGwYiIbVmIMYkED7HxKa9mTAfG4EoqrBWizTvfQ2crT5ePscjmNx3eWAmWfy-bVWtn5hUW3kSOL_0k2SCLjsyfU1UarNTG6nvUZmBaIKsTl61SixbWnUVWG-kHS-feoOrxUglJ_nZ8bPSUtGGnyeKGauf0RDyhr8p3t8llOZH4KJ3wVCvWzXODvG1bIDb8VFEncwA9FSlSNKnFGo7FzM2ziVw5R81aPtMmsxaOiQg_IJKl08W0L3uD4T5aE04OK8YPaQqIvQd94PsdYFLmpc9VFKEVHjGZam6p0Q4BnshCbc30uh2YnSzhP8MSodaGRC6Z0QueoCgx4Mv-xFX2zUV49heu-ryfMQrVadcFDJqZcbLMet6Sr1cY2NLTfdXA9NfxdwZiMJaPkekiHrPTqNoPZbJCfitqa06Q-G13woQaK71OgevqdSnY0doVGU7lqvyhvKIWjCt9NVeHjZUP4AXAtUNnohlD3W7OV32uWo9ffxhTPbOChjneHXMgKAbUzfnizH3f_T7n1J0vPTWl5OLfDfAat0Q20yvVn0NgAqkeQZSX83eyh1Afqo89DXuj5MZ5m09D2AYAJQqG9_KJlZfOPdQlORrokiwUD61SMk9fI&uuid=587da7b1-385e-4f57-b974-930d862b28eb&authuser=0&nonce=dtqqb3hno8e5a&user=13035833597927717980&hash=et903huso2ppujqjcdum6gnhlocl1bru"
	Write-Host "'$software' is NOT installed.";
	#install software
	Start-Process $installerPath
	Start-Sleep -Seconds 10
}
 else {

	Write-Host "'$software' is already installed."
	
}

#FOR SSMS
$installerExeName = "\SSMS-Setup-ENU.exe"
$installerPath = $technicalInstallerDestination+$installerExeName
$software = "SQL Server Management Studio"
$installed = $null -ne (Get-WmiObject -Class Win32_product  | Select-Object Name | Where-Object { $_.Name -eq "SQL Server Management Studio" }) 

If(-Not $installed) {
	Write-Host "'$software' is NOT installed.";
	#install software
	Start-Process $installerPath
	Start-Sleep -Seconds 10
	
} 
 else {
	Write-Host "'$software' is already installed."
}


#FOR DOCKER DESKTOP
# need to be fixed
#Method for checking docker ->
#foreach($obj in $InstalledSoftware){write-host $obj.GetValue('DisplayName')}
$installerExeName = "\Docker Desktop Installer.exe"
$installerPath = $technicalInstallerDestination+$installerExeName
$software = "Docker Desktop"
$InstalledSoftware = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
$installed = foreach($obj in $InstalledSoftware){$obj.GetValue('DisplayName')} Where-Object { $_.DisplayName -eq "Docker Desktop"}
$veryfying = $null -ne (Where-Object {$installed -eq $software})
#$source = "https://desktop.docker.com/win/main/amd64/Docker%20Desktop%20Installer.exe"
If(-Not $veryfying) {
	Write-Host "'$software' is NOT installed.";
	#install software
	Start-Process $installerPath
	Start-Sleep -Seconds 10
	
} 
 else {
	Write-Host "'$software' is already installed."	
} 

#FOR VSCODE
#name in powershell - Microsoft Visual Studio Tools for Applications 2017 x64 Hosting Support
#WSL extemnsion install = code --install-extension ms-vscode-remote.remote-wsl 
$exeName = "vscode.exe"
$software = "Visual Studio Code"
$installed = $null -ne (Get-WmiObject -Class Win32_product  | Select-Object Name | Where-Object { $_.Name -eq "Microsoft Visual Studio Tools for Applications 2017 x64 Hosting Support" }) 

If(-Not $installed) {
	$source = "https://az764295.vo.msecnd.net/stable/97dec172d3256f8ca4bfb2143f3f76b503ca0534/VSCodeUserSetup-x64-1.74.3.exe"
	Write-Host "'$software' is NOT installed.";
	#Download software
	Invoke-WebRequest -URI $source -OutFile $destination$exeName
	#install software
	Start-Process $destination$exeName
	Start-Sleep -Seconds 10
	
} 
	Timeout /NoBreak 30
	if ($installed) {
		#method to list all extensions= code --list-extensionsh
		#method to install extension= code --install-extenstion ms-vscode-remote.remote-wsl
		code --install-extension ms-vscode-remote.remote-wsl
	}

else {
	Write-Host "'$software' is already installed."
	
}

#FOR WSL2 AND UBUNTU DISTRO
$flag = "true"
If($flag -eq "true"){
	Write-Host "WSL2 installtion process has started"
	wsl --install -d UBUNTU
	Get-WindowsUpdate -AcceptAll -Install -AutoReboot
	Write-Host "Installed windows updates"
	Write-Host "Ctrl+C to cancel the restart"
	Timeout /NoBreak 600
	Restart-Computer
}
else{
	Write-Host "Updates already installed"
}
####################################################################
#TECHNICAL USER SCRIPT ENDS HERE
####################################################################