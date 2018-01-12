#add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
#Download Letsencrypt-win-simple: https://github.com/Lone-Coder/letsencrypt-win-simple/releases
#Make sure the owa site is reachable on http port 80 from outside of your network. (Configure NAT on firewall to IIS Server port 80)

Add-PSsnapin *Exchange* -ErrorAction SilentlyContinue

#Be sure to correctly set variables below.
#Get Url For External site.
$WebAddress = (Get-OWAVirtualDirectory).ExternalUrl.Host

#Be sure to change password in letsencrypt.exe.config
$mypwd = ConvertTo-SecureString -String "SamplePassword" -Force -AsPlainText
$CertifacateDirectory = "C:\LetsEncrypt\Certificates"
$LetsEncryptLocation = "C:\LetsEncrypt"
$IISRootDirectory = "C:\inetpub\wwwroot"
$EmailAddress = "test@example.com"
$ScriptLocation = "C:\LetsEncrypt\ExchangeLetsEncryptCertificate.ps1"
$RenewalTime = 70
# possible parameters: None | IMAP | POP | UM | IIS | SMTP | Federation | UMCallRouter seperate with comma for multiple services.
$ServicesToEnAble = "IIS"


#Creates the directory for validation files.
$ChallengeDir = $IISRootDirectory + "\.well-known\acme-challenge"
New-Item -ItemType Directory -Path $IISRootDirectory\.well-known\acme-challenge -ErrorAction SilentlyContinue
Copy-Item $LetsEncryptLocation\Web_Config.xml $ChallengeDir -ErrorAction SilentlyContinue

#Disables SSL for the challenge directory
Set-WebConfiguration //system.webserver/security/access -metadata overrideMode -value Allow -PSPath IIS:/
Set-WebConfiguration -PSPath "IIS:\Sites\Default Web Site\.well-known\acme-challenge" -Filter 'system.webserver/security/access' -Value 'None'

#Enable anonymous authentication on challenge directory
Set-WebConfiguration system.webServer/security/authentication/anonymousAuthentication -metadata overrideMode -value Allow -PSPath IIS:/
Set-WebConfigurationProperty -Filter "/system.webServer/security/authentication/windowsAuthentication" -Name Enabled -Value True -PSPath "IIS:\Sites\Default Web Site\.well-known\acme-challenge"
Set-WebConfiguration system.webServer/security/authentication/anonymousAuthentication -PSPath "IIS:\Sites\Default Web Site\.well-known\acme-challenge" -Location "Default Web Site" -Value @{enabled="True"}

#Create subdirectory to store new certificate
$CertifacateDirectory =  $CertifacateDirectory + "\" +$((Get-Date).ToString('yyyy-MM-dd'))
New-Item -ItemType Directory -Path $CertifacateDirectory -Force
write-host "Created directory: " $CertifacateDirectory -foregroundcolor "Yellow"


cd $LetsEncryptLocation 
.\LetsEncrypt.exe --verbose --notaskscheduler --centralsslstore $CertifacateDirectory --webroot $IISRootDirectory --plugin manual --manualhost $WebAddress [--validationmode http-01] --validation selfhosting --emailaddress $EmailAddress --accepttos --forcerenewal --installation none

Start-Sleep -s 20
$ScheduledTaskName = "Renew Exchange LetsEncrypt Certificate"
Unregister-ScheduledTask -TaskName $ScheduledTaskName -Confirm:$false

if(Test-Path ($CertifacateDirectory + "\" + $WebAddress + ".pfx")){
    
#Store thumbPrint of imported certificate.
    $ThumbPrint = (Import-ExchangeCertificate -FileData ([Byte[]]$(Get-Content -Path ($CertifacateDirectory + "\" + $WebAddress + ".pfx")  -Encoding byte -ReadCount 0)) -Password ($mypwd)).thumbprint
    Enable-ExchangeCertificate -Thumbprint $ThumbPrint -Services $ServicesToEnAble

    #Create scheduled task for the next renewal.
    $NextRenewJob = New-ScheduledTaskAction -Execute Powershell.exe -Command $ScriptLocation
    $NextRenewTime = New-ScheduledTaskTrigger -Once -At ((get-date).AddDays($RenewalTime))

    Register-ScheduledTask -Action $NextRenewJob -Trigger $NextRenewTime -TaskName $ScheduledTaskName -Description "Job to renew the exchange certificate automatically." 
}

#Retry the script tomorrow if there is no certificate generated.
else{    
    $NextRenewJob = New-ScheduledTaskAction -Execute Powershell.exe -Command $ScriptLocation
    $NextRenewTime = New-ScheduledTaskTrigger -Once -At ((get-date).AddDays(1))

    Register-ScheduledTask -Action $NextRenewJob -Trigger $NextRenewTime -TaskName $ScheduledTaskName -Description "Job to renew the exchange certificate automatically." 
}
