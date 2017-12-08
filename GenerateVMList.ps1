function GenerateVMList([parameter(Mandatory=$true)][string]$VMHostList ,[parameter(Mandatory=$true)][string]$ExportPath){

$VMHostCredentials = (get-credential)

    if(Test-Path $VMHostList){
    
        Get-Content $VMHostList |  Foreach-Object {
        Connect-VIServer -Server $_ -Credential $VMHostCredentials  
        }       
         
    $ExportVmList =  Get-VM | Select -Property VMHost, name ,@{N='IP';E={[string]::Join(',',$_.Guest.IPAddress)}},@{N="DNS";E={[string]::Join(',',($_.ExtensionData.Guest.IpStack.DnsConfig.IpAddress))}} 
    $ExportVmList |Export-Excel -NoNumberConversion IP, DNS $ExportPath
    }
}
