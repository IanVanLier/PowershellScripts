#Function to enable RDP on another pc, pass in a .csv .txt or a single pc name and wether to enable or disable them for rdp
function SetRemotePcRDP([parameter(Mandatory=$true)][string]$RemoteComputer,[parameter(Mandatory=$true)][int]$EnableOrDisable)
{
    #Tests if we passed in a path (txt or csv) or a single pc name
    if((Test-Path $RemoteComputer)) {
         Write-Host "passed in a valid path"

          Get-Content $RemoteComputer |  Foreach-Object {          
          RemoteRDPHandler -RemoteComputer $_ -EnableOrDisable $EnableOrDisable
          }

    }
    else{
        RemoteRDPHandler -RemoteComputer $RemoteComputer -EnableOrDisable $EnableOrDisable
    }
} 

function RemoteRDPHandler($RemoteComputer, $EnableOrDisable){
     write-host "Changing RDP settings on $RemoteComputer" -ForegroundColor Yellow
     
     if ($EnableOrDisable -eq 0)
        {

         Invoke-Command –Computername $RemoteComputer {
    
            # Enables RDP on remote ps (value 1 to disable, 0 to enable)	
            Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0
            # Enables RDP traffic firewall rule	
            Enable-NetFirewallRule -DisplayGroup "Remote Desktop"                  

            } 
              Write-Host "Enabled RDP $RemoteComputer `n" -ForegroundColor Green
    
         }
        else{
        Invoke-Command –Computername $RemoteComputer {
            # Enables RDP on remote ps (value 1 to disable, 0 to enable)	
            Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 1
            # Disable RDP traffic firewall rule	
            Disable-NetFirewallRule -DisplayGroup "Remote Desktop"

            }

            
            Write-Host "Disabled RDP $RemoteComputer `n" -ForegroundColor Red
        }

}

function PrintPcNames($InTextFile){

if(!(Test-Path $InTextFile)) {
    Write-Host "passed in a single pc name"}


else {
    Write-Host "Passed in text file"

    
    Get-Content $InTextFile |  Foreach-Object {
    Write-Host "$_"
    }
}

}




