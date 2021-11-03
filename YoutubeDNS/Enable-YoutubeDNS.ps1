<#
.Synopsis
   This module removes a forward lookup zone for youtube.com and tiktok.com in the local DNS server
.DESCRIPTION
   This module removes a forward lookup zone for youtube.com and tiktok.com in the local DNS server
.EXAMPLE
   Disable-YoutubeDNS.ps1
.NOTES
#>

$DNSServer = "pdc2.ross-wa.net"
$AddressList = @("youtube.com","tiktok.com")

Import-Module dnsserver

if (Test-NetConnection -ComputerName $DNSServer -CommonTCPPort WINRM -InformationLevel Quiet) {
    $CurrentZones = Get-DnsServerZone -ComputerName $DNSServer
    Foreach ($Address in $AddressList) {
        if ($CurrentZones.ZoneName -like $Address){
            Write-Host "Deleting DNS Zone $Address on $DNSServer"
            Remove-DnsServerZone -ComputerName $DNSServer -Name $Address -Force
        }
        else{
            Write-Host "Zone $Address is already deleted on $DNSServer"
        }
    }
}