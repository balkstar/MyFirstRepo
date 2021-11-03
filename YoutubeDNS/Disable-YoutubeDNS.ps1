<#
.Synopsis
   This module creates a forward lookup zone for youtube.com and tiktok.com in the local DNS server
.DESCRIPTION
   This module creates a forward lookup zone for youtube.com and tiktok.com in the local DNS server
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
        if ($CurrentZones.ZoneName -notlike $Address){
            Write-Host "Creating DNS Zone $Address on $DNSServer"
            Add-DnsServerPrimaryZone -ComputerName $DNSServer -Name $Address -ZoneFile "$Address.DNS"
        }
        else{
            Write-Host "Zone $Address already exists on $DNSServer"
        }
    }
}