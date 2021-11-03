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
$AddressList - @("youtube.com","tiktok.com")

Import-Module dnsserver

if (Test-NetConnection -ComputerName $DNSServer -CommonTCPPort WINRM -InformationLevel Quiet) {
    $CurrentZones = Get-DnsServerZone -ComputerName $DNSServer
    Foreach ($Address in $AddressList) {
        if ($Address -notlike $CurrentZones.ZoneName){
            Write-Host "Creating DNS Zone $Address on $Server"
            Add-DnsServerPrimaryZone -ComputerName $DNSServer -Name $Address -ZoneFile "$Address.DNS" -whatif
        }
    }
}