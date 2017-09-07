Import-Module ".\Convert-Office365NetworksData\ConvertFrom-O365AddressesXMLFile\ConvertFrom-O365AddressesXMLFile.ps1" -Force
$a = ConvertFrom-O365AddressesXMLFile -RemoveFileAfterParsing

# TODO:
# - Fix maximum length on FQDN's
# - Parameter
# - List available Services, Types
# - delete duplicate IP Addresses
# - integration with FortiManager API

# Target Formatation:
# edit G-EXTERN-MS-O365-UPDATE-110117-137.116.157.126/32
# set subnet 137.116.157.126/32
# next
#config firewall addrgrp
#    edit "G-INTERN-GRP-CAOP"
#        set member "G-INTERN-CAOP-Prod" "G-INTERN-CAOP-Test" "G-INTERN-svr-sflext-01" "G-INTERN-ATWPSFLEX01" "G-INTERN-CAOP-Test01"
#    next
#end


## Parameters

# Maximum support addresses per group all Fortigate models: 300
$addrGrpLimit = 250

# Name of address object
$StartStringAddressObject = "G-EXTERN-MS-O365"

# Name of group object
$StartStringGroupObject = "G-EXTERN-GRP-O365"

# Address Object types (IPv4, IPV6, URL)
$selectedTypes = "IPv6"
#$selectedTypes = ($a.Type | Sort-Object | unique)

# O365 Service selected (CRLs, EOP, OfficeIpad, Identity, etc.)
$selectedServices = ($a.Service | Sort-Object | unique)




$o365AddressObjects = @(@{},@{},@{})

$FortiIPV4Text = "config firewall address`n"
$FortiIPV6Text = "config firewall address6`n"
$FortiIPFQDNText = "config firewall address`n"

ForEach ($Service in $selectedServices) {

    $temparray_v4 = New-Object System.Collections.ArrayList
    $temparray_v6 = New-Object System.Collections.ArrayList
    $temparray_fqdn = New-Object System.Collections.ArrayList

    ForEach ($object in ($a | Where-Object {$_.Service -eq $Service})) {
        

        if ($object.Type -eq 'IPv4' -and ($selectedTypes -contains $object.Type)) {

            $obj_ip = "$($object.IPAddress)/$($object.SubnetMaskLength)"

            $FortiIPV4Text += "edit $StartStringAddressObject-$obj_ip`n"
            $FortiIPV4Text += "set subnet $obj_ip`n"
            $FortiIPV4Text += "set comment $($object.Service)`n"
            $FortiIPV4Text += "next`n"

            $temparray_v4.Add("$StartStringAddressObject-$obj_ip") >$null

        }


        if ($object.Type -eq 'IPv6' -and ($selectedTypes -contains $object.Type)) {

            $obj_ip = "$($object.IPAddress)/$($object.SubnetMaskLength)"
            
            $FortiIPV6Text += "edit $StartStringAddressObject-$obj_ip`n"
            $FortiIPV6Text += "set ip6 $obj_ip`n"
            $FortiIPV6Text += "set comment $($object.Service)`n"
            $FortiIPV6Text += "next`n"

            $temparray_v6.Add("$StartStringAddressObject-$obj_ip") >$null
        }

        if ($object.Type -eq 'URL' -and ($selectedTypes -contains $object.Type)) {

            $obj_fqdn = "$($object.Url)"
            
            $FortiIPFQDNText += "edit $StartStringAddressObject-$obj_fqdn`n"
            $FortiIPFQDNText += "set type wildcard-fqdn`n"
            $FortiIPFQDNText += "set wildcard-fqdn $obj_fqdn`n"
            $FortiIPFQDNText += "set comment $($object.Service)`n"
            $FortiIPFQDNText += "next`n"

            $temparray_fqdn.Add("$StartStringAddressObject-$obj_fqdn") >$null
        }
    }
    if ($temparray_v4.Count){
        $o365AddressObjects[0].Add($Service, $temparray_v4)}
    if ($temparray_v6.Count){
        $o365AddressObjects[1].Add($Service, $temparray_v6)}
    if ($temparray_fqdn.Count){
        $o365AddressObjects[2].Add($Service, $temparray_fqdn)}

}

$FortiIPV4Text += "end`n"
$FortiIPV6Text += "end`n"
$FortiIPFQDNText += "end`n"

$FortiIPV4Text
$FortiIPV6Text
$FortiIPFQDNText


for ($i=0;$i -le 2;$i++){

    if ($o365AddressObjects[$i].keys -eq 0){ continue};

switch ($i)
{
    0 {write-output "config firewall addrgrp"}
    1 {write-output "config firewall addrgrp6"}
    2 {write-output "config firewall addrgrp"}
}

Foreach ($serviceType in $o365AddressObjects[$i].keys) {

    $tempstring = ""

    if ($o365AddressObjects[$i][$serviceType].Count -eq 0){ continue};

    switch ($i){
        0 {write-output "edit $StartStringGroupObject-$serviceType"}
        1 {write-output "edit $StartStringGroupObject-$serviceType"}
        2 {write-output "edit $StartStringGroupObject-$serviceType-FQDN"}
    }

    Foreach ($element in $o365AddressObjects[$i][$serviceType]) {

        $counter++;

        $tempstring += "`"$element`" "


        if ($counter % $addrGrpLimit -eq 0) {

            write-output "set member $tempstring"
            write-output "next"
            $groupcounter++
            write-output "edit $StartStringGroupObject-$serviceType-$groupcounter"
            $tempstring = ""
        }

    }

    if ($tempstring.Length -gt 0) {
        "set member $tempstring"
    }

    "next"
    $counter = 0
    $groupcounter = 0

}
"end"
}