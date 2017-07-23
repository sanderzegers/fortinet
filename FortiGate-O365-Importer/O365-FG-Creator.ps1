Import-Module ".\Convert-Office365NetworksData\ConvertFrom-O365AddressesXMLFile\ConvertFrom-O365AddressesXMLFile.ps1" -Force
$a = ConvertFrom-O365AddressesXMLFile -RemoveFileAfterParsing

# TODO:
# - Don't print empty groups
# - IPv6 and Domain Name support
# - Parameter
# - List available Services, Types
# - delete duplicate IP Addresses
# - integration with FortiManager API

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

# Name of address group object
$startString = "G-EXTERN-MS-O365"

# Address Object types (IPv4, IPV6, URL)
$selectedTypes = "IPv4"
#$selectedTypes = ($a.Type | Sort-Object | unique)

# O365 Service selected (CRLs, EOP, OfficeIpad, Identity, etc.)
$selectedServices = ($a.Service | Sort-Object | unique)




$c = @{}

write-output "config firewall address"

ForEach ($Service in $selectedServices) {

    $temparray = New-Object System.Collections.ArrayList

    ForEach ($object in ($a | Where-Object {$_.Service -eq $Service})) {
        

        if ($object.Type -eq 'IPv4' -and ($selectedTypes -contains $object.Type)) {

            #$obj_service = $object.Service
            $obj_ip = "$($object.IPAddress)/$($object.SubnetMaskLength)"

            write-output "edit $startString-$obj_ip"
            write-output "set subnet $obj_ip"
            write-output "next"

            $temparray.Add("$startString-$obj_ip") >$null

        }


        if ($object.Type -eq 'IPv6' -and ($selectedTypes -contains $object.Type)) {

            #$obj_service = $object.Service
            $obj_ip = "$($object.IPAddress)"

            #$obj_service
            $obj_ip
        }

        if ($object.Type -eq 'URL' -and ($selectedTypes -contains $object.Type)) {

            #$obj_service = $object.Service
            $obj_ip = "$($object.Url)"

            #$obj_service
            $obj_ip
        }
    }
    $c.Add($Service, $temparray)
}

write-output "end"

write-output "config firewall addrgrp"

Foreach ($serviceType in $c.keys) {

    $tempstring = ""


    write-output "edit G-EXTERN-GRP-O365-$serviceType"$

    Foreach ($element in $c[$serviceType]) {

        $counter++;

        $tempstring += "`"$element`" "


        if ($counter % $addrGrpLimit -eq 0) {

            write-output "set member $tempstring"
            write-output "next"
            $groupcounter++
            write-output "edit G-EXTERN-GRP-O365-$serviceType-$groupcounter"
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