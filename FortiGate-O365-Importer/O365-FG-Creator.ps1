# TODO:
# - Empty members not allowed

# List available Services, Types
# Set Maximum Object limit
# delete duplicate IP Addresses

# edit G-EXTERN-MS-O365-UPDATE-110117-137.116.157.126/32
# set subnet 137.116.157.126/32
# next
#config firewall addrgrp
#    edit "G-INTERN-GRP-CAOP"
#        set member "G-INTERN-CAOP-Prod" "G-INTERN-CAOP-Test" "G-INTERN-svr-sflext-01" "G-INTERN-ATWPSFLEX01" "G-INTERN-CAOP-Test01"
#    next
#end


Import-Module ".\ConvertFrom-O365AddressesXMLFile.ps1" -Force

#$a = ConvertFrom-O365AddressesXMLFile -RemoveFileAfterParsing
$a = ConvertFrom-O365AddressesXMLFile

$addrGrpLimit = 250

$bTypes = ($a.Type |sort | unique)

$bServices = ($a.Service | sort | unique)

$bTypes = "IPv4"
#$bServices = "SPO","RCA"

$startString = "G-EXTERN-MS-O365"

$c = @{}

"config firewall address"

ForEach ($Service in $bServices) {

$temparray = New-Object System.Collections.ArrayList

    ForEach ($object in ($a | Where-Object {$_.Service -eq $Service})){
        

        if ($object.Type -eq 'IPv4' -and ($bTypes -contains $object.Type)){

            $obj_service = $object.Service
            $obj_ip = "$($object.IPAddress)/$($object.SubnetMaskLength)"

            "edit $startString-$obj_ip"
            "set subnet $obj_ip"
            "next"

            $temparray.Add("$startString-$obj_ip") >$null

        }


        if ($object.Type -eq 'IPv6' -and ($bTypes -contains $object.Type)){

            $obj_service = $object.Service
            $obj_ip = "$($object.IPAddress)"

            $obj_service
            $obj_ip
        }

        if ($object.Type -eq 'URL' -and ($bTypes -contains $object.Type)){

            $obj_service = $object.Service
            $obj_ip = "$($object.Url)"

            $obj_service
            $obj_ip
        }

     

    }


    $c.Add($Service,$temparray)
    
    
}

"end"
"config firewall addrgrp"

$cunter=1

Foreach ($serviceType in $c.keys) {

$tempstring =""

"edit G-EXTERN-GRP-O365-$serviceType"

Foreach ($element in $c[$serviceType]){

$counter++;

$tempstring += "`"$element`" "


if ($counter % $addrGrpLimit -eq 0){

"set member $tempstring"
"next"
$groupcounter++
"edit G-EXTERN-GRP-O365-$serviceType-$groupcounter"
$tempstring=""
}

}

if ($tempstring.Length -gt 0)
{
"set member $tempstring"
}

"next"
$counter=0
$groupcounter=0

}

"end"