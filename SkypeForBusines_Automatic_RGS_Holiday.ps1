### Skype for Business Automatic RGS Holidayset ###
### Version 1.0                                 ###
### Author: Alexander Holmeset                  ###
### Email: alexander.holmeset@gmail.com         ###
### Twitter: twitter.com/Holmez85               ###
### Blog: skype4bworld.wordpress.com            ###


# This script catch holidays from officeholidays.com and creates a RGS Holiday Set for Skype for Business.
# I got the inspiration from Chris Hayward script, that could get the dates from a JSON website and set them as RGS Holiday Sets
# http://chrishayward.co.uk/2017/07/07/skype-for-business-automatically-set-rgs-holiday-sets-with-powershell-and-json/
# The problem is that there is no service that for free provides JSON holiday info.
# On officeholidays.com you have updated holiday information for 133 countries.
# The dates are placed in tables. This script finds the tables, and creates powershell variables out of the data.
#
# Comments: The dates have 3 types of classes: holiday, publicholiday and regional. 
# As there is no standard for how to mark holidays if its "a day off from work" holiday, you should adjust this script to suit your needs.
# You can do this by removing the holiday class you want from the Where-Object command. You could also remove days you dont need with Jamie Schwinn´s tool: http://waveformation.com/holidayseteditor/
# Will take a look to see if i can automate removing days. Next step will be to see if this can be converted to work with autoatetendants in O365.
#
# Inside the two functions you can see what are valid countries.
#
# Other sources of inspiration:
# Ken Lasko: http://ucken.blogspot.no/2012/05/holiday-sets-for-lync-response-groups.html
# James Arber: http://www.skype4badmin.com/skype4b-and-lync-tools/
# Paul Bloem: https://ucsorted.com/2012/11/20/rgs-hours-of-business-holiday-sets/
# Yoav Barzilay: https://y0av.me/2016/01/05/irish-public-holidays-2016-set-for-skype-for-business-server/
# Andrew Morpeth: https://ucgeek.co/2013/12/lync-response-group-holiday-sets/

$PoolName = "FE01.contoso.local"
$HolidaySetName = "Norwegian Holidays"
$Country = "Norway"

function Mostcountries {

        param(
        [Parameter(Position=0)]
        [ValidateSet("Algeria","Angola","Armenia","Argentina","Australia","Austria","Azerbaijan","Bahamas","Bahrain","Bangladesh","Barbados","Belarus","Belgium","Bolivia","Bosnia_and_Herzegovina","Botswana","Brazil","Brunei","Bulgaria","Burundi","Cambodia","Canada","Cayman_Islands","Chile","China","Colombia","Costa_Rica",
        "Croatia","Cyprus","Czech_Republic","Denmark","Dominican_Republic","Ecuador","Egypt","El_Salvador","Estonia","Ethiopia","Fiji","Finland","France","Georgia","Germany","Ghana","Gibraltar","Greece","Grenada","Guernsey",
        "Honduras","Hong_Kong","Hungary","Iceland","India","Indonesia","Iraq","Ireland","Isle_of_Man","Israel","Italy","Jamaica","Japan","Jersey","Jordan","Kazakhstan","Kenya","Kuwait","Lao",
        "Latvia","Lebanon","Libya","Liechtenstein","Lithuania","Luxembourg","Macau","Macedonia","Maldives","Malta","Mauritius","Mexico","Moldova","Monaco","Montenegro",
        "Malaysia","Morocco","Mozambique","Myanmar","Netherlands","New_Zealand","Nigeria","Norway","Oman","Pakistan","Panama","Paraguay","Peru","Philippines","Poland","Portugal","Qatar",
        "Romania","Russia","Rwanda","Saint_Lucia","Saudi_Arabia","Serbia","Singapore","Slovakia","Slovenia","South_Africa","South_Korea","Spain","Sri_Lanka","Sweden",
        "Switzerland","Taiwan","Tanzania","Thailand","Tonga","Trinidad_and_Tobago","Tunisia","Turkey","Turks_and_Caicos_Islands","Uganda","Uganda","Ukraine",
        "United_Arab_Emirates","United_Kingdom","Uruguay","USA","Venezuela","Vietnam","Yemen","Zambia","Zimbabwe")]
        [System.String]$Country
        )

$uri = "http://www.officeholidays.com/countries/$Country/index.php"
$html = Invoke-WebRequest -Uri $uri
$tables = $html.ParsedHtml.getElementsByTagName('tr') |
Where-Object {$_.classname -eq 'holiday' -or $_.classname -eq 'regional' -or $_.classname -eq 'publicholiday' } |
Select-Object -exp innerHTML
$script:holidays = foreach ($table In $tables){ 
$day= (($table -split "<TD>")[1] -split "</TD>")[0] ;

$Date = (($table -split "<SPAN class=ad_head_728>")[1] -split "</SPAN>")[0]; 

$Holiday = ((($table -split "<TD><A title=")[1] -split ">")[1] -split "</A")[0]
$Remarks = (($table -split "class=remarks>")[1] -split "<")[0]; 
[PSCustomObject]@{
        Date = [DateTime]$Date  ; Holiday = $Holiday
        Remarks = If ($Holiday -ne $Remarks){$Remarks}}

 }
}

function Fewcountries {


    param(
        [Parameter(Position=0)]
        [ValidateSet("Albania","Switzerland")]
        [System.String]$Country
        )


$uri = "http://www.officeholidays.com/countries/$Country/index.php"
$html = Invoke-WebRequest -Uri $uri
$table = $html.ParsedHtml.getElementsByTagName('tr') |
Where-Object {$_.classname -eq 'holiday' -or $_.classname -eq 'regional' -or $_.classname -eq 'publicholiday'} |
Select-Object -exp innerHTML

$script:holidays = foreach ($t in $table){
    [void]($t -match 'SPAN title="(.*?)"') ; $m1 = $Matches[1]
    [void]($t -match 'tooltip>(.*)') ; $m2 = $Matches[1]
  .  [void]($t -match 'remarks>(.*) ') ; $m3 = $Matches[1]
    [PSCustomObject]{
        Date = $m1 ; Holiday = $m2
        Remarks = If ($m2 -ne $m3){$m3}}
}
}




#


$ErrorActionPreference = "SilentlyContinue"




Mostcountries -country $Country
Fewcountries -country $Country

# Get the Holiday Set
$y = Try {Get-CsRgsHolidaySet -Identity "service:ApplicationServer:$PoolName" -Name $HolidaySetName}catch{$null}
 
# Clear the current Holiday List if it exists (Stops duplicates)
If ($y.name -eq $HolidaySetName){
    $y.HolidayList.Clear()
    Set-CsRgsHolidaySet -Instance $y
}
 
ForEach ($hol in $holidays){
 
    $StartDate = $hol.date | get-date -Format yyyy-MM-dd
    $EndDate = $([DateTime]$hol.date).AddDays(1)
    
    # Ignore any previous public holidays
    If (([DateTime]$hol.date) -gt $(get-date -Format yyyy-MM-dd)){
 
        $x = New-CsRgsHoliday -StartDate $StartDate -EndDate $EndDate -Name $hol.holiday
 
        # If the current Holiday set exists add day to Holiday List
        If ($y.name -eq $HolidaySetName){
            $y.HolidayList.Add($x)
            Set-CsRgsHolidaySet -Instance $y
        }Else{
        # Create the holiday set
            New-CsRgsHolidaySet -Parent "service:ApplicationServer:$PoolName" -Name $HolidaySetName -HolidayList($x)
        }
    }
}


#Lists the holidays added. 
Get-CsRgsHolidaySet | Select-Object -ExpandProperty HolidayList




