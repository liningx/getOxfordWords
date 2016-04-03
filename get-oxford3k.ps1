function downOxford($uri) {
    $html = Invoke-WebRequest $uri
    $wordlist = $html.ParsedHtml.getElementById("entrylist1").outertext

    $wordlistInArray = @()
    $wordlist.split("`r") | %{ $wordlistInArray += $_}

    $array_20_5 = New-Object 'object[,]' 20,5
    $wordlistInArray | %{ 
        $index = [array]::IndexOf($wordlistInArray,$_)
        $rownum = $index % 20
        $colnum = [math]::floor($index / 20)
        $array_20_5[$rownum, $colnum] = $_.trim()
    }

    $outtxt = @();
    for($i = 0; $i -lt 20; $i++){
        $rowcom = ""
        for($j=0; $j -lt 5; $j++){
            $cellValue = $array_20_5[$i, $j]
            if ($cellValue -ne $null){
                $cellValue = $cellValue.trim()    
            }

            if ($j -eq 0) { 
                $rowcom = $cellValue        
            } else {
                $rowcom += "," + $cellValue    
            }
        }
        $outtxt += $rowcom
    }

    return $outtxt
}


$preUri = 'http://www.oxfordlearnersdictionaries.com/us/wordlist/english/oxford3000'

$fullDownload=@()
$toBeDown = @{"A-B"="5"; "C-D"="6"; "E-G"="5"; 
       "H-K"="4"; "L-N"="4"; "O-P"="5"; 
       "Q-R"="3"; "S"="5"; "T"="3"; "U-Z"="3"}
$toBeDown.GetEnumerator() | sort key | % {
    for($page = 1; $page -le $_.value; $page++){
        $uri = $preUri + "/Oxford3000_$($_.key)/?page=$($page)"
        write-host "downloading > $uri"
        $pageDown = downOxford $uri
        $fullDownload += $pageDown + "$($_.key),=,=,=,=" 
    }
}

$fullDownload > oxford3k.csv
write-host "done."
