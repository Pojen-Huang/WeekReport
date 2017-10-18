copy ..\LogHome\WS-NPDC-FAS3220-06.txt .\RAWDATA
#Get Hostname from log file
$hostname = cat .\RAWDATA | grep WS-NPDC-FAS3220-06> | sed -n '15p' | sed 's/>//'
echo "1 - {Hostname} = $hostname "
#Get Model Name
$model = cat .\RAWDATA | grep 'Model Name:' | sed -n '1p' | awk '{print $3}'
echo "2 - {Model} = $model"
#Get Memory size
$MEM_size = cat .\RAWDATA | grep 'Memory Size:' | awk '{print $3,$4}'
echo "3 - {Memory} = $MEM_size"
#Get UPtime
$uptime = cat .\RAWDATA | grep  -A1 uptime | sed -n '2p' | awk '{print $3,$4,$5}'
echo "4 - {Uptime} = $uptime"
#Get Serial number
$serial = cat .\RAWDATA | grep 'System Serial Number:' | awk '{print $4}'
echo "5 - {Serial} =  $serial"

##FAS HardDisk Usage##

$aggr0 =  cat .\RAWDATA | grep aggr0 | sed -n '1p' | awk '{print $5}'
$aggr1 =  cat .\RAWDATA | grep aggr1 | sed -n '1p' | awk '{print $5}' 
$OSW1_dr = cat .\RAWDATA | grep "/vol/dr_OS_windows_01/" | sed -n '1p' | awk '{print $5}'
$OSW2_dr = cat .\RAWDATA | grep "/vol/dr_OS_windows_02/" | sed -n '1p' | awk '{print $5}'
$OSL1_dr = cat .\RAWDATA | grep "/vol/dr_OS_linux_01/" | sed -n '1p' | awk '{print $5}'
$OSL2_dr = cat .\RAWDATA | grep "/vol/dr_OS_linux_02/" | sed -n '1p' | awk '{print $5}'
$EWDR = cat .\RAWDATA | grep "dr_EMIC_WD" | sed -n '1p' | awk '{print $5}'
$MSPDR = cat .\RAWDATA | grep "dr_MSP_Monitor_FS" | sed -n '1p' | awk '{print $5}'
$ARCDBDR = cat .\RAWDATA | grep "dr_arc_db" | sed -n '1p' | awk '{print $5}'
$EMCDXDR = cat .\RAWDATA | grep "dr_EMIC_CDX" | sed -n '1p' | awk '{print $5}'
$ARCGIS = cat .\RAWDATA | grep "dr_EMIC_ArcGIS" | sed -n '1p' | awk '{print $5}'
$SRMPLACE = cat .\RAWDATA | grep "srm_placehold" | sed -n '1p' | awk '{print $5}'
$FORSRMPLACE = cat .\RAWDATA | grep "for_srm_placehold" | sed -n '1p' | awk '{print $5}'
echo "6 - {aggr0} = $aggr0"
echo "7 - {aggr1} = $aggr1"
echo "8 - {OSW1_dr} = $OSW1_dr"
echo "9 - {OSW2_dr} = $OSW2_dr"
echo "10 - {OSL1_dr} = $OSL1_dr"
echo "11 - {OSL2_dr} = $OSL2_dr"
echo "12 - {EWDR} = $EWDR"
echo "13 - {MSPDR} = $MSPDR"
echo "14 - {ARCDBDR} = $ARCDBDR"
echo "15 - {EMCDXDR} = $EMCDXDR"
echo "16 - {ARCGIS} = $ARCGIS"
echo "17 - {FORSRMPLACE} =$FORSRMPLACE"
echo "18 - {SRMPLACE} =$SRMPLACE"
del RAWDATA
$date_check = Get-Date 
switch ($date_check.DayOfWeek)
    {
        星期六{
				$year = $date_check.AdDaysd(-1).Year
                $mon = $date_check.Add(-1).Month
				$day = $date_check.AddDays(-1).Day
			  }
        星期日{
				$year = $date_check.AddDays(-2)
                $mon = $date_check.AddDays(-2).Month
				$day = $date_check.AddDays(-2).Day
			  }
        星期一{
				$mon = $date_check.AddDays(-3).Month
				$day = $date_check.AddDays(-3).Day
                $year = $date_check.AddDays(-3).Year
			  }
        星期二{
				$mon = $date_check.AddDays(-4).Month
				$day = $date_check.AddDays(-4).Day
                $year = $date_check.AddDays(-4).Year
			  }
        星期三{
				$mon = $date_check.AddDays(-5).Month
				$day = $date_check.AddDays(-5).Day
                $year = $date_check.AddDays(-5).Year
			  }
        星期四{
				$mon = $date_check.AddDays(-6).Month
                $year = $date_check.AddDaysDays(-6).Year
				$day = $date_check.AddDays(-6).Day
			  }
        星期五{
				$mon = $date_check.AddDays(-0).Month
				$day = $date_check.AddDays(-0).Day
                $year = $date_check.AddDays(-0)
			  }
		Friday{
                $year= $date_check.AddDays(-0).Year
                $mon = $date_check.AddDays(-0).Month
                $day = $date_check.AddDays(-0).Day
                }
        Saturday{
                $year= $date_check.AddDays(-1).Year
                $mon = $date_check.AddDays(-1).Month
                $day = $date_check.AddDays(-1).Day
                    }
                                Sunday{
            $year= $date_check.AddDays(-2).Year
            $mon = $date_check.AddDays(-2).Month
            $day = $date_check.AddDays(-2).Day
                }
                            Monday{
            $year= $date_check.AddDays(-3).Year
            $mon = $date_check.AddDays(-3).Month
            $day = $date_check.AddDays(-3).Day
                }
                            Tuesday{
            $year= $date_check.AddDays(-4).Year
            $mon = $date_check.AddDays(-4).Month
            $day = $date_check.AddDays(-4).Day
                }
                            Wednesday{
            $year= $date_check.AddDays(-5).Year
            $mon = $date_check.AddDays(-5).Month
            $day = $date_check.AddDays(-5).Day
                }
                            Thursday{
            $year= $date_check.AddDays(-6).Year
            $mon = $date_check.AddDays(-6).Month
            $day = $date_check.AddDays(-6).Day
                }
    
    }
$date = "106  年  $mon  月  $day  日"
echo "18 - {date} = $date"

$count_check = Get-Date -uFormat %A
switch ($count_check)
    {
        星期六{$count_adjust = 170}
        星期日{$count_adjust = 170}
        星期一{$count_adjust = 169}
        星期二{$count_adjust = 169}
        星期三{$count_adjust = 169}
        星期四{$count_adjust = 169}
        星期五{$count_adjust = 169}
        Saturday{$count_adjust = 170}
        Sunday{$count_adjust = 170}
        Monday{$count_adjust = 169} 
		Tuesday{$count_adjust = 169}
        Wednesday{$count_adjust = 169}
        Thursday{$count_adjust = 169}
        Friday{$count_adjust = 169}
    }
$count = $count_adjust + (Get-Date -uFormat %W)
echo "19 - {count} = $count"	


$filename1='WS-FAS3220-06_(第'
$filename2='次).doc'
$foldername1 = "EMIC週報$year$mon$day(第"
$foldername2 ='次)'
$OutFile= $filename1 + $count + $filename2
$OutFolder="..\ReportHome\$foldername1$count$foldername2"

Copy-Item ..\Template\temp_FAS3220-06.doc .\working.doc
$CurrentPath = Get-Location
$folderPath = $CurrentPath.Path  
$fileType =  "working.doc"          

echo "Get the Template file for Replacement"
copy ..\Template\temp_FAS3220-06.doc .\working.doc 
$textToReplace = @{
	"{count}" = $count
	"{date}"  = $date
    "{Hostname}" = $hostname
	"{Model}"  = $Model
	"{Memory}"  = $MEM_size
	"{Uptime}" = $uptime
	"{Serial}" = $serial
	#"{CPU}"    = $cpu
	#"{CPU_Users}"    = $cpu_user
	#"{CPU_Kernel}}"    = $cpu_kernel
	#"{Memory_used}" = $MEM_used 
	#"{Memory_total}" = $MEM_total
	#"{Memory_usage_percent}" = $MEM_usage
	"{aggr0}" = $aggr0
	"{aggr1}" = $aggr1
	#"{mgmt_dr}" =$mgmt_dr
	#"{Sybaseg_Min}"  = $Sybaseg_Min
	#"{Sybaseg_Max}"  = $Sybaseg_Max
    "{OSW1_dr}" = $OSW1_dr
    "{OSW2_dr}" = $OSW2_dr
    "{OSL1_dr}" = $OSL1_dr
    "{OSL2_dr}" = $OSL2_dr
    "{EWDR}" = $EWDR
    "{MSPDR}" = $MSPDR
    "{ARCDBDR}" = $ARCDBDR
    "{EMCDXDR}" = $EMCDXDR
    "{ARCGIS}" = $ARCGIS
    "{SRMPLACE}" =$SRMPLACE
	"{FORSRMPLACE}" =$FORSRMPLACE
	"{redundantNode}" = "WS-NPDC-FAS3220-05"
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false

$matchCase = $true
$matchWholeWord = $false
$matchWildcards = $false
$matchSoundsLike = $false
$matchAllWordForms = $false
$forward = $true
$findWrap = [Microsoft.Office.Interop.Word.WdReplace]::wdReplaceAll
$format = $false
$replace = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindContinue

$countf = 0 #count files
$countr = 0 #count replacements

Function findAndReplace($objFind, $FindText, $ReplaceWith) {
    #simple Find and Replace to execute on a Find object
    $objFind.Execute($FindText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $findWrap, $format, $ReplaceWith, $replace)
}

Function findAndReplaceAll($objFind, $FindText, $ReplaceWith) {
    $count = 0
    $count += findAndReplace $objFind $FindText $ReplaceWith
    While ($objFind.Found) {
        $count += findAndReplace $objFind $FindText $ReplaceWith
    }
    return $count
}

Function findAndReplaceMultiple($objFind, $lookupTable) {
    #apply multiple Find and Replace on the same Find object
    $count = 0
    $lookupTable.GetEnumerator() | ForEach-Object {
        $count += findAndReplaceAll $objFind $_.Key $_.Value
    }
    return $count
}

Function findAndReplaceMultipleWholeDoc($Document, $lookupTable) {
    $count = 0
    # Loop through each StoryRange
    ForEach ($storyRge in $Document.StoryRanges) {
        Do {
            $count += findAndReplaceMultiple $storyRge.Find $lookupTable
            #check for linked Ranges
            $storyRge = $storyRge.NextStoryRange
        } Until (!$storyRge) #non-null is True

    }
    # Loop through each Section's Headers and Footers Shapes
    ForEach ($section in $Document.Sections) {
        # https://msdn.microsoft.com/en-us/vba/word-vba/articles/shapes-object-word
        # "The Count property for this collection in a document returns the number of items in the main story only.
        #  To count the shapes in all the headers and footers, use the Shapes collection with any HeaderFooter object."
        # Hence the .Item(1) which should be able to collect all Shapes
        If ($section.Headers.Item(1).Shapes.Count) {
            ForEach ($shp in $section.Headers.Item(1).Shapes) {
                If ($shp.TextFrame.HasText) {
                    $count += findAndReplaceMultiple $shp.TextFrame.TextRange.Find $lookupTable
                }
            }
        }
        If ($section.Footers.Item(1).Shapes.Count) {
            ForEach ($shp in $section.Footers.Item(1).Shapes) {
                If ($shp.TextFrame.HasText) {
                    $count += findAndReplaceMultiple $shp.TextFrame.TextRange.Find $lookupTable
                }
            }
        }
    }
    return $count
}

Function processDoc {
    $count = 0
    $doc = $word.Documents.Open($_.FullName)
    $count += findAndReplaceMultipleWholeDoc $doc $textToReplace
    $doc.Close([ref]$true)
    return $count
}

$sw = [Diagnostics.Stopwatch]::StartNew()
Get-ChildItem -Path $folderPath -Recurse -Filter $fileType | ForEach-Object { 
    $countr = 0
    Write-Host "Processing \`"$($_.Name)\`"..."
    $countr += processDoc
    Write-Host "$countr replacements made."
    $countf++
}
$sw.Stop()
$elapsed = $sw.Elapsed.toString()
Write-Host "`nDone. $countf files processed in $elapsed"

$word.Quit()
$word = $null
[gc]::collect() 
[gc]::WaitForPendingFinalizers()

 
copy .\working.doc   $OutFolder\$OutFile
remove-item .\working.doc


