copy-item ..\LogHome\E7-NPDC-FG1000-2.txt  .\RAWDATA
$sn=0 #list number start
#Get Uptime
$uptime = cat RAWDATA | grep Uptime | sed -n '1p' 
echo "($sn+1) - {Uptime} = $uptime"

#Get CPU Usage
$cpu_user =  cat RAWDATA  | grep 'CPU states' | awk '{print $3}' | sed 's/%//'
$cpu_system =  cat RAWDATA  | grep 'CPU states' | awk '{print $5}' | sed 's/%//'
$cpu = $cpu_use + $cpu_system
echo "($sn+1) - {CPU} = $CPU"

#Get Memory Usage
$MEM_usage = cat RAWDATA  | grep 'Memory states:' | awk '{print $3}'
echo "($sn+1) - {Memory} = $MEM_usage"


$date_check = Get-Date 
switch ($date_check.DayOfWeek)
    {
        �P����{
				$year = $date_check.AdDaysd(-1).Year
                $mon = $date_check.Add(-1).Month
				$day = $date_check.AddDays(-1).Day
			  }
        �P����{
				$year = $date_check.AddDays(-2)
                $mon = $date_check.AddDays(-2).Month
				$day = $date_check.AddDays(-2).Day
			  }
        �P���@{
				$mon = $date_check.AddDays(-3).Month
				$day = $date_check.AddDays(-3).Day
                $year = $date_check.AddDays(-3).Year
			  }
        �P���G{
				$mon = $date_check.AddDays(-4).Month
				$day = $date_check.AddDays(-4).Day
                $year = $date_check.AddDays(-4).Year
			  }
        �P���T{
				$mon = $date_check.AddDays(-5).Month
				$day = $date_check.AddDays(-5).Day
                $year = $date_check.AddDays(-5).Year
			  }
        �P���|{
				$mon = $date_check.AddDays(-6).Month
                $year = $date_check.AddDaysDays(-6).Year
				$day = $date_check.AddDays(-6).Day
			  }
        �P����{
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
$date = "  106  �~  $mon  ��  $day  ��"
echo "8 - {date} = $date"

$count_check = Get-Date -uFormat %A
switch ($count_check)
    {
        �P����{$count_adjust = 170}
        �P����{$count_adjust = 170}
        �P���@{$count_adjust = 169}
        �P���G{$count_adjust = 169}
        �P���T{$count_adjust = 169}
        �P���|{$count_adjust = 169}
        �P����{$count_adjust = 169}
        Saturday{$count_adjust = 170}
        Sunday{$count_adjust = 170}
        Monday{$count_adjust = 169}
        Tuesday{$count_adjust = 169}
		Wednesday{$count_adjust = 169}
        Thursday{$count_adjust = 169}
        Friday{$count_adjust = 169}
    }
$count = $count_adjust + (Get-Date -uFormat %W)

#Get Path
$filename1="Fortinet FG-02-1000C_(��"
$filename2='��).doc'
$filename = $filename1 + $count + $filename2
$CurrentPath = Get-Location
$folderPath = $CurrentPath.Path  
$fileType = $filename         
echo "9 - {count} = $count"	
copy ..\Template\temp_FG-02.doc .\$filename   

$textToReplace = @{
	"{count}" = $count
	"{date}"  = $date

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


