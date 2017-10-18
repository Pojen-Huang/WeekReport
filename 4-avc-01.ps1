#Get the Source Data

#Get Hostname from log file
#[string]$hostname =  cat E7-NPDC-GTM-01.txt | grep hostname | awk   '{print $2}'
#echo "1 - E7-NPDC-FAS3220-02  = $hostname "
#Get CPU Usage
#[string]$cpu = cat E7-NPDC-GTM-01.txt | grep -A2 'System CPU Usage(%' | awk '{print $3}' | sed -n '3p'
#echo "4 - {CPU} = $cpu"
#Get memory used
#[int]$MEM_used = cat WS-NPDC-3750X-07.txt | grep 'Processor Pool' | awk '{print $6}'
#echo "5 - {Memory_Used} = $MEM_used"
#Get memory total
#[int]$MEM_total= cat WS-NPDC-3750X-07.txt | grep 'Processor Pool' | awk '{print $4}'
#echo "6 - {Memory_total} = $MEM_total"
#caculate memory usage
#[int]$MEM_usage = ($MEM_used/$MEM_total)*100
#echo "7 - {Memory_usage_percent} = $MEM_usage"

$date_check = Get-Date 
switch ($date_check.DayOfWeek)
    {
        星期六{
				$mon = $date_check.AddDays(-1).Month
				$day = $date_check.AddDays(-1).Day
			  }
        星期日{
				$mon = $date_check.AddDays(-2).Month
				$day = $date_check.AddDays(-2).Day
			  }
        星期一{
				$mon = $date_check.AddDays(-3).Month
				$day = $date_check.AddDays(-3).Day
			  }
        星期二{
				$mon = $date_check.AddDays(-4).Month
				$day = $date_check.AddDays(-4).Day
			  }
        星期三{
				$mon = $date_check.AddDays(-5).Month
				$day = $date_check.AddDays(-5).Day
			  }
        星期四{
				$mon = $date_check.AddDays(-6).Month
				$day = $date_check.AddDays(-6).Day
			  }
        星期五{
				$mon = $date_check.AddDays(-0).Month
				$day = $date_check.AddDays(-0).Day女
			  }
		Saturday{
				$mon = $date_check.AddDays(-1).Month
				$day = $date_check.AddDays(-1).Day
			  }
        Sunday{
				$mon = $date_check.AddDays(-2).Month
				$day = $date_check.AddDays(-2).Day
			  }
        Monday{
				$mon = $date_check.AddDays(-3).Month
				$day = $date_check.AddDays(-3).Day
			  }
	    Tuesday{
				$mon = $date_check.AddDays(-4).Month
				$day = $date_check.AddDays(-4).Day
			  }
        Wednesday{
				$mon = $date_check.AddDays(-5).Month
				$day = $date_check.AddDays(-5).Day
			  }
        Thursday{
				$mon = $date_check.AddDays(-6).Month
				$day = $date_check.AddDays(-6).Day
			  }
        Friday{
				$mon = $date_check.AddDays(-0).Month
				$day = $date_check.AddDays(-0).Day
			  }
    
    }
$date = "106  年  $mon  月  $day  日"
echo "8 - {date} = $date"

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

#Get Path

$filename ="DeepSecurity_avc_01.doc"
$CurrentPath = Get-Location
$folderPath = $CurrentPath.Path  
$fileType = "working.doc"        
echo "9 - {count} = $count"	
copy ..\Template\temp_$filename .\$filename   

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


