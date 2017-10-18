copy ..\LogHome\WS-NPDC-MX80-01.txt  .\
#Get Hostname from log file
[string]$hostname = cat WS-NPDC-MX80-01.txt | grep Hostname: | awk '{print $2}' | sed -n '1p'
echo "1 - {Hostname} = $hostname "
#Get Model
[string]$model = cat WS-NPDC-MX80-01.txt | grep Model | awk '{print $2}' | sed -n '1p'
echo "2 - {Model} = $model"
#Get UPtime
[string]$uptime = cat WS-NPDC-MX80-01.txt | grep Uptime | awk '{print $2,$3,$4,$5,$6,$7,$8,$9}' | sed -n '1p'
echo "3 - {Uptime} = $uptime"
#Get Serial number
[string]$serial = cat WS-NPDC-MX80-01.txt | grep 'Chassis' | awk '{print $2}'
echo "4 - {Serial} = $Serial"
[string]$sid = cat WS-NPDC-MX80-01.txt | grep 'Serial ID' | awk '{print $4}'
echo "4 - {SID} = $sid"
#Get CPU Usage
[int]$cpu_user = cat WS-NPDC-MX80-01.txt | grep User | sed -n '1p' | awk '{print $2}'
[int]$cpu_kernel = cat WS-NPDC-MX80-01.txt | grep Kernel | sed -n '3p' | awk '{print $2}'
[int]$cpu = $cpu_user + $cpu_kernel
echo "5 - {CPU_User} = $cpu_user"
echo "6 - {CPU_Kernel} = $cpu_kernel"
echo "7 - {CPU} = $cpu"
#Get memory used
#[int]$MEM_used = cat WS-NPDC-MX80-01.txt | grep 'Processor Pool' | awk '{print $6}'
#echo "5 - {Memory_used} = $MEM_used"
#Get memory total
#[int]$MEM_total= cat WS-NPDC-MX80-01.txt | grep 'Processor Pool' | awk '{print $4}'
#echo "6 - {Memory_total} = $MEM_total"
#caculate memory usage
$MEM_usage = cat WS-NPDC-MX80-01.txt | grep 'Memory utilization' | awk '{print $3}' | sed -n '1p'
echo "8 - {Memory} = $MEM_usage"

$date_check = Get-Date 
switch ($date_check.DayOfWeek)
    {
        �P����{
				$mon = $date_check.AddDays(-1).Month
				$day = $date_check.AddDays(-1).Day
			  }
        �P����{
				$mon = $date_check.AddDays(-2).Month
				$day = $date_check.AddDays(-2).Day
			  }
        �P���@{
				$mon = $date_check.AddDays(-3).Month
				$day = $date_check.AddDays(-3).Day
			  }
        �P���G{
				$mon = $date_check.AddDays(-4).Month
				$day = $date_check.AddDays(-4).Day
			  }
        �P���T{
				$mon = $date_check.AddDays(-5).Month
				$day = $date_check.AddDays(-5).Day
			  }
        �P���|{
				$mon = $date_check.AddDays(-6).Month
				$day = $date_check.AddDays(-6).Day
			  }
        �P����{
				$mon = $date_check.AddDays(-0).Month
				$day = $date_check.AddDays(-0).Day
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
$date = "106  �~  $mon  ��  $day  ��"
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
echo "9 - {count} = $count"	
$filename1="Juniper MX80-01_(��"  
$filename2='��).doc'
$filename = $filename1 + $count + $filename2
$CurrentPath = Get-Location
$folderPath = $CurrentPath.Path  
$fileType =  $filename          

echo "Get the Template file for Replacement"
copy ..\Template\temp_MX80.doc .\$filename 

$textToReplace = @{
	"{count}" = $count
	"{date}"  = $date
    "{Hostname}" = $hostname
	"{Model}" = $model
	"{Uptime}" = $uptime
	"{Serial}" = $serial
	"{SID}" = $sid
	"{CPU}"    = $cpu
	"{CPU_User}" = $cpu_user
	"{CPU_Kernel}" = $cpu_kernel
	#"{Memory_used}" = $MEM_used 
	#"{Memory_total}" = $MEM_total
	"{Memory}" = $MEM_usage

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


