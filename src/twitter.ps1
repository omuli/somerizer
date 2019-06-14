#constants
$url_search_atrsoft='https://twitter.com/search?q=%23atrsoft%20since%3A2019-01-01'
$url_search_customtools='https://twitter.com/search?q=%23customtools%20since%3A2019-01-01'

$excelFileName = "desktop\out{0}.xlsx" -f (Get-Date -Format HH:mm:ss.fff)
$outputpath = join-path -Path $env:USERPROFILE -ChildPath $excelFileName


#create excel file
$excelFile = New-Object -ComObject excel.application
$excelFile.visible = $true

$excelWorkbook = $excelFile.Workbooks.Add()

$excelWorksheet= $excelWorkbook.Worksheets.Item(1) 
$excelWorksheet.Name = 'Twitter'

#create excel header row
$excelWorksheet.Cells.Item(1,1)="Handle"
$excelWorksheet.Cells.Item(1,2)="Name"
$excelWorksheet.Cells.Item(1,3)="Link"
$excelWorksheet.Cells.Item(1,4)="Retweets"
$excelWorksheet.Cells.Item(1,5)="Likes"
$excelWorksheet.Cells.Item(1,6)="Message"

$excelWorksheet.Columns("F").ColumnWidth = 75

$rowNbr = 2;

#invoke web request
$webResponseObject = Invoke-WebRequest -Uri $url_search_customtools

$webResponseObject.AllElements | ForEach-Object {
    if ($_.tagName = "div") {
        # looking for tweets with the class name "original-tweet"
        if ($_.class -Match 'original-tweet') {
            # "found tweet"
            $excelWorksheet.Cells.Item($rowNbr,1)="@" + $_.'data-screen-name'
            $excelWorksheet.Cells.Item($rowNbr,2)=$_.'data-name'
            $excelWorksheet.Hyperlinks.Add(
                $excelWorksheet.Cells.Item($rowNbr,3),
                    "https://twitter.com" + $_.'data-permalink-path') | Out-Null
            $excelWorksheet.Cells.Item($rowNbr,6)=$_.innerText
            # likes: <SPAN class="ProfileTweet-action--favorite u-hiddenVisually"><SPAN class=ProfileTweet-actionCount data-tweet-stat-count="2">
            # retweets: <SPAN class="ProfileTweet-action--retweet u-hiddenVisually"><SPAN class=ProfileTweet-actionCount data-tweet-stat-count="1">
            $rowNbr++
            }
        }
    }


#save Excel Workbook
$excelWorkbook.SaveAs($outputpath)
$excelFile.Quit()