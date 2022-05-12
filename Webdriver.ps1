$folderPath = "C:\Users\vincent.vincent\Documents\Powershell"    
$fileName = Get-ChildItem -Path $folderPath -File -Filter *.xlsb -Recurse
$filePath = Join-Path -Path $folderPath -ChildPath $fileName
$excl = New-Object -ComObject Excel.Application
$excl.Visible = $true
$excl.Workbooks.Open($filePath)
Start-Sleep -s 5
$worksheet = $excl.Worksheets.item("WEBDRIVER RESULT")
Start-Sleep -s 5
For ($i = 7;;$i++){
    $text = $worksheet.Cells(4,$i).Text
    if ($text -eq ""){
        $lastcolumn = $i - 1
        break
    }
}

$test = $lastcolumn - 3
$worksheet.Cells.Item(2,1) = "4"
$worksheet.Cells.Item(3,2) = "$test"
$excl.Worksheets.item("WEBDRIVER RESULT").Activate()


Write-Host $lastcolumn