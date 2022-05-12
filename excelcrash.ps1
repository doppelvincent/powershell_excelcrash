#$xl = New-Object -ComObject Excel.Application

#$xl.Visible = $true

#$xl.Workbooks.Open("C:\Users\vincent.vincent\Documents\Powershell\2021_Tax_26-04-2022_v1al.xlsb")
Function Test-FileLock {
    Param(
        [parameter(Mandatory=$True)]
        [string]$Path
    )
    $OFile = New-Object System.IO.FileInfo $Path
    If ((Test-Path -Path $Path -PathType Leaf) -eq $False) {Return $False}
    Else {
        Try {
            $OStream = $OFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            If ($OStream) {$OStream.Close()}
            Return $False
        } 
        Catch {Return $True}
    }
}


$folderPath = "H:\PS"    
$fileName = Get-ChildItem -Path $folderPath -File -Filter *.xlsb -Recurse

$excl = New-Object -ComObject "Excel.Application"
while ($true) {
    $filePath = Join-Path -Path $folderPath -ChildPath $fileName
    $isOpen = Test-FileLock($filePath)
    while ($isOpen) {
        Start-Sleep -s 5
        $isOpen = Test-FileLock($filePath)
    }

    $excl.Visible = $true
    $workbook_hehe = $excl.Workbooks.Open($filePath)
    Start-Sleep -s 5
    $worksheet = $workbook_hehe.worksheets.item('Dashboard für ZG im CI')
    Write-Output $worksheet.name
    New-ItemProperty -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security -Name "VBAWarnings" -Value "1" -PropertyType DWORD -Force | Out-Null
    try {
        $excl.Run("EverythingInOne")
    }
    catch {
        $isOpen = Test-FileLock($filePath)
        "ERROR"
        if ($isOpen -eq "$false") {
            
        }
    }
}