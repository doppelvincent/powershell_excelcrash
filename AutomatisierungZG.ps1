$folderPath = "C:\Users\vincent.vincent\Documents\Powershell"    
$fileName = Get-ChildItem -Precurath $folderPath -File -Filter *.xlsb -Recurse


$filePath = Join-Path -Path $folderPath -ChildPath $fileName
Function OpenAndRunZG {
    Param([string]$Path)
    $excl = New-Object -ComObject Excel.Application
    Start-Sleep -s 10
    Invoke-Item $Path
    New-ItemProperty -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security -Name "VBAWarnings" -Value "1" -PropertyType DWORD -Force | Out-Null
    Start-Sleep -s 20
    $popup = New-Object -ComObject wscript.shell
    $popup.AppActivate("Excel")
    $popup.SendKeys("{ESC}")
    Start-Sleep -s 60
    $worksheet = $excl.Worksheets.item('Dashboard für ZG im CI').Activate()
    Write-Output $worksheet.name
    $excl.Run("EverythingInOne")
}

Function OpenAndRunZG_WD{
    Param([string]$Path)
    $excl = New-Object -ComObject Excel.Application
    Start-Sleep -s 10
    Invoke-Item $Path
    New-ItemProperty -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security -Name "VBAWarnings" -Value "1" -PropertyType DWORD -Force | Out-Null
    Start-Sleep -s 20
    $popup = New-Object -ComObject wscript.shell
    $popup.AppActivate("Excel")
    $popup.SendKeys("{ESC}")
    Start-Sleep -s 60
    $worksheet = $excl.Worksheets.item('Dashboard für ZG im CI').Activate()
    Write-Output $worksheet.name
    $excl.Run("EverythingInOne")
    Start-Sleep -s 120
    $worksheet = $excl.Worksheets.item('WEBDRIVER RESULT')
    For ($i = 7;;$i++){
    $text = $worksheet.Cells(4,$i).Text
    if ($text -eq ""){
        $lastcolumn = $i - 1
        break
    }

    }
    $anzahl = $lastcolumn - 3
    $worksheet.Cells.Item(2,2) = "4"
    $worksheet.Cells.Item(3,2) = "$anzahl"
    Start-Sleep -s 5
    $excl.Worksheets.item('WEBDRIVER RESULT').Activate
    Start-Sleep -s 30
    $excl.Run("RunWebDriverCode")
}
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

$userInput = Read-Host "[a] Zufallsgenerator [b] Zufallsgenerator + Webdriver"
for(;;){
try{
    $isLocked = Test-FileLock($filePath)
    
    If (!(Get-Process -Name excel -ErrorAction SilentlyContinue) -or (!($isLocked))){
        If ($userInput -eq "a"){
            OpenAndRunZG($filePath)
        }
        elseif ($userInput -eq "b"){
           OpenAndRunZG_WD($filePath) 
        }
        }
    
    }
catch{    
}
    Start-sleep -s 5
}