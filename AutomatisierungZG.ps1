$folderPath = "C:\Users\Guest\Desktop\Powershell_AutomatisierungZG"    
$fileName = Get-ChildItem -Path $folderPath -File -Filter *.xlsb -Recurse


$filePath = Join-Path -Path $folderPath -ChildPath $fileName

Function OpenAndRunZG {
    Param([string]$Path)
    $excl = New-Object -ComObject Excel.Application
    #$excl.Visible = $true
    #$excl.Workbooks.Open($Path)
    Invoke-Item $Path
    Start-Sleep -s 5
    New-ItemProperty -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security -Name "VBAWarnings" -Value "1" -PropertyType DWORD -Force | Out-Null
    Start-Sleep -s 20
    $popup = New-Object -ComObject wscript.shell
    $popup.AppActivate("Excel")
    $popup.SendKeys("{ESC}")
    Start-Sleep -s 60
    $worksheet = $excl.Worksheets.item('Dashboard für ZG im CI').Activate()
    Write-Output $worksheet.name
    Start-Sleep -s 3
    $excl.Run("EverythingInOne")

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("ALLES IST FERTIG!! Das Excel ist $counter Mal abgestürzt!", "SteuerCHECK", 0, [System.Windows.Forms.MessageBoxIcon]::Information)

}
Function OpenAndRunZG_WD{
    Param([string]$Path)
    $excl = New-Object -ComObject Excel.Application
    #$excl.Visible = $true
    #$excl.Workbooks.Open($Path)
    Invoke-Item $Path
    Start-Sleep -s 5
    New-ItemProperty -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security -Name "VBAWarnings" -Value "1" -PropertyType DWORD -Force | Out-Null
    Start-Sleep -s 20
    $popup = New-Object -ComObject wscript.shell
    $popup.AppActivate("Excel")
    $popup.SendKeys("{ESC}")

    Start-Sleep -s 60

    $worksheet = $excl.Worksheets.item('Dashboard für ZG im CI').Activate()
    Write-Output $worksheet.name

    Start-Sleep -s 3

    $excl.Run("EverythingInOne")

    Start-Sleep -s 180
    Write-Host ("DER ZG IST DURCHGELAUFEN")

    $worksheet = $excl.Worksheets.item('WEBDRIVER RESULT')
    For ($i = 7;;$i++){
        $text = $worksheet.Cells(4,$i).Text
        if ($text -eq ""){
            $lastcolumn = $i - 1
            break
        }
    
    }
    Start-Sleep -s 3
    $anzahl = $lastcolumn - 3
    $worksheet.Cells.Item(2,2) = "4"
    $worksheet.Cells.Item(3,2) = "$anzahl"
    Write-Host ("DER WEBDRIVER WIRD IN 30 SEKUNDEN BEGINNEN")
    Start-Sleep -s 30
    $excl.Run("RunWebDriverCode")

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("ALLES IST FERTIG!! Das Excel ist $counter Mal abgestürzt!", "SteuerCHECK", 0, [System.Windows.Forms.MessageBoxIcon]::Information)

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
Remove-Item -Path "C:\Users\Guest\Desktop\doku_*.txt"
$counter = 0
for(;;){
    try{
        $isLocked = Test-FileLock($filePath)
        
        If (!($isLocked)){
            Start-Sleep -s 10
    
            $proc = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
            If ($userInput -eq "a"){
                If ($proc){
                    $proc.Kill()
                    $counter += 1
                    Write-Host $counter
                    OpenAndRunZG($filePath)
                    Break
                }
                Else {
                OpenAndRunZG($filePath) 
                Break
                }
            }
            Elseif ($userInput -eq "b"){
                If ($proc) {
                    $proc.Kill()
                    $counter += 1
                    Write-Host $counter
                    OpenAndRunZG_WD($filePath)
                    Break
                }
                Else {
                    OpenAndRunZG_WD($filePath)
                    Break
                }
            }

                
            }
        Else {
            $proc = Get-Process -Name EXCEL | Sort-Object -Property ProcessName -Unique -ErrorAction SilentlyContinue
            $i = 1
            while ($true){
                if ($proc.Responding -eq $false){
                    $i += 1
                    if ($i -eq 15){
                        $proc.Kill()
                        break
                    }
                }
                else {
                    break
                }
                Start-Sleep -s 20
            }
        }


        }
    catch{    
    }
    Start-sleep -s 5
    
}
