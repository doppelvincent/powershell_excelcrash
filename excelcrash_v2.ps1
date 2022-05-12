$folderPath = "H:\PS"    
$fileName = Get-ChildItem -Path $folderPath -File -Filter *.xlsb -Recurse


$filePath = Join-Path -Path $folderPath -ChildPath $fileName

Function Kill-PopUp(){
    kill (Get-Event -SourceIdentifier ChildPID).Messagedata
    Get-Job | Stop-Job
    Get-Job | Remove-Job
}

Function OpenAndRun{
    Param([string]$Path)
    $excl = New-Object -ComObject "Excel.Application"
    $excl.Visible = $true
    $excl.Workbooks.Open($Path)
    Start-Sleep -s 5
    #Kill-PopUp
    New-ItemProperty -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security -Name "VBAWarnings" -Value "1" -PropertyType DWORD -Force | Out-Null
    $worksheet = $excl.worksheets.item('Dashboard für ZG im CI').Activate()
    Start-Sleep -s 5
    $excl.Run("EverythingInOne")
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


for(;;){
try{
    $whatisthis = Get-Process -Name excel -ErrorAction SilentlyContinue
    $isLocked = Test-FileLock($filePath)
    If (!(Get-Process -Name excel -ErrorAction SilentlyContinue) -or (!($isLocked))){
        OpenAndRun($filePath)
        }
    $proc = Get-Process -Name excel | Sort-Object -Property ProcessName -Unique -ErrorAction SilentlyContinue
    $test = $proc.Responding
    If (!$proc) {
        $proc.Kill()
        Start-Sleep -s 10
        OpenAndRun($filePath)
        }
    }
catch    {    
}
    Start-sleep -s 10
}