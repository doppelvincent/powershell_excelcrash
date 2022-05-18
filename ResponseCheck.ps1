while ($true){
  $proc = Get-Process -Name EXCEL | Sort-Object -Property ProcessName -Unique -ErrorAction SilentlyContinue
  if ($proc.Responding -eq $false){
      Write-Host "Excel is not responding $i"
      $i += 1
      if ($i -eq 1200){
          $proc.Kill()
          break
      }
  }
  else {
      break
  }
  Start-Sleep -s 1
}