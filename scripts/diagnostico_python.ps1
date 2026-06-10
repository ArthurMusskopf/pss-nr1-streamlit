Write-Host "== Diagnostico Python ==" -ForegroundColor Cyan
Write-Host "PowerShell:" $PSVersionTable.PSVersion
Write-Host "\nwhere.exe python" -ForegroundColor Yellow
where.exe python
Write-Host "\nwhere.exe py" -ForegroundColor Yellow
where.exe py
Write-Host "\npython --version" -ForegroundColor Yellow
python --version
Write-Host "\npy -3 --version" -ForegroundColor Yellow
py -3 --version
