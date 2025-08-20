python -m venv venv
Set-ExecutionPolicy Unrestricted -Scope Process
venv\Scripts\Activate.ps1 
pip install playwright
pip install pyautogui
pip install pytz
pip install xmltodict
playwright install
$origem = (Get-Location).Path
$destino = "$origem\venv\Lib\site-packages"
Move-Item "$origem\Update-Downloader-Xmls.exe" -Destination $destino
Move-Item "$origem\Renomear-Xmls.exe" -Destination $destino
Move-Item "$origem\start.bat" -Destination $destino
Move-Item "$origem\config.ini" -Destination $destino
Move-Item "$origem\log.txt" -Destination $destino