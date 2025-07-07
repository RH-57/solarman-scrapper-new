@echo off
echo Wait until VPN is On...
timeout /t 45 /nobreak

cd /d "C:\WPy64-31241\notebooks\Scrape Solarman Mariadb"
start "" "C:\WPy64-31241\python-3.12.4.amd64\python.exe" solarman_scrape_new.py
exit