@echo off
title Python Flask Server
echo Dang khoi dong Flask Server trong mang LAN...

:: 1. Di chuyen den thu muc du an (Thay o: bang o dia cua ban)
cd /d "C:\Cham_cong_Ngay_phep_Ngoai_gio"

:: 2. Kich hoat moi truong ao (venv)
call venv\Scripts\activate

:: 3. Chay Flask app
python app.py