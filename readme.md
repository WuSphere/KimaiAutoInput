## exeファイルを生成
```powershell
pyinstaller --onefile --hidden-import pandas --hidden-import proxy_utils --add-data "config.json;." --add-data "勤怠データ.xlsx;." KimaiAutoInput.py
```