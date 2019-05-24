SET project_path="C:\Program Files (x86)\AorusFusion\AorusFusionMod"

cd %project_path%
pyinstaller AorusFusionMod.py --noupx --onefile --noconsole --distpath %project_path%

:: pause
