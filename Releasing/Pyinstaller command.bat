set "Project_Basedir=%~dp0"
set "Script_Path=%Project_Basedir%..\Codebase\Mariser_Script.py"


pyinstaller --onefile --add-data "%Project_Basedir%\..\Assets\Images;Assets\Images" --noconsole %Script_Path%