# C-programing

Unlock excel.

Rename to free. Drap and drop excel file to "Desbloquear_excel" and click on the .bat . Few minutes later your excel worksheet with password will be unlock.

Work only for windows.


CODE of .bat file


 dir /b | Findstr  "free" >inicio.txt 

 set /p bruno=<inicio.txt

echo %bruno%
ren %bruno% arquivo.zip

if exist *.zip goto existe
if not exist *.zip goto nao
:existe
%arquivo%=*.zip
echo arquivo enconrtado
echo %arquivo%
goto fim

:nao

echo deu pau fodeu!
:fim

echo %cd%>caminho.txt
 dir /b | Findstr  ".zip" >ar.txt 


powershell.exe -nologo -noprofile -command "&" { Add-Type -A 'System.IO.Compression.FileSystem';$bruno=cat ar.txt ;[IO.Compression.ZipFile]::ExtractToDirectory($bruno, 'gabrielbgab'); }"






Desbloquear_planilhas_Excel.exe  %cd%\gabrielbgab\xl\worksheets %cd%



powershell.exe -nologo -noprofile -command "&" { Add-Type -A 'System.IO.Compression.FileSystem';$bruna=cat inicio.txt ;[IO.Compression.ZipFile]::CreateFromDirectory('gabrielbgab', 'gabrielbgab.zip'); }"




 set /p bruno=<inicio.txt
echo %bruno%
ren gabrielbgab.zip %bruno% 

Echo Planilha Desbloqueada com sucesso!!
pause


