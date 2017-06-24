# C-programing
Desbloquear planilhas excel

Renomeio o arquivo excel para o nome free .  Exemplo  planilha_de_compras.xslx renomeie para free.xslx
copie o arquivo renomeado para a pasta desbloqueia_arquivo
Aperte o .bat
Demora um pouco, eficiencia O(n³).

Unlock excel.

Rename to free. Drap and drop excel file to "Desbloquear_excel" and click on the .bat . Few minutes later your excel worksheet with password will be unlock.

Work only for windows.


Logica do algoritmo e o que ele faz:
1)Muda a extensão de .xlsx para zip
2)Descompactac o zip e cria uma pasta gabrielbgab com os arquivos
3)Executa o algoritmo em c que le no diretorio gabrielbgab/xl/worksheets/  os arquivos sheet1.xml,sheet2.xml..ect
3)Desbloquea alterando o arquivo sheet1.xml, sheet2.xml...etc.. e copia para a pasta atual
4)Move os arquivos modificados para a pasta gabrielbgab/xl/worksheets/
5)compacta em .zip a pasta gabrielbgab
6)muda o nome de .zip para free.xls
7)a planilha está com as worksheets desblouqueadas


