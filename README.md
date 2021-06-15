#Exemplo de integração de VBA com linguagem C com .BAT e com windons powershell<BR>

Renomeio o arquivo excel para o nome free .  Exemplo  planilha_de_compras.xslx renomeie para free.xslx<BR>
copie o arquivo renomeado para a pasta desbloqueia_arquivo<BR>
Aperte o .bat<BR>
Demora um pouco, eficiencia O(n³).<BR>

Unlock excel.<BR>

Rename to free. Drap and drop excel file to "Desbloquear_excel" and click on the .bat . Few minutes later your excel worksheet with password will be unlock.<BR>

Work only for windows.<BR>


Logica do algoritmo e o que ele faz:<BR>
1)Muda a extensão de .xlsx para zip<BR>
2)Descompactac o zip e cria uma pasta gabrielbgab com os arquivos<BR>
3)Executa o algoritmo em c que le no diretorio gabrielbgab/xl/worksheets/  os arquivos sheet1.xml,sheet2.xml..ect<BR>
3)Desbloquea alterando o arquivo sheet1.xml, sheet2.xml...etc.. e copia para a pasta atual<BR>
4)Move os arquivos modificados para a pasta gabrielbgab/xl/worksheets/<BR>
5)compacta em .zip a pasta gabrielbgab<BR>
6)muda o nome de .zip para free.xls<BR>
7)a planilha está com as worksheets desblouqueadas<BR>


