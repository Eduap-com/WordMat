Este � o ficheiro LEIA-ME para Windows.


Ficheiros bin�rios inclu�dos nesta distribui��o
-----------------------------------------------

A vers�o Windows do pacote do Maxima inclui ficheiros bin�rios
provenientes de outros projectos de Software Livre tamb�m
alojados no Sourceforge.

gcc: 

gcc.exe, cc1.exe e os ficheiros nas direct�rios lib/gcc-lib e
include/ prov�m da vers�o mingw do gcc. Essa vers�o encontra-se
dispon�vel em http://prdownloads.sf.net/mingw/
     

binutils:

as.exe faz parte da vers�o de binutils em mingw
(http://www.mingw.org/), dispon�vel em
http://prdownloads.sf.net/mingw/


gnuplot:

Os ficheiros wgnuplot.exe, wgnuplot.hlp e wgnuplot.mnu prov�m da
distribui��o de gnuplot para Windows, em
http://gnuplot.sourceforge.net


wxMaxima:

Os ficheiros no direct�rio wxMaxima s�o da distribui��o de
wxMaxima para Windows, dispon�vel em http://wxmaxima.sourceforge.net


Interface gr�fica do Maxima e firewalls
---------------------------------------

Por vezes, as interfaces gr�ficas do Maxima (xmaxima ou wxMaxima)
n�o conseguem arrancar o Maxima, ou enviam uma mensagem de fim
do tempo de espera ou n�o obt�m nenhuma resposta aso comandos
do Maxima. O mais prov�vel � que esses problemas sejam causados
por um programa de firewall ou de antiv�rus. A interface gr�fica
do Maxima comunica-se com o motor do Maxima atrav�s de uma porta
(socket). Os programas de antiv�rus e/ou firewall ver�o isso como
um risco de seguran�a e tentar�o bloque�-lo (porque alguns v�rus
costumam tamb�m usar portas para causar estragos).

Para resolver esse problema:

1.  Tente encontrar o painel de controlo do antiv�rus e/ou firewall.

2.  Procure o nome da interface gr�fica do Maxima entre a lista de
    programas bloqueados e desactive esse bloqueio. A interface
    gr�fica poder� aparecer identificada como um programa "Tcl/Tk" 
    (nome da linguagem em que a interface xmaxima est� feita).


Falta de Espa�o para Ambiente
-----------------------------

Quando tentar executar o maxima numa m�quina com Windows 9x,
poder� obter o seguinte erro:

    Espa�o para ambiente esgotado (Out of environment space)

A seguinte dica, no Artigo 230205 da Base de Conhecimento da
Microsoft, poder� ser �til.

CAUSA

Essa situa��o acontece quando n�o tiver suficiente mem�ria para
colocar uma vari�vel de ambiente no ambiente do MS_DOS.

RESOLU��O

Para resolver o problema, aumente o espa�o para ambiente dispon�vel,
por omiss�o, para programas do MS-DOS. Para fazer isso, use algum
dos m�todos seguintes.

Modificar o ambiente para todos os programas em MS-DOS

Para aumentar o valor por omiss�o do espa�o do ambiente para todos
os programas em MS-DOS executados no Windows, edite o comando
Shell no ficheiro Config.sys. Isso � feito com os passos seguintes:

1.  Clique em Come�ar e a seguir em Executar.

2.  Na caixa que aparece, escreva sysedit e carregue em OK.

3.  Clique na janela do Config.sys. 

4.  No in�cio da linha Shell=, escreva REM, e carregue na barra
    de espa�o.

    Se n�o existir a linha Shell=, avance para o passo seguinte.

5.  Carregue na tecla HOME. 

6.  Escreva a linha seguinte, para criar uma nova linha Shell=, e
    a seguir carregue em ENTER:

      SHELL=C:\COMMAND.COM /E:4096 /P 

7.  No menu de Ficheiros, clique em Guardar.

8.  No ficheiro de menu, clique em Sair.

9.  Reinicie o computador.

Modificar o ambiente para um programa em MS-DOS espec�fico

Para modificar o espa�o por omiss�o unicamente para o ambiente do
programa maxima.bat, siga os passos seguintes:

1.  Clique com o bot�o direito no �cone do programa maxima.bat, e
    a seguir clique em Propriedades.

2.  Clique na sec��o de Memoria.

3.  Na caixa de Ambiente Inicial, escreva o n�mero de quilo-bytes (KB)
    que o programa precisa, e a seguir clique em OK.

NOTA: O tamanho m�ximo de memoria que pode destinar � 4096 KB.
