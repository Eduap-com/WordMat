Esse  o arquivo README do Windows.

Arquivos bin�rios inclusos nessa distribui��o
--------------------------------------------

A vers�o do Maxima para Windows inclui arquivos bin�rios 
de outros projetos de c�digo aberto tamb�m hospedados no Sourceforge.

gcc: 

gcc.exe, cc1.exe e os arquivos nos subdiret�rios lib/gcc-lib
e include/ s�o da vers�o mingw do gcc.  Essa vers�o est�
dispon�vel  em http://prdownloads.sf.net/mingw/
     

binutils:

as.exe � o port mingw (http://www.mingw.org/) do binutils
dispon�vel em http://prdownloads.sf.net/mingw/


gnuplot:

Os arquivo wgnuplot.exe, wgnuplot.hlp e wgnuplot.mnu s�o da
distribui��o para Windows do gnuplot localizado em http://gnuplot.sourceforge.net


wxMaxima:

Os arquivos no subdiret�rio wxMaxima s�o da distribui��o para Windows
do wxMaxima dispon�vel em http://wxmaxima.sourceforge.net


Interface Gr�fica do Maxima e firewall
--------------------------------------

Algumas vezes a interface gr�fica do Maxima (xmaxima ou wxMaxima) n�o consegue ativar o Maxima
ou mostra uma mensagem de tempo esgotado ou n�o recebe as respostas dos comandos do Maxima.  
Muito provavelmente o problema � causando pelo firewall e/ou software antivirus.  
A interface gr�fica do Maxima conversa com o mecanismo computacional atrav�s de um socket.  
Antivirus e/ou programas de firewall v�em e tentam bloque�-lo 
(porque alguns programas suspeitos abrem sockets tamb�m).  

Para resolver o problema:

1.  Tente encontrar o painel de controle para o antivirus e/ou firewall. 

2.  Encontre a interface gr�fica do Maxima na lista de programas bloqueados e desabilite
    o bloqueio para o Maxima.  O programa de interface gr�fica pode aparecer como "Tcl/Tk" 
    (o nome da cole��o de ferramentas utilizadas na montagem da interface gr�fica xmaxima).


Out of Environment Space
------------------------

Quando voc� tenta rodar Maxima em uma m�quina Windows 9x
pode aparecer a seguinte mensagem de erro 

    Out of environment space

O seguinte boletim do Microsoft Knowledge Base Article 230205
(http://support.microsoft.com/support/kb/articles/Q230/2/05.ASP)
pode ajudar.

MOTIVO

Essa mensagem pode ocorrer se voc� n�o tiver mem�ria livre no ambiente
MS-DOS para disponibilizar uma vari�vel de ambiente. 

SOLU��O

Para resolver essa mensagem, aumente o valor do espa�o padr�o para o ambiente 
para programas MS-DOS. Para fazer isso, use um dos seguintes m�todos. 

Modifique o ambiente para todos os programas MS-DOS

Para aumentar o espa�o padr�o de ambiente para todos os programas MS-DOS 
rodando no Windows, edite o Shell de comandos no arquivo Config.sys. 
Para fazer isso, siga os passos abaixo: 

1.  Iniciar, e ent�o em Executar. 

2.  Na caixa aberta,digite sysedit, e ent�o clique em OK. 

3.  Clique na Janela do Config.sys. 

4.  No come�o da linha contendo Shell=, digite REM, e ent�o pressione
    a barrade espa�o. 

    Se a linha contendo Shell= n�o existir, conforme os passos seguintes. 

5.  Pressione a tecla HOME. 

6.  Digite a seguinte linha para criar uma nova linha contendo Shell=, e ent�o
    pressione ENTER: 

      SHELL=C:\COMMAND.COM /E:4096 /P 

7.  No menu Arquivo, clique em Salvar. 

8.  No menu Arquivo, clique Sair. 

9.  Reinicie o computador. 


Modificar o ambiente para um programa espec�fico do MS-DOS

ara aumentar espa�o de ambiente padr�o apenas para maxima.bat, siga
esses passos: 

1.  D� um clique com o bot�o direito do mouse sobre o �cone do programa maxima.bat, e ent�o clique em Propriedades. 

2.  Clique na guia de Mem�ria. 

3.  Na caixa inicial do Ambiente, digite o n�mero de kilobytes (KB) 
    que o programa requer, e ent�o clique em OK. 

NOTA: O m�ximo de me�ria que voc� pode disponibilizar � 4096 KB.

