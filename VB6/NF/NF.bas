Attribute VB_Name = "mod_NF"
Option Explicit

'===============================================================================
'
'�                                M�DULO NF
'�                         IMPRESS�O DE NOTA FISCAL
'�                      _______________________________
'�
'�                          EDI��O = 073
'�                          DATA   = 10.MAI.2022
'�                      _______________________________
'�
'�
'===============================================================================
'
'�_____________________________________________________________________________
'|                                                                             |
'|  H I S T � R I C O    D A S    A L T E R A � � E S                          |
'|_____________________________________________________________________________|
'|          |      |                                                           |
'|  DATA    | RESP | ALTERA��O                                                 |
'|__________|______|___________________________________________________________|
'|19.08.2003| HHO  | In�cio v 1.00                                             |
'|__________|______|___________________________________________________________|
'|10.11.2003| HHO  | V 1.01                                                    |
'|          |      | Implementa��o das seguintes altera��es:                   |
'|          |      |   1) Permitir a digita��o de uma s�rie de pedidos p/ im-  |
'|          |      |      primir em uma �nica nota.                            |
'|          |      |   2) Criar um campo p/ especificar uma loja destinat�ria, |
'|          |      |      sendo que quando este campo est� preenchido, os dados|
'|          |      |      do destinat�rio s�o preenchidos com os dados da loja.|
'|          |      |   3) Preenchimento autom�tico do campo "observa��es II" c/|
'|          |      |      o n�mero da nota gerado automaticamente.             |
'|__________|______|___________________________________________________________|
'|22.11.2005| HHO  | V 1.02                                                    |
'|          |      | Implementa��o das seguintes altera��es:                   |
'|          |      |   1) Carga da rela��o de C.F.O.P. a partir de um arquivo- |
'|          |      |      texto, sendo que a ordem de exibi��o � a mesma ordem |
'|          |      |      dos itens no arquivo.                                |
'|          |      |   2) Bot�o para consultar os dados do destinat�rio, pois  |
'|          |      |      em alguns casos � necess�rio conhecer alguns dados   |
'|          |      |      para poder preencher os demais campos. Exemplo: se � |
'|          |      |      pessoa f�sica ou jur�dica, a UF de destino, etc.     |
'|__________|______|___________________________________________________________|
'|30.11.2005| HHO  | V 1.03                                                    |
'|          |      | Inclus�o da op��o de ICMS de 0%                           |
'|__________|______|___________________________________________________________|
'|27.10.2006| HHO  | V 1.04                                                    |
'|          |      | Altera��o do pre�o de venda p/ pre�o de NF.               |
'|__________|______|___________________________________________________________|
'|04.12.2007| HHO  | V 1.05                                                    |
'|          |      |  1) Preencher o endere�o de entrega no campo "Dados adi-  |
'|          |      |     cionais".                                             |
'|          |      |  2) Preencher a raz�o social da transportadora.           |
'|__________|______|___________________________________________________________|
'|05.12.2007| HHO  | V 1.06                                                    |
'|          |      | Altera��o na regra p/ inserir o endere�o de entrega: o    |
'|          |      | operador � obrigado a clicar no bot�o "Insere Endere�o de |
'|          |      | Entrega" antes de imprimir a NF.                          |
'|__________|______|___________________________________________________________|
'|15.04.2008| HHO  | V 1.07                                                    |
'|          |      | 1) Impress�o do novo endere�o da empresa e riscar o ende- |
'|          |      |    re�o antigo.                                           |
'|          |      | 2) Impress�o de texto fixo relativo � responsabilidade    |
'|          |      |    pelo servi�o de instala��o/manuten��o.                 |
'|__________|______|___________________________________________________________|
'|15.04.2008| HHO  | V 1.08                                                    |
'|          |      | A pedido do Carlos, foi refeita a altera��o implementada  |
'|          |      | na vers�o 1.07 sobre a vers�o 1.05 a fim de suprimir a    |
'|          |      | altera��o introduzida na vers�o 1.06.                     |
'|__________|______|___________________________________________________________|
'|16.04.2008| HHO  | V 1.09                                                    |
'|          |      | Ao "riscar" o endere�o antigo, passa a ser usado o carac- |
'|          |      | ter "X" ao inv�s de pintar uma tarja preta.               |
'|__________|______|___________________________________________________________|
'|14.05.2008| HHO  | V 1.10                                                    |
'|          |      | Ignorar os produtos (t_PEDIDO_ITEM) que tenham 'preco_NF' |
'|          |      | igual a zero, pois s�o brindes e v�o dentro da caixa de   |
'|          |      | outro produto.                                            |
'|__________|______|___________________________________________________________|
'|08.09.2008| HHO  | V 1.11                                                    |
'|          |      | Novo recurso para "chavear" a conex�o do BD entre o site  |
'|          |      | Artven Bonshop e Artven Fabricante.                       |
'|__________|______|___________________________________________________________|
'|23.09.2008| HHO  | V 1.12                                                    |
'|          |      | Remo��o da rotina que risca o endere�o antigo da NF e im- |
'|          |      | prime o endere�o novo devido � chegada do novo formul�rio |
'|          |      | que j� est� c/ endere�o correto.                          |
'|__________|______|___________________________________________________________|
'|12.11.2008| HHO  | V 1.13                                                    |
'|          |      | Imprime uma mensagem no campo 'Dados adicionais' informan-|
'|          |      | do 'Bem de uso e consumo' quando o campo StBemUsoConsumo  |
'|          |      | for igual a 1.                                            |
'|__________|______|___________________________________________________________|
'|05.01.2009| HHO  | V 1.14                                                    |
'|          |      | Impede a impress�o quando o campo 'Entrega Imediata' do   |
'|          |      | pedido for 'N�o'.                                         |
'|__________|______|___________________________________________________________|
'|22.05.2009| HHO  | V 1.15                                                    |
'|          |      | Altera��o no texto de alerta que vai no rodap� do quadro  |
'|          |      | com os itens da NF.                                       |
'|          |      | Como o novo texto consome um espa�o maior, foi retirado o |
'|          |      | espa�amento extra de 1mm que havia entre cada linha nos   |
'|          |      | itens da NF. Al�m disso, foi reduzida a quantidade m�xima |
'|          |      | de linhas impressas de 16 p/ 15 linhas.                   |
'|__________|______|___________________________________________________________|
'|26.05.2009| HHO  | V 1.16                                                    |
'|          |      | Possibilita a impress�o quando o campo 'Entrega Imediata' |
'|          |      | do pedido for 'N�o' mediante confirma��o.                 |
'|__________|______|___________________________________________________________|
'|09.06.2009| HHO  | V 1.17                                                    |
'|          |      | 1) Altera��o no layout do texto impresso no rodap� do qua-|
'|          |      |    dro com os itens da NF para reduzir em uma linha o es- |
'|          |      |    pa�o gasto. Com isso, o m�ximo de linhas impressas     |
'|          |      |    volta a ser de 16 ao inv�s de 15.                      |
'|          |      | 2) Verifica��o se algum dos pedidos possui pagamento por  |
'|          |      |    boleto, pois em caso negativo, a mensagem de alerta    |
'|          |      |    � omitida. Neste caso, o m�ximo de linhas impressas    |
'|          |      |    passa de 16 p/ 18 linhas.                              |
'|__________|______|___________________________________________________________|
'|13.08.2009| HHO  | V 1.17B                                                   |
'|          |      | Adapta��o ao novo modelo de NF.                           |
'|__________|______|___________________________________________________________|
'|20.07.2009| HHO  | V 1.18                                                    |
'|          |      | Ao imprimir a NF de um pedido que ser� pago por boleto,   |
'|          |      | ser�o calculados a quantidade de parcelas, datas e valores|
'|          |      | dos boletos. Esses dados ser�o impressos na NF e ser�o    |
'|          |      | armazenados no BD para serem usados na gera��o dos boletos|
'|          |      | no arquivo de remessa.                                    |
'|__________|______|___________________________________________________________|
'|14.09.2009| HHO  | V 1.19                                                    |
'|          |      | Ao calcular as parcelas dos boletos no caso em que h� mais|
'|          |      | de um pedido, verifica se h� conflito do plano de contas, |
'|          |      | j� que este � definido por loja.                          |
'|__________|______|___________________________________________________________|
'|02.10.2009| HHO  | V 1.19                                                    |
'|          |      | Foi mantido o mesmo n�mero de vers�o.                     |
'|          |      | Altera��o da op��o default do combo "Frete por Conta" de  |
'|          |      | "2 - Destinat�rio" para "1 - Emitente"                    |
'|__________|______|___________________________________________________________|
'|15.10.2009| HHO  | V 1.19                                                    |
'|          |      | Como uma solu��o tempor�ria, o Rog�rio est� usando o campo|
'|          |      | "obs_2" do pedido p/ armazenar uma letra durante uma roti-|
'|          |      | na de an�lise.                                            |
'|          |      | Para n�o causar impactos ao operador durante a impress�o  |
'|          |      | da NF, foi realizada uma altera��o que trata o conte�do do|
'|          |      | campo "obs_2" como se estivesse vazio caso esteja preen-  |
'|          |      | chido com uma letra.                                      |
'|__________|______|___________________________________________________________|
'|11.03.2010| HHO  | V 1.20                                                    |
'|          |      | Como prepara��o para a emiss�o da NFe, todos os cadastros |
'|          |      | que cont�m endere�os foram alterados para separar as in-  |
'|          |      | forma��es "n�mero" e "complemento" do endere�o.           |
'|          |      | Como consequ�ncia disso, este m�dulo foi adaptado p/ mon- |
'|          |      | tar o endere�o considerando a possibilidade do endere�o   |
'|          |      | j� estar no formato novo ou n�o.                          |
'|__________|______|___________________________________________________________|
'|29.03.2010| HHO  | V 1.21                                                    |
'|          |      | Implementa��o da integra��o com o sistema de Nota Fiscal  |
'|          |      | Eletr�nica da empresa Target One.                         |
'|          |      | A integra��o ser� realizada atrav�s do envio de dados da  |
'|          |      | NF via banco de dados em SQL Server hospedado no mesmo    |
'|          |      | servidor do sistema da Artven.                            |
'|          |      | H� duas inst�ncias do sistema de NFe da Target One insta- |
'|          |      | ladas no servidor, cada qual com seu pr�prio banco de     |
'|          |      | dados. Uma inst�ncia � para a OLD03 e outra para a OLD01. |
'|__________|______|___________________________________________________________|
'|19.04.2010| HHO  | V 1.22                                                    |
'|          |      | Implementa��o de novo painel para emitir NFe manualmente, |
'|          |      | ou seja, n�o � informado nenhum n�mero de pedido, os dados|
'|          |      | s�o informados atrav�s da digita��o do CNPJ/CPF do desti- |
'|          |      | nat�rio, c�digo/valor/qtde dos produtos, sele��o do emi-  |
'|          |      | tente da NFe, etc.                                        |
'|__________|______|___________________________________________________________|
'|24.06.2010| HHO  | V 1.23                                                    |
'|          |      | Implementa��o de altera��es para emitir NFe em situa��es  |
'|          |      | que envolvam ST (substitui��o tribut�ria).                |
'|          |      | Al�m disso, para possibilitar a emiss�o de NFe complemen- |
'|          |      | tar futuramente, todos os dados da NFe emitida est�o sendo|
'|          |      | armazenados em novas tabelas.                             |
'|__________|______|___________________________________________________________|
'|28.06.2010| HHO  | V 1.24                                                    |
'|          |      | Ajustes na emiss�o da NFe ainda durante a fase de desen-  |
'|          |      | volvimento e testes em ambiente de homologa��o no sistema |
'|          |      | da SEFAZ (Secretaria da Fazenda).                         |
'|__________|______|___________________________________________________________|
'|29.06.2010| HHO  | V 1.25                                                    |
'|          |      | Ajustes na emiss�o da NFe ainda durante a fase de desen-  |
'|          |      | volvimento e testes em ambiente de homologa��o no sistema |
'|          |      | da SEFAZ (Secretaria da Fazenda).                         |
'|__________|______|___________________________________________________________|
'|30.06.2010| HHO  | V 1.26                                                    |
'|          |      | Vers�o com ajustes finais para emiss�o de NFe em ambiente |
'|          |      | de produ��o da SEFAZ (Secretaria da Fazenda).             |
'|__________|______|___________________________________________________________|
'|05.07.2010| HHO  | V 1.27                                                    |
'|          |      | Como h� muitos clientes cadastrados com o DDD do telefone |
'|          |      | usando 3 d�gitos (0XX), foi implementada uma altera��o p/ |
'|          |      | retirar automaticamente o zero da esquerda durante a emis-|
'|          |      | s�o da NFe.                                               |
'|__________|______|___________________________________________________________|
'|09.07.2010| HHO  | V 1.28                                                    |
'|          |      | Inclus�o de novo bot�o para emitir a NFe informando ma-   |
'|          |      | nualmente o n�mero da nota. O objetivo � evitar que hajam |
'|          |      | lacunas na numera��o decorrentes das seguintes situa��es: |
'|          |      |   1) Ocorre algum erro inesperado no processamento entre  |
'|          |      |      a gera��o do n�mero e o envio dos dados p/ o sistema |
'|          |      |      da Target One. Neste caso, o n�mero pode ser utiliza-|
'|          |      |      do p/ emitir qualquer NFe, j� que na primeira tenta- |
'|          |      |      tiva, os dados n�o chegaram ao sistema da Target One.|
'|          |      |   2) A NFe � enviada para o sistema da Target One, mas    |
'|          |      |      ocorre algum problema na valida��o dos dados. Neste  |
'|          |      |      caso, deve-se reenviar a mesma NFe com as corre��es  |
'|          |      |      necess�rias.                                         |
'|          |      |                                                           |
'|          |      | Lembrando que os n�meros n�o utilizados devem ser cance-  |
'|          |      | lados dentro de um prazo determinado e que o cancelamen-  |
'|          |      | to exige que seja informada uma justificativa.            |
'|__________|______|___________________________________________________________|
'|12.07.2010| HHO  | V 1.29                                                    |
'|          |      | Corre��o da rotina de emiss�o de NFe para gravar o n�mero |
'|          |      | da NF junto c/ os dados gerados p/ a emiss�o de boletos.  |
'|          |      | A rotina continuava usando a vari�vel global que armaze-  |
'|          |      | nava o n� da NF quando a numera��o ainda era controlada   |
'|          |      | atrav�s da grava��o do n� no registry da m�quina. Assim,  |
'|          |      | o valor dessa vari�vel estava sempre com zero.            |
'|__________|______|___________________________________________________________|
'|15.07.2010| HHO  | V 1.30                                                    |
'|          |      | Realizados os seguintes ajustes:                          |
'|          |      |   1) Na digita��o do campo 'Dados Adicionais', agora s�o  |
'|          |      |      aceitos caracteres min�sculos e mai�sculos, ao inv�s |
'|          |      |      de somente mai�sculos.                               |
'|          |      |   2) Na emiss�o de NFe atrav�s de um pedido, o n� do pedi-|
'|          |      |      do e a eventual informa��o 'Bem de Uso e Consumo' fo-|
'|          |      |      ram colocados na parte de baixo do quadro de produtos|
'|          |      |      ao inv�s do quadro 'Dados Adicionais'.               |
'|__________|______|___________________________________________________________|
'|30.07.2010| HHO  | V 1.31                                                    |
'|          |      | Realizados os seguintes ajustes:                          |
'|          |      |   1) Verifica��o se o n� da NFe a ser emitida foi inuti-  |
'|          |      |      lizado antes de prosseguir com a emiss�o.            |
'|          |      |   2) Inclus�o de campos para permitir o preenchimento de  |
'|          |      |      informa��es no campo "infAdProd" para cada item (pro-|
'|          |      |      duto) contendo informa��es adicionais do produto. No |
'|          |      |      caso do painel de emiss�o autom�tico, foi necess�rio |
'|          |      |      adicionar campos na tela para exibir todos os itens  |
'|          |      |      do pedido.                                           |
'|__________|______|___________________________________________________________|
'|06.08.2010| HHO  | V 1.32                                                    |
'|          |      | Implementa��o de tratamento para as seguintes situa��es   |
'|          |      | no caso de haver produtos com CST 60 (ICMS ST):           |
'|          |      |   A) Emiss�o p/ estados n�o conveniados                   |
'|          |      |   B) Devolu��o de mercadorias                             |
'|          |      | Nestes casos, o CST � alterado p/ "00" e o ICMS � calcula-|
'|          |      | do normalmente usando a al�quota selecionada.             |
'|__________|______|___________________________________________________________|
'|12.11.2010| HHO  | V 1.33                                                    |
'|          |      | Implementa��o de uma nova l�gica de funcionamento no modo |
'|          |      | de emiss�o autom�tica:                                    |
'|          |      |   1) No relat�rio de Solicita��o de Coletas, o operador   |
'|          |      |      pode selecionar os pedidos p/ os quais deseja solici-|
'|          |      |      tar a emiss�o da NFe.                                |
'|          |      |   2) Os pedidos selecionados s�o colocados em uma fila e  |
'|          |      |      este programa de emiss�o de NFe trata automaticamente|
'|          |      |      os pedidos da fila.                                  |
'|__________|______|___________________________________________________________|
'|15.12.2010| HHO  | V 1.34                                                    |
'|          |      | Implementa��o de consist�ncias na emiss�o de NFe:         |
'|          |      | 1) D�gito verificador do n�mero de IE                     |
'|          |      | 2) Munic�pio do destinat�rio consta na rela��o de muni-   |
'|          |      |    c�pios do IBGE                                         |
'|          |      | Al�m disso, foi criada uma opera��o p/ realizar o download|
'|          |      | de todos os PDF's de DANFE de uma determinada data.       |
'|__________|______|___________________________________________________________|
'|10.01.2011| HHO  | V 1.35                                                    |
'|          |      | Altera��es na emiss�o de NFe de entrada atrav�s do painel |
'|          |      | de emiss�o autom�tica:                                    |
'|          |      | 1) A emiss�o de nota de entrada n�o deve registrar o n�   |
'|          |      |    da NFe no campo obs_2 do pedido, mas sim no novo campo |
'|          |      |    criado p/ isso na tabela de itens devolvidos.          |
'|          |      | 2) Analisar o pedido p/ verificar se a totalidade dos pro-|
'|          |      |    dutos foi devolvida antes de prosseguir c/ a emiss�o   |
'|          |      |    da NFe de entrada, pois caso a devolu��o tenha sido    |
'|          |      |    parcial, o operador dever� fazer a emiss�o da NFe atra-|
'|          |      |    v�s do painel de emiss�o manual e editar o pedido tam- |
'|          |      |    b�m manualmente p/ anotar o n� da NFe no item que foi  |
'|          |      |    devolvido.                                             |
'|__________|______|___________________________________________________________|
'|24.03.2011| HHO  | V 1.36                                                    |
'|          |      | Revis�o e ajustes para adequar � vers�o 2.0 do layout do  |
'|          |      | xml da NFe.                                               |
'|__________|______|___________________________________________________________|
'|10.06.2011| HHO  | V 1.37                                                    |
'|          |      | Inclus�o da informa��o da cubagem total para ser impressa |
'|          |      | na DANFE. Como n�o h� campo espec�fico p/ tal informa��o, |
'|          |      | ela ser� inclu�da junto com o texto do campo "dados adi-  |
'|          |      | cionais".                                                 |
'|__________|______|___________________________________________________________|
'|11.08.2011| HHO  | V 1.38                                                    |
'|          |      | Inclus�o de um campo edit�vel para a quantidade total de  |
'|          |      | volumes nos pain�is de emiss�o autom�tica e manual.       |
'|          |      | Se o campo estiver preenchido, informar esse valor p/ a   |
'|          |      | emiss�o da NFe.                                           |
'|__________|______|___________________________________________________________|
'|27.09.2011| HHO  | V 1.39                                                    |
'|          |      | Altera��o no painel de emiss�o manual para incluir mais 2 |
'|          |      | linhas p/ os produtos e permitir a emiss�o manual de notas|
'|          |      | que atendam os pedidos que utilizem todas as linhas dis-  |
'|          |      | pon�veis.                                                 |
'|__________|______|___________________________________________________________|
'|30.03.2012| HHO  | V 1.40                                                    |
'|          |      | Implementa��o das seguintes altera��es:                   |
'|          |      |  1) Preenchimento dos campos "cPais" e "xPais" informando |
'|          |      |     sempre como Brasil.                                   |
'|          |      |  2) Inclus�o do campo "outras despesas acess�rias" no pai-|
'|          |      |     nel de emiss�o manual para que seja usado p/ infor-   |
'|          |      |     mar o valor do IPI em notas de devolu��o ao fornece-  |
'|          |      |     dor, j� que o valor total da nota deve representar o  |
'|          |      |     valor total dos produtos mais o valor do IPI.         |
'|          |      |  3) Retirada da altera��o autom�tica do CST "60" para "00"|
'|          |      |     que existia p/ alguns CFOP's devido aos estados n�o   |
'|          |      |     conveniados e tamb�m p/ devolu��o de mercadorias. N�o |
'|          |      |     h� nenhuma situa��o agora que necessite dessa altera- |
'|          |      |     ��o autom�tica, principalmente porque essa altera��o  |
'|          |      |     acarreta em novo recolhimento de ICMS.                |
'|__________|______|___________________________________________________________|
'|30.03.2012| HHO  | V 1.41                                                    |
'|          |      | Para evitar o erro "O campo 'vOutro' n�o esta de acordo." |
'|          |      | que passou a ocorrer na emiss�o manual, foi adicionada a  |
'|          |      | verifica��o que somente informa o campo se o valor for    |
'|          |      | diferente de zero no bloco dos itens.                     |
'|__________|______|___________________________________________________________|
'|04.04.2012| HHO  | V 1.42                                                    |
'|          |      | Inclus�o do c�lculo do PIS/COFINS nas NFe's emitidas atra-|
'|          |      | v�s dos pain�is autom�tico e manual.                      |
'|__________|______|___________________________________________________________|
'|26.04.2012| HHO  | V 1.43                                                    |
'|          |      | Altera��o do painel de emiss�o manual para incluir os cam-|
'|          |      | pos CST e CFOP para cada um dos itens, permitindo a edi-  |
'|          |      | ��o de forma individual.                                  |
'|          |      | Retirada a palavra 'PEDIDO' que precedia o n�mero do pedi-|
'|          |      | do no campo de dados adicionais no painel de emiss�o auto-|
'|          |      | m�tica.                                                   |
'|__________|______|___________________________________________________________|
'|09.08.2012| HHO  | V 1.44                                                    |
'|          |      | Ajustes na rotina de consist�ncia para aceitar telefones  |
'|          |      | com 9 d�gitos devido � inclus�o do 9� d�gito nos telefones|
'|          |      | celulares de S�o Paulo.                                   |
'|__________|______|___________________________________________________________|
'|30.08.2012| HHO  | V 1.45                                                    |
'|          |      | Implementa��o das seguintes altera��es:                   |
'|          |      |  1) No painel de emiss�o autom�tico foram adicionados os  |
'|          |      |     campos: "outras despesas acess�rias", CST e CFOP para |
'|          |      |     permitir um maior grau de liberdade na edi��o da nota |
'|          |      |     a ser emitida.                                        |
'|          |      |  2) No painel de emiss�o manual passa a ser poss�vel in-  |
'|          |      |     formar dados de produtos que n�o est�o cadastrados no |
'|          |      |     BD. Este recurso foi adicionado devido �s NFe's emiti-|
'|          |      |     das em decorr�ncia de produtos em assist�ncia t�cnica |
'|          |      |     (lembrando que na assist�ncia t�cnica podem estar sen-|
'|          |      |     do usados c�digos diferentes ou produtos diferentes do|
'|          |      |     cadastro de produtos do sistema de produ��o).         |
'|          |      |     Alguns campos foram adicionados p/ que as informa��es |
'|          |      |     obtidas do BD pudessem ser fornecidas manualmente:    |
'|          |      |     NCM, peso bruto total e peso l�quido total. Al�m disso|
'|          |      |     � necess�rio preencher a descri��o do produto, sendo  |
'|          |      |     que o campo "descri��o" fica edit�vel nessa situa��o. |
'|          |      |     Para emitir uma nota informando produtos n�o cadastra-|
'|          |      |     dos, � necess�rio liberar o modo de edi��o manual cli-|
'|          |      |     cando no bot�o "Liberar Edi��o". Neste momento, ser�  |
'|          |      |     exibido um alerta e solicitada a confirma��o atrav�s  |
'|          |      |     da digita��o da senha. No banco de dados ficar� regis-|
'|          |      |     trado que a nota foi emitida com o modo de edi��o ma- |
'|          |      |     nual liberado.                                        |
'|          |      |     Quando um c�digo de fabricante e/ou produto n�o � en- |
'|          |      |     contrado no BD, o respectivo campo fica em vermelho p/|
'|          |      |     indicar essa situa��o ao usu�rio.                     |
'|__________|______|___________________________________________________________|
'|31.08.2012| HHO  | V 1.46                                                    |
'|          |      | Corre��o de bug ao emitir NF no painel manual: quando a   |
'|          |      | edi��o manual n�o est� ativada e um produto cadastrado no |
'|          |      | BD � informado, a descri��o era exibida automaticamente na|
'|          |      | tela, mas n�o estava sendo executada o trecho da rotina   |
'|          |      | que escolhia entre a descri��o cadastrada no BD e a des-  |
'|          |      | cri��o preenchida na tela. Por isso, era exibida a men-   |
'|          |      | sagem de erro "O produto 999999 N�O possui descri��o!!".  |
'|__________|______|___________________________________________________________|
'|03.09.2012| HHO  | V 1.47                                                    |
'|          |      | Ajustes no painel de emiss�o manual devido a problemas de |
'|          |      | arredondamento nos c�lculos do peso bruto total e peso l�-|
'|          |      | quido total. Altera��o do tipo de dados 'Double' para     |
'|          |      | 'Single' para compatibilizar com o tipo de dados 'Real' do|
'|          |      | dos campos 'peso' e 'cubagem' do BD.                      |
'|__________|______|___________________________________________________________|
'|30.01.2013| HHO  | V 1.48                                                    |
'|          |      | Altera��o na emiss�o autom�tica p/ obter os dados de NCM  |
'|          |      | e CST a partir das informa��es cadastradas na opera��o de |
'|          |      | entrada de mercadorias no estoque (tabela t_ESTOQUE_ITEM).|
'|          |      | Um mesmo produto pode ter o NCM alterado pelo seu fabri-  |
'|          |      | cante, mas, por outro lado, deve haver um "batimento"     |
'|          |      | entre a quantidade que foi comprada e vendida. Por isso   |
'|          |      | surgiu a necessidade de se memorizar essas informa��es    |
'|          |      | para cada lote de produtos durante a entrada de mercado-  |
'|          |      | rias no estoque. Se por acaso um pedido possuir um produ- |
'|          |      | to que misture c�digos diferentes de NCM ou CST, o pro-   |
'|          |      | grama deve desmembrar em linhas diferentes.               |
'|          |      | Inclus�o de campo para al�quota de ICMS para cada produto,|
'|          |      | pois entrou em vigor uma nova lei em que nas vendas inter-|
'|          |      | estaduais de mercadorias importadas a al�quota do ICMS �  |
'|          |      | de 4%. Com isso, se uma NFe mistura produtos importados c/|
'|          |      | produtos nacionais, passa a haver al�quotas diferentes de |
'|          |      | ICMS em uma mesma NFe. O campo � um combo-box do tipo edi-|
'|          |      | t�vel.                                                    |
'|__________|______|___________________________________________________________|
'|22.05.2013| HHO  | V 1.49                                                    |
'|          |      | Inclus�o de campos 'combo-box' p/ indicar os casos em que |
'|          |      | as al�quotas de PIS e COFINS devem ser zero e qual o c�di-|
'|          |      | go de CST a ser usado.                                    |
'|          |      | Supress�o do alerta de poss�vel incoer�ncia na al�quota   |
'|          |      | de ICMS de 4% em vendas interestaduais de mercadorias im- |
'|          |      | portadas quando o destinat�rio for PF ou se for PJ isenta |
'|          |      | de I.E.                                                   |
'|__________|______|___________________________________________________________|
'|05.06.2013| HHO  | V 1.50                                                    |
'|          |      | Para atender � lei 12.741/2012, foram feitas as altera��es|
'|          |      | p/ calcular o valor total estimado dos tributos usando os |
'|          |      | percentuais fornecidos pelo IBPT.                         |
'|          |      | O IBPT fornece os dados atrav�s de um arquivo CSV. Foi de-|
'|          |      | senvolvido um novo m�dulo chamado ADM2 com a rotina para  |
'|          |      | carregar os dados do arquivo p/ o BD.                     |
'|__________|______|___________________________________________________________|
'|11.09.2013| HHO  | V 1.51                                                    |
'|          |      | Inclus�o do campo 'xPed' (n� do pedido de compra) na se��o|
'|          |      | dos itens da NFe.                                         |
'|__________|______|___________________________________________________________|
'|19.09.2013| HHO  | V 1.52                                                    |
'|          |      | Inclus�o do email da transportadora junto com o email do  |
'|          |      | cliente (separados por ponto e v�rgula) devido � necessi- |
'|          |      | dade de enviar o XML da NF p/ as transportadoras.         |
'|__________|______|___________________________________________________________|
'|24.10.2013| HHO  | V 1.53                                                    |
'|          |      | Corre��o de um problema que ocorre quando h� mais de uma  |
'|          |      | linha com o mesmo produto. Isso pode acontecer quando o   |
'|          |      | pedido consome de lotes do estoque que tenham sido cadas- |
'|          |      | trados c/ NCM e/ou CST distintos. O problema estava na    |
'|          |      | forma como os campos edit�veis da tela eram lidos, pois   |
'|          |      | todas as ocorr�ncias do produto estavam obtendo sempre a  |
'|          |      | primeira linha da tela em que o produto aparecia. Assim,  |
'|          |      | se as demais ocorr�ncias tinham um c�digo de NCM e/ou CST |
'|          |      | diferentes, o programa acabava assumindo que o usu�rio    |
'|          |      | editou esses campos e todas as ocorr�ncias ficavam c/ os  |
'|          |      | mesmos valores. Mesmo quando nenhuma edi��o havia sido    |
'|          |      | feita, o programa acabava entendendo que as demais ocor-  |
'|          |      | r�ncias do produto haviam sido editadas.                  |
'|__________|______|___________________________________________________________|
'|09.12.2013| HHO  | V 1.54                                                    |
'|          |      | Altera��o do painel de emiss�o manual para permitir a edi-|
'|          |      | ��o do endere�o. Isso tornar� mais f�cil a emiss�o de no- |
'|          |      | tas de simples remessa at� que uma solu��o mais completa  |
'|          |      | e automatizada seja desenvolvida.                         |
'|__________|______|___________________________________________________________|
'|31.01.2014| HHO  | V 1.55                                                    |
'|          |      | Implementa��o da fun��o de download dos PDF's de DANFE por|
'|          |      | per�odo de datas. Tamb�m foi implementado tratamento p/ a |
'|          |      | situa��o em que n�o h� dados do PDF no BD, pois nesse caso|
'|          |      | estavam sendo criados arquivos vazios de PDF. Se houvessem|
'|          |      | arquivos baixados anteriormente c/ conte�do v�lido, havia |
'|          |      | o risco de que esses arquivos fossem substitu�dos por ar- |
'|          |      | quivos inv�lidos.                                         |
'|          |      | No painel de emiss�o autom�tico, quando um pedido possui  |
'|          |      | endere�o de entrega que seja na mesma UF do endere�o do   |
'|          |      | cliente, o endere�o de entrega est� sendo preenchido auto-|
'|          |      | maticamente no campo "Dados Adicionais".                  |
'|__________|______|___________________________________________________________|
'|03.02.2014| HHO  | V 1.56                                                    |
'|          |      | Altera��o do tratamento do endere�o de entrega: foi rever-|
'|          |      | tida a �ltima altera��o em que se preenchia automaticamen-|
'|          |      | te o endere�o de entrega no campo "Dados Adicionais" quan-|
'|          |      | do dentro da mesma UF. Percebeu-se que o endere�o de en-  |
'|          |      | trega j� era adicionado automaticamente no campo "Dados   |
'|          |      | Adicionais" na integra��o com o sistema da Target One, o  |
'|          |      | que causava duplicidade da informa��o.                    |
'|          |      | Entretanto, a verifica��o quanto ser dentro da mesma UF   |
'|          |      | n�o estava sendo realizada, o que foi implementado agora. |
'|__________|______|___________________________________________________________|
'|22.04.2014| HHO  | V 1.57                                                    |
'|          |      | Inclus�o de mensagem alertando que o fabricante n�o cobre |
'|          |      | avarias em pe�as pl�sticas e que � necess�rio verificar   |
'|          |      | o produto no ato da entrega.                              |
'|__________|______|___________________________________________________________|
'|15.01.2015| HHO  | V 1.58                                                    |
'|          |      | Altera��o da rotina de decriptografia da senha no login   |
'|          |      | devido � nova senha usada no servidor dedicado n�o ser    |
'|          |      | suportada pelo algoritmo antigo de criptografia.          |
'|__________|______|___________________________________________________________|
'|24.03.2015| LHGX | V 1.59                                                    |
'|          |      | Implementa��o das altera��es necess�rias para o novo      |
'|          |      | layout 3.10 da NFe.                                       |
'|__________|______|___________________________________________________________|
'|01.04.2015| LHGX | V 1.60                                                    |
'|          |      | Vers�o homologada ap�s ajustes para o novo layout da NFe  |
'|          |      | v3.10                                                     |
'|__________|______|___________________________________________________________|
'|01.04.2015| LHGX | V 1.61                                                    |
'|          |      | Corre��o do SQL usado no preenchimento dos itens de um    |
'|          |      | pedido no painel de emiss�o autom�tico.                   |
'|__________|______|___________________________________________________________|
'|14.04.2015| HHO  | V 1.62                                                    |
'|          |      | Corre��o do SQL usado no preenchimento dos itens de um    |
'|          |      | pedido no painel de emiss�o autom�tico.                   |
'|__________|______|___________________________________________________________|
'|05.06.2015| LHGX | V 1.63                                                    |
'|          |      | NFe 3.10 - Verifica��o se cliente � contribuinte ICMS     |
'|          |      | NFe 3.10 - Se � produtor rural, verifica se est� c/       |
'|          |      |            al�quota de ICMS espec�fica (para mercadoria   |
'|          |      |            importada em venda interestadual)              |
'|          |      | NFe 3.10 - Inclus�o do e-mail do destinat�rio na tag      |
'|          |      |            correspondente, se campo email_xml for         |
'|          |      |            preenchido                                     |
'|__________|______|___________________________________________________________|
'|09.06.2015| LHGX | V 1.64                                                    |
'|          |      | Implementa��o da possibilidade de informar mais de uma    |
'|          |      | chave de acesso para NFes referenciadas                   |
'|__________|______|___________________________________________________________|
'|04.10.2015| LHGX |V 1.65                                                     |
'|          |      | Implementa��o do tratamento de m�ltiplos Centros de       |
'|          |      | Distribui��o                                              |
'|          |      | Defini��o de preced�ncia de telefones do cliente a serem  |
'|          |      | preenchidos na NF (Celular, Residencial e Comercial)      |
'|__________|______|___________________________________________________________|
'|04.10.2015| LHGX |V 1.66                                                     |
'|          |      | Aumento da Descri��o do Produto para 80 caracteres no     |
'|          |      | Painel de Emiss�o Manual                                  |
'|__________|______|___________________________________________________________|
'|09.10.2015| LHGX |V 1.67                                                     |
'|          |      | Inclus�o da coluna Unidade no Painel de Emiss�o Manual,   |
'|          |      | para informar a unidade comercial/tribut�ria desejada     |
'|__________|______|___________________________________________________________|
'|23.12.2015| LHGX |V 1.68                                                     |
'|          |      | Inclus�o das al�quotas referentes � UF de ES              |
'|__________|______|___________________________________________________________|
'|29.02.2016| LHGX |V 1.69                                                     |
'|          |      | Remo��o dos 07 dias adicionais na primeira data de        |
'|          |      | vencimento, no caso de pagamentos parcelados              |
'|__________|______|___________________________________________________________|
'|16.03.2016| LHGX |V 1.70                                                     |
'|          |      | Inclus�o do campo 'nItemPed' (item do pedido de compra)   |
'|          |      | na se��o dos itens da Nfe.                                |
'|__________|______|___________________________________________________________|
'|05.05.2016| LHGX |V 1.71                                                     |
'|          |      | Altera��o para o layout 3.11 da NFe, com inclus�o do      |
'|          |      | campo CEST                                                |
'|          |      | Altera��es para funcionamento em diversos ambientes       |
'|          |      | (entrada DIS)                                             |
'|__________|______|___________________________________________________________|
'|27.06.2016| LHGX |V 1.72                                                     |
'|          |      | Conforme orienta��o da contabilidade, os contribuintes    |
'|          |      | isentos passam a ser enviados como n�o contribuintes      |
'|          |      | � SEFAZ                                                   |
'|          |      | Inclus�o de painel para editar parcelas de boletos        |
'|__________|______|___________________________________________________________|
'|29.06.2016| LHGX |V 1.73                                                     |
'|          |      | Altera��o de al�quotas internas do ICMS                   |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|13.07.2016| LHGX |V 1.74                                                     |
'|          |      | Inclus�o de campo com totaliza��o do ICMS                 |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|14.07.2016| LHGX |V 1.75                                                     |
'|          |      | Ajuste para n�o enviar informa��es de partilha se o CFOP  |
'|          |      | for de remessa                                            |
'|__________|______|___________________________________________________________|
'|18.07.2016| LHGX |V 1.76                                                     |
'|          |      | Previs�o dos novos campos na t_PEDIDO:                    |
'|          |      | 'NFe_texto_constar' (para acr�scimo em Dados Adicionais)  |
'|          |      | e 'NFe_xPed' (para preenchimento dos campos xPed)         |
'|          |      | Corre��o no arredondamento do c�lculo do ICMS interesta-  |
'|          |      | dual                                                      |
'|__________|______|___________________________________________________________|
'|27.07.2016| LHGX |V 1.77                                                     |
'|          |      | Corre��o de bug (teste de condi��o de parcelamento no     |
'|          |      | modo Manual quando h� fila de impress�o)                  |
'|__________|______|___________________________________________________________|
'|20.09.2016| LHGX |V 1.78                                                     |
'|          |      | Inclus�o de arquivo para cadastramento de CFOP's para os  |
'|          |      | quais n�o � necess�rio o envio de informa��es de partilha |
'|          |      | de ICMS.                                                  |
'|__________|______|___________________________________________________________|
'|11.10.2016| LHGX |V 1.79                                                     |
'|          |      | Corre��o de arredondamento para o campo de valor estimado |
'|          |      | do total de tributos.                                     |
'|__________|______|___________________________________________________________|
'|25.10.2016| LHGX |V 1.80                                                     |
'|          |      | Sinaliza��o de que o endere�o foi editado no painel de    |
'|          |      | emiss�o manual.                                           |
'|__________|______|___________________________________________________________|
'|16.11.2016| LHGX |V 1.81                                                     |
'|          |      | Corre��o da rotina para altera��o de datas nos parcela-   |
'|          |      | mentos de boletos.                                        |
'|__________|______|___________________________________________________________|
'|10.01.2017| LHGX |V 1.82                                                     |
'|          |      | Inclus�o da possibilidade de altern�ncia de CD's na fila  |
'|          |      | de pedidos e no painel de emiss�o manual.                 |
'|          |      | Possibilidade de escolher se o percentual de partilha do  |
'|          |      | ICMS ser� o do ano atual ou do ano anterior, nos casos    |
'|          |      | de nota de entrada com chave de acesso de nota referen-   |
'|          |      | ciada (painel de emiss�o manual).                         |
'|__________|______|___________________________________________________________|
'|19.01.2017| LHGX |V 1.83                                                     |
'|          |      | Possibilidade de escolher se o percentual de partilha do  |
'|          |      | ICMS ser� o do ano atual ou do ano anterior, nos casos    |
'|          |      | de nota de entrada com chave de acesso de nota referen-   |
'|          |      | ciada (painel de emiss�o autom�tica).                     |
'|__________|______|___________________________________________________________|
'|31.03.2017| LHGX |V 1.84                                                     |
'|          |      | Remo��o do antigo controle por CD's.                      |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|03.05.2017| LHGX |V 1.85                                                     |
'|          |      | Filtro por UF no download de DANFE por per�odo            |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|09.06.2017| LHGX |V 1.86                                                     |
'|          |      | Cria��o da Tabela de Al�quotas Internas de ICMS           |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|28.06.2017| LHGX |V 1.87                                                     |
'|          |      | Emiss�o de Nota Fiscal Complementar de ICMS atrav�s do    |
'|          |      | painel de emiss�o manual                                  |
'|__________|______|___________________________________________________________|
'|XX.XX.XXXX| LHGX |V 1.88                                                     |
'|          |      | Adi��o do campo IE nos forms para indica��o da modalidade |
'|          |      | de contribui��o de ICMS (C = contribuinte, NC = n�o con-  |
'|          |      | tribuinte, I = isento, PR = produtor rural, em branco =   |
'|          |      | pessoa f�sica)                                            |
'|__________|______|___________________________________________________________|
'|31.08.2017| LHGX |V 1.89                                                     |
'|          |      | Grava��o do campo cEANTrib conforme exig�ncia futura da   |
'|          |      | SEFAZ                                                     |
'|__________|______|___________________________________________________________|
'|17.10.2017| LHGX |V 1.90                                                     |
'|          |      | Tratamento na fila autom�tica para ignorar pedidos can-   |
'|          |      | celados                                                   |
'|__________|______|___________________________________________________________|
'|13.11.2017| LHGX |V 1.91                                                     |
'|          |      | Tela para Emiss�o de Notas Triangulares                   |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|28.12.2017| LHGX |V 1.92                                                     |
'|          |      | Ajuste al�quota 12% para ES                               |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|06.02.2018| LHGX |V 1.93                                                     |
'|          |      | - Ajuste no conte�do do campo idDest na opera��o          |
'|          |      |   triangular                                              |
'|          |      | - Consist�ncia para al�quota de ICMS 0% em opera��es      |
'|          |      |   interestaduais                                          |
'|__________|______|___________________________________________________________|
'|28.02.2018| LHGX |V 1.94                                                     |
'|          |      | - Inclus�o do campo Pedido na Tela de Emiss�o Manual      |
'|          |      | - Ajustes nas rotinas que consultam dados do CEP para     |
'|          |      |   que acessem o novo banco de dados de CEP, sendo que a   |
'|          |      |   estrutura das tabelas tamb�m foi alterada.              |
'|__________|______|___________________________________________________________|
'|23.04.2018| lhgx |V 1.95                                                     |
'|          |      | - Corre��o NFe Triangular (ap�strofe)                     |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|14.06.2018| LHGX |V 1.96                                                     |
'|          |      | - Altera��o hard code do telefone (11)3934-4420           |
'|          |      |   para (11)4858-2431                                      |
'|__________|______|___________________________________________________________|
'|XX.XX.XXXX| LHGX |V 1.97                                                     |
'|          |      | - Adapta��o para o layout 4.0 da NFe                      |
'|          |      | - Altera��o na tela de emiss�o triangular para deixar     |
'|          |      |   os controles "Local de Destino(Remessa)" e "Natureza    |
'|          |      |   da Opera��o(Remessa)" pr�-selecionados para opera��es   |
'|          |      |   interestaduais                                          |
'|          |      | - Corre��o do bug de atualiza��o do campo emissao_status  |
'|          |      |   na emiss�o de notas triangulares previamente canceladas |
'|          |      |                                                           |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|02.08.2018| LHGX |V 2.00 (substituiu V 1.99 e 1.98 que tinham bug)           |
'|          |      | - Enviar VAlor de Pagamento 0 para tPag=90                |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|09.08.2018| LHGX |V 2.01                                                     |
'|          |      | - Corre��o de bug para emiss�o de nota de remessa para    |
'|          |      |   PJ no painel de emiss�o triangular                      |
'|__________|______|___________________________________________________________|
'|27.08.2018| LHGX |V 2.02                                                     |
'|          |      | - Op��o de Incluir Dados de Parcelas no Campo Informa��es |
'|          |      |   Adicionais                                              |
'|          |      | - Pesquisa de IE na Nota de Remessa de Opera��es Trian-   |
'|          |      |   gulares                                                 |
'|          |      |                                                           |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|15.09.2018| LHGX |V 2.03                                                     |
'|          |      | - Inclus�o novamente da tag nDup no XML, ap�s ajustes     |
'|          |      |   e testes com a Target                                   |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|01.11.2018| LHGX |V 2.04                                                     |
'|          |      | - Ajuste de tag's de totaliza��o do FCP                   |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|26.11.2018| LHGX |V 2.05                                                     |
'|          |      | - Novo ajuste de tag's de totaliza��o do FCP              |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|02.01.2019| LHGX |V 2.07                                                     |
'|          |      | - Corre��o no arrendondamento do ICMS interestadual       |
'|          |      |   (campo vICMSUFRemet)                                    |
'|__________|______|___________________________________________________________|
'|25.04.2019| LHGX |V 2.08                                                     |
'|          |      | -  Zerar PIS/COFINS nas notas com os CFOP's abaixo:       |
'|          |      |    5.915 - 6.152 -5.949 - 6.117 - 6.923 - 6.910           |
'|__________|______|___________________________________________________________|
'|02.07.2019| LHGX |V 2.09                                                     |
'|          |      | - Inclus�o da forma de pagamento "Cart�o com Maquineta"   |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|04.09.2019| LHGX |V 2.10                                                     |
'|          |      | - Cria��o da tabela t_NFe_EMITENTE_NUMERACAO              |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|18.11.2019| LHGX |V 2.11                                                     |
'|          |      | - Corre��o de bug (Emitir NFe com N�mero Manual)          |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|24.01.2020| LHGX |V 2.12                                                     |
'|          |      | - Cria��o do arquivo UFS_INSCRICAO_VIRTUAL.CFG, para que  |
'|          |      |   o texto sobre DIFAL/Partilha n�o seja impresso na NF    |
'|__________|______|___________________________________________________________|
'|10.02.2020| LHGX |V 2.13                                                     |
'|          |      | - Ajustes sobre o UFS_INSCRICAO_VIRTUAL.CFG, para que     |
'|          |      |   a defini��o das UF's seja feita por emitente            |
'|__________|______|___________________________________________________________|
'|24.04.2020| LHGX |V 2.14                                                     |
'|          |      | - Implementa��o da memoriza��o do endere�o do cliente     |
'|          |      |   na t_PEDIDO                                             |
'|__________|______|___________________________________________________________|
'|14.11.2020| LHGX |V 2.15                                                     |
'|          |      | - Ajuste para trazer as informa��es memorizadas no        |
'|          |      |   pedido no quadro Informa��es do Pedido                  |
'|__________|______|___________________________________________________________|
'|16.02.2021| LHGX |V 2.16                                                     |
'|          |      | - Flag para definir se ser� usado o endere�o de cobran�a  |
'|          |      |   ou de entrega nas notas de PF com memoriza��o           |
'|__________|______|___________________________________________________________|
'|04.04.2021| LHGX |V 2.17                                                     |
'|          |      | - Tratamento das informa��es de intermediador para vendas |
'|          |      |   de marketplace                                          |
'|          |      | - Novos meios de pagamento e restri��o ao 99 - Outros     |
'|__________|______|___________________________________________________________|
'|01.06.2021| LHGX |V 2.18                                                     |
'|          |      | - Corre��o da UF do endere�o de venda no painel de        |
'|          |      |   Informa��es do Pedido na tela principal                 |
'|__________|______|___________________________________________________________|
'|06.06.2021| LHGX |V 2.19                                                     |
'|          |      | - Venda Futura / Pagamento Antecipado                     |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|15.07.2021| LHGX |V 2.20                                                     |
'|          |      | - Venda Futura (ajuste painel triangular)                 |
'|          |      | - Obrigatoriedade de preenchimento do campo xBairro       |
'|          |      |   para n�o haver rejei��o na SEFAZ                        |
'|__________|______|___________________________________________________________|
'|31.08.2021| LHGX |V 2.21                                                     |
'|          |      | - Cria��o de par�metros e altera��es para excluir o ICMS  |
'|          |      |   e o DIFAL das bases de c�lculo de PIS e COFINS - de     |
'|          |      |   acordo com decis�o do STF                               |
'|          |      | - Cria��o de par�metros de conting�ncia para informar     |
'|          |      |   emergencialmente o meio de pagamento "99-Outros" no     |
'|          |      |   campo tPag acompanhado da descri��o no campo xPag       |
'|__________|______|___________________________________________________________|
'|02.09.2021| LHGX |V 2.22                                                     |
'|          |      | - Painel de Emiss�o Manual: corre��o para sempre informar |
'|          |      |   vPag=0 quando tPag=90 (sem pagamento)                   |
'|__________|______|___________________________________________________________|
'|18.01.2022| LHGX |V 2.23                                                     |
'|          |      | - Ajuste grava��o campo nItemPed                          |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|03.02.2022| LHGX |V 2.24                                                     |
'|          |      | - Funcionalidade para que o DIFAL n�o seja calculado      |
'|          |      |   em caso de liminar a favor                              |
'|__________|______|___________________________________________________________|
'|06.02.2022| LHGX |V 2.25                                                     |
'|          |      | - Ajuste da vers�o anterior para que as NF's internas     |
'|          |      |   n�o exibam a mensagem sobre a n�o-cobran�a  do DIFAL    |
'|__________|______|___________________________________________________________|
'|25.04.2022| LHGX |V 2.26                                                     |
'|          |      | - Informa��es do intermediador de pagamento nas opera-    |
'|          |      |   ��es envolvendo marketplace                             |
'|          |      | - Registro no log da tela de origem da emiss�o da NFe     |
'|          |      |   (autom�tica, manual, triangular)                        |
'|          |      | - Zerar PIS/COFINS quando natureza da opera��o for 6949   |
'|          |      | - Mudan�a da nomenclatura da pasta local de grava��o      |
'|          |      |   de DANFEs                                               |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|10.05.2022| LHGX |V 2.27                                                     |
'|          |      | - Tratamento do campo t_TRANSPORTADORA.email2             |
'|          |      | - Limitar comprimento de e-mails na tag 'operacional'     |
'|__________|______|___________________________________________________________|
'|XX.XX.XXXX| LHGX |V 2.28                                                     |
'|          |      | - Corre��o de bug em mensagem sobre entrega futura        |
'|          |      |   n�o quitada                                             |
'|__________|______|___________________________________________________________|
'|30.11.2022| LHGX |V 2.29                                                     |
'|          |      | - N�o exibir a mensagem sobre os valores aproximados      |
'|          |      |   de tributos (IBPT) para opera��es de transfer�ncia      |
'|          |      |   de estoque entre filiais (CFOP 5152)                    |
'|__________|______|___________________________________________________________|
'|10.03.2023| LHGX |V 2.30                                                     |
'|          |      | - Inclus�o tag infRespTec                                 |
'|          |      | - Adi��o do par�metro NF_Informa_Resp_Tec para ativar ou  |
'|          |      |   desativar o envio de informa��es da tag infRespTec      |
'|__________|______|___________________________________________________________|
'|XX.XX.XXXX| XXXX |V X.XX                                                     |
'|          |      |                                                           |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|XX.XX.XXXX| XXXX |V X.XX                                                     |
'|          |      |                                                           |
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'


Global Const m_id_versao = "2.30"
Global Const m_id = "Nota Fiscal  v" & m_id_versao & "  10/03/2023"

' N� VERS�O ATUAL DO LAYOUT DOS DADOS DA NFe
Global Const ID_VERSAO_LAYOUT_NFe = "4.00"


' TEXTOS DE MENSAGENS
Global Const TEXTO_LEI_CST_ICMS_60 = "IMPOSTO RECOLHIDO POR SUBSTITUI��O TRIBUT�RIA CONFORME ART.313-Z111 DO RICMS/00"

' AL�QUOTAS
Global Const PERC_PIS_ALIQUOTA_NORMAL = 1.65  ' 1,65%
Global Const PERC_COFINS_ALIQUOTA_NORMAL = 7.6  ' 7,6%
Global Const PERC_ICMS_ALIQUOTA_VENDA_INTERESTADUAL_MERCADORIA_IMPORTADA = 4  ' 4,0%

' A PARTILHA DO ICMS INTERESTADUAL EST� ATIVA?
Global Const PARTILHA_ICMS_ATIVA = True

' FINALIDADE DE EMISS�O
Global Const NFE_FINALIDADE_NFE_NORMAL = "1" '1-Normal  2-Complementar  3-Ajuste 4 - Devolu��o
Global Const NFE_FINALIDADE_NFE_COMPLEMENTAR = "2"
Global Const NFE_FINALIDADE_NFE_AJUSTE = "3"
Global Const NFE_FINALIDADE_NFE_DEVOLUCAO = "4"

' C�DIGOS P/ SOLICITA��O DE EMISS�O DE NFe
Global Const COD_NFE_EMISSAO_SOLICITADA__PENDENTE = "0"
Global Const COD_NFE_EMISSAO_SOLICITADA__ATENDIDA = "1"
Global Const COD_NFE_EMISSAO_SOLICITADA__CANCELADA = "2"

' TIMEOUT P/ SOLICITA��O DE EMISS�O DE NFe REQUISITADA DA FILA E QUE N�O FOI PROCESSADA AT� O FINAL, OU SEJA, N�O TEVE A NFE EMITIDA
Global Const MAX_TIMEOUT_REGISTRO_REQUISITADO_FILA_EM_SEG = 1 * 60

' SEPARA SUFIXO DO PEDIDO FILHOTE
Global Const COD_SEPARADOR_FILHOTE = "-"

  ' C�DIGOS QUE IDENTIFICAM SE � PESSOA F�SICA OU JUR�DICA
Global Const ID_PF = "PF"
Global Const ID_PJ = "PJ"

Global Const TAM_MIN_NUM_NF = 6
Global Const TAM_MIN_NUM_PEDIDO = 6    '�SOMENTE PARTE NUM�RICA DO N�MERO DO PEDIDO
Global Const TAM_MIN_ID_PEDIDO = 7     '�PARTE NUM�RICA DO N�MERO DO PEDIDO + LETRA REFERENTE AO ANO
Global Const TAM_MIN_FABRICANTE = 3
Global Const TAM_MAX_FABRICANTE = 4
Global Const TAM_MIN_PRODUTO = 6
Global Const TAM_MAX_PRODUTO = 8
Global Const TAM_MIN_LOJA = 2
Global Const TAM_MAX_LOJA = 3


'   STATUS DE ENTREGA DO PEDIDO
Global Const ST_ENTREGA_ESPERAR = "ESP"         ' NENHUMA MERCADORIA SOLICITADA EST� DISPON�VEL
Global Const ST_ENTREGA_SPLIT_POSSIVEL = "SPL"  ' PARTE DA MERCADORIA EST� DISPON�VEL PARA ENTREGA
Global Const ST_ENTREGA_SEPARAR = "SEP"         ' TODA A MERCADORIA EST� DISPON�VEL E J� PODE SER SEPARADA PARA ENTREGA
Global Const ST_ENTREGA_A_ENTREGAR = "AET"      ' A TRANSPORTADORA J� SEPAROU A MERCADORIA PARA ENTREGA
Global Const ST_ENTREGA_ENTREGUE = "ETG"        ' MERCADORIA FOI ENTREGUE
Global Const ST_ENTREGA_CANCELADO = "CAN"       ' VENDA FOI CANCELADA

'   TIPOS DE ESTOQUE
Global Const ID_ESTOQUE_VENDA = "VDA"
Global Const ID_ESTOQUE_VENDIDO = "VDO"
Global Const ID_ESTOQUE_SEM_PRESENCA = "SPE"
Global Const ID_ESTOQUE_KIT = "KIT"
Global Const ID_ESTOQUE_SHOW_ROOM = "SHR"
Global Const ID_ESTOQUE_DANIFICADOS = "DAN"
Global Const ID_ESTOQUE_DEVOLUCAO = "DEV"
Global Const ID_ESTOQUE_ROUBO = "ROU"
Global Const ID_ESTOQUE_ENTREGUE = "ETG"

'   FORMA DE PAGAMENTO
Global Const COD_FORMA_PAGTO_A_VISTA = "1"
Global Const COD_FORMA_PAGTO_PARCELADO_CARTAO = "2"
Global Const COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA = "3"
Global Const COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA = "4"
Global Const COD_FORMA_PAGTO_PARCELA_UNICA = "5"
Global Const COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA = "6"
    
Global Const ID_FORMA_PAGTO_DINHEIRO = "1"
Global Const ID_FORMA_PAGTO_DEPOSITO = "2"
Global Const ID_FORMA_PAGTO_CHEQUE = "3"
Global Const ID_FORMA_PAGTO_BOLETO = "4"
Global Const ID_FORMA_PAGTO_CARTAO = "5"
Global Const ID_FORMA_PAGTO_BOLETO_AV = "6"
Global Const ID_FORMA_PAGTO_CARTAO_MAQUINETA = "7"

'   C�DIGOS DE STATUS
Global Const NF_PARCELA_PAGTO__STATUS_INICIAL = "0"
Global Const NF_PARCELA_PAGTO__STATUS_CANCELADO = "1"
Global Const NF_PARCELA_PAGTO__STATUS_TRATADO = "2"

'   NSU
Global Const NSU_T_FIN_NF_PARCELA_PAGTO = "t_FIN_NF_PARCELA_PAGTO"
Global Const NSU_T_FIN_NF_PARCELA_PAGTO_ITEM = "t_FIN_NF_PARCELA_PAGTO_ITEM"
Global Const NSU_T_NFe_EMISSAO = "t_NFe_EMISSAO"
Global Const NSU_T_NFe_IMAGEM = "t_NFe_IMAGEM"
Global Const NSU_T_NFe_IMAGEM_ITEM = "t_NFe_IMAGEM_ITEM"
Global Const NSU_T_NFe_IMAGEM_TAG_DUP = "t_NFe_IMAGEM_TAG_DUP"
Global Const NSU_T_NFe_IMAGEM_NFe_REFERENCIADA = "t_NFe_IMAGEM_NFe_REFERENCIADA"
Global Const NSU_T_NFe_IMAGEM_PAG = "T_NFe_IMAGEM_PAG"

'   CONTROLE DE ACESSO
Global Const OP_CEN_MODULO_NF_ACESSO_AO_MODULO = 21900

' C�DIGOS PARA REGISTRO DE OPERA��ES NO LOG
Global Const OP_LOG_NF_IMPRESSAO = "NF IMPRESSAO"
Global Const OP_LOG_NFE_EMISSAO = "NFe EMISSAO"
Global Const OP_LOG_NFE_EMISSAO_MANUAL = "NFe EMISSAO MANUAL"
Global Const OP_LOG_DOWNLOAD_DANFE_EM_BATCH = "DownloadDanfeEmBatch"
Global Const OP_LOG_NFE_EMISSAO_TRIANGULAR = "NFe OP TRIANGULAR"

' C�DIGOS PARA CONTRIBUINTE ICMS
Global Const COD_ST_CLIENTE_CONTRIBUINTE_ICMS_INICIAL = 0
Global Const COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO = 1
Global Const COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM = 2
Global Const COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO = 3
    
' C�DIGOS PARA PRODUTOR RURAL
Global Const COD_ST_CLIENTE_PRODUTOR_RURAL_INICIAL = 0
Global Const COD_ST_CLIENTE_PRODUTOR_RURAL_NAO = 1
Global Const COD_ST_CLIENTE_PRODUTOR_RURAL_SIM = 2
    
Const REG_CHAVE_USUARIO_HORARIO_VERAO = "SOFTWARE\PRAGMATICA\Sistema Contratos\AMB\Horario de Verao"
Const REG_CHAVE_USUARIO_PARCELAS_INFO = "SOFTWARE\PRAGMATICA\Sistema Contratos\AMB\InfAdic Parcelas"

Type TIPO_DUAS_COLUNAS
    c1 As String
    c2 As String
    End Type

Type TIPO_TRES_COLUNAS
    c1 As String
    c2 As String
    c3 As String
    End Type

Type TIPO_LINHA_NOTA_FISCAL
    fabricante As String
    produto As String
    descricao As String
    EAN As String
    ncm As String
    NCM_bd As String
    NCM_tela As String
    cst As String
    CST_bd As String
    CST_tela As String
    qtde_total As Long
    valor_total As Currency
    qtde_volumes_total As Long
    peso_total As Single
    cubagem_total As Single
    perc_MVA_ST As Single
    infAdProd As String
    vl_outras_despesas_acessorias As Currency
    cfop As String
    CFOP_formatado As String
    CFOP_tela As String
    CFOP_tela_formatado As String
    ICMS As String
    ICMS_tela As String
    tem_dados_IBPT As Boolean
    percAliqNac As Single
    percAliqImp As Single
    xPed As String
    nItemPed As String
    fcp As String
    End Type
    
Type TIPO_LINHA_NFe_EMISSAO_MANUAL
    fabricante As String
    produto As String
    descricao As String
    descricao_bd As String
    descricao_tela As String
    EAN As String
    ncm As String
    NCM_bd As String
    NCM_tela As String
    cst As String
    CST_bd As String
    CST_tela As String
    qtde As Long
    vl_unitario As Currency
    vl_outras_despesas_acessorias As Currency
    qtde_volumes_total As Long
    peso_total As Single
    cubagem_total As Single
    perc_MVA_ST As Single
    infAdProd As String
    cfop As String
    CFOP_formatado As String
    CFOP_tela As String
    CFOP_tela_formatado As String
    ICMS As String
    ICMS_tela As String
    tem_dados_IBPT As Boolean
    percAliqNac As Single
    percAliqImp As Single
    xPed As String
    nItemPed As String
    uCom As String
    uTrib As String
    fcp As String
    End Type
    
Type TIPO_LISTA_CFOP
    codigo As String
    descricao As String
    End Type

Type TIPO_PEDIDO_CALCULO_PARCELAS_BOLETO
    pedido As String
    vlTotalFamiliaPedidos As Currency
    vlTotalDestePedido As Currency
    razaoValorPedidoFilhote As Double
    tipo_parcelamento As Integer
    av_forma_pagto As Integer
    pc_qtde_parcelas As Integer
    pc_valor_parcela As Currency
    pce_forma_pagto_entrada As Integer
    pce_forma_pagto_prestacao As Integer
    pce_entrada_valor As Currency
    pce_prestacao_qtde As Integer
    pce_prestacao_valor As Currency
    pce_prestacao_periodo As Integer
    pse_forma_pagto_prim_prest As Integer
    pse_forma_pagto_demais_prest As Integer
    pse_prim_prest_valor As Currency
    pse_prim_prest_apos As Integer
    pse_demais_prest_qtde As Integer
    pse_demais_prest_valor As Currency
    pse_demais_prest_periodo As Integer
    pu_forma_pagto As Integer
    pu_valor As Currency
    pu_vencto_apos As Integer
    End Type

Type TIPO_NF_LINHA_DADOS_PARCELA_PAGTO
    intNumDestaParcela As Integer
    intNumTotalParcelas As Integer
    id_forma_pagto As String
    dtVencto As Date
    vlValor As Currency
    strDadosRateio As String
    End Type
    
    
' DADOS COMPLETOS DA NFE
Type TIPO_NFe_IMG
    id As Long
    id_nfe_emitente As Integer
    versao_layout_NFe As String
    NFe_serie_NF As Long
    NFe_numero_NF As Long
    pedido As String
    operacional__email As String
    ide__natOp As String
    ide__indPag As String
    ide__serie As String
    ide__nNF As String
    ide__dEmi As String
    ide__dEmiUTC As String
    ide__dSaiEnt As String
    ide__tpNF As String
    ide__idDest As String
    ide__cMunFG As String
    ide__tpAmb As String
    ide__finNFe As String
    ide__indFinal As String
    ide__indPres As String
    ide__IEST As String
    dest__CNPJ As String
    dest__CPF As String
    dest__xNome As String
    dest__xLgr As String
    dest__nro As String
    dest__xCpl As String
    dest__xBairro As String
    dest__cMun As String
    dest__xMun As String
    dest__UF As String
    dest__CEP As String
    dest__cPais As String
    dest__xPais As String
    dest__fone As String
    dest__IE As String
    dest__ISUF As String
    dest__idEstrangeiro As String
    dest__indIEDest As String
    dest__email As String
    entrega__CNPJ As String
    entrega__CPF As String
    entrega__xLgr As String
    entrega__nro As String
    entrega__xCpl As String
    entrega__xBairro As String
    entrega__cMun As String
    entrega__xMun As String
    entrega__UF As String
    total__vBC As String
    total__vICMS As String
    total__vICMSDeson As String
    total__vFCPUFDest As String
    total__vICMSUFDest As String
    total__vICMSUFRemet As String
    total__vFCP As String
    total__vFCPST As String
    total__vFCPSTRet As String
    total__vIPIDevol As String
    total__vBCST As String
    total__vST As String
    total__vProd As String
    total__vFrete As String
    total__vSeg As String
    total__vDesc As String
    total__vII As String
    total__vIPI As String
    total__vPIS As String
    total__vCOFINS As String
    total__vOutro As String
    total__vNF As String
    total__vTotTrib As String
    transp__modFrete As String
    transporta__CNPJ As String
    transporta__CPF As String
    transporta__xNome As String
    transporta__IE As String
    transporta__xEnder As String
    transporta__xMun As String
    transporta__UF As String
    vol__qVol As String
    vol__esp As String
    vol__marca As String
    vol__nVol As String
    vol__pesoL As String
    vol__pesoB As String
    vol_nLacre As String
    infAdic__infAdFisco As String
    infAdic__infCpl As String
    codigo_retorno_NFe_T1 As String
    msg_retorno_NFe_T1 As String
    End Type
    
Type TIPO_NFe_IMG_ITEM
    id As Long
    id_nfe_imagem As Long
    fabricante As String
    produto As String
    det__nItem As String
    det__cProd As String
    det__cEAN As String
    det__xProd As String
    det__NCM As String
    det__CEST As String
    det__indEscala As String
    det__EXTIPI As String
    det__genero As String
    det__CFOP As String
    det__uCom As String
    det__qCom As String
    det__vUnCom As String
    det__vProd As String
    det__cEANTrib As String
    det__uTrib As String
    det__qTrib As String
    det__vUnTrib As String
    det__vFrete As String
    det__vSeg As String
    det__vDesc As String
    ICMS__orig As String
    ICMS__CST As String
    ICMS__modBC As String
    ICMS__pRedBC As String
    ICMS__vBC As String
    ICMS__pICMS As String
    ICMS__vICMS As String
    ICMS__vICMSDeson As String
    ICMS__modBCST As String
    ICMS__pMVAST As String
    ICMS__pRedBCST As String
    ICMS__vBCST As String
    ICMS__pICMSST As String
    ICMS__vICMSST As String
    PIS__CST As String
    PIS__vBC As String
    PIS__pPIS As String
    PIS__vPIS As String
    PIS__qBCProd As String
    PIS__vAliqProd As String
    COFINS__CST As String
    COFINS__vBC As String
    COFINS__pCOFINS As String
    COFINS__vCOFINS As String
    COFINS__qBCProd As String
    COFINS__vAliqProd As String
    IPI__CST As String
    IPI__clEnq As String
    IPI__CNPJProd As String
    IPI__cSelo As String
    IPI__qSelo As String
    IPI__cEnq As String
    IPI__vBC As String
    IPI__qUnid As String
    IPI__vUnid As String
    IPI__pIPI As String
    IPI__vIPI As String
    det__infAdProd As String
    det__vOutro As String
    det__indTot As String
    det__xPed As String
    det__nItemPed As String
    det__vTotTrib As String
    ICMS__vBCSTRet As String
    ICMS__vICMSSTRet As String
    ICMSUFDest__vBCUFDest As String
    ICMSUFDest__pFCPUFDest As String
    ICMSUFDest__pICMSUFDest As String
    ICMSUFDest__pICMSInter As String
    ICMSUFDest__pICMSInterPart As String
    ICMSUFDest__vFCPUFDest As String
    ICMSUFDest__vICMSUFDest As String
    ICMSUFDest__vICMSUFRemet As String
    End Type

Type TIPO_NFe_IMG_TAG_DUP
    id As Long
    id_nfe_imagem As Long
    nDup As String
    dVenc As String
    vDup As String
    End Type
    
Type TIPO_NFe_IMG_NFe_REFERENCIADA
    id As Long
    id_nfe_imagem As Long
    refNFe As String
    NFe_serie_NF_referenciada As Long
    NFe_numero_NF_referenciada As Long
    End Type
    
Type TIPO_NFe_IMG_PAG
    id As Long
    id_nfe_imagem As Long
    pag__indPag As String
    pag__tPag As String
    pag__vPag As String
    End Type

' DECLARA��ES P/ FUN��ES DE IMPRESS�O
Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
    End Type

Global vCFOPsSemPartilha() As TIPO_LISTA_CFOP

Global vCUFsInscricaoVirtual() As TIPO_DUAS_COLUNAS
    

Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
  
  '�DrawText() Format Flags
    Global Const DT_TOP = &H0
    Global Const DT_LEFT = &H0
    Global Const DT_CENTER = &H1
    Global Const DT_RIGHT = &H2
    Global Const DT_VCENTER = &H4
    Global Const DT_BOTTOM = &H8
    Global Const DT_WORDBREAK = &H10
    Global Const DT_SINGLELINE = &H20
    Global Const DT_EXPANDTABS = &H40
    Global Const DT_TABSTOP = &H80
    Global Const DT_NOCLIP = &H100
    Global Const DT_EXTERNALLEADING = &H200
    Global Const DT_CALCRECT = &H400
    Global Const DT_NOPREFIX = &H800
    Global Const DT_INTERNAL = &H1000
    
    Global Const MM_TEXT = 1
    Global Const MM_TWIPS = 6
    Global Const MM_ISOTROPIC = 7
  

Declare Function ConsisteInscricaoEstadual Lib "DllInscE32.dll" (ByVal Insc As String, ByVal uf As String) As Integer

Function nfe_chave_acesso_ok(ByVal chave As String) As Integer

Dim d As Integer
Dim i As Integer


Const p = "4329876543298765432987654329876543298765432"

    If Trim$(chave) = "" Then Exit Function

    If Len(chave) <> 44 Then Exit Function


'�
'�  VERIFICA O CHECK DIGIT
'�
    d = 0
    For i = 1 To 43
        d = d + Val(Mid$(p, i, 1)) * Val(Mid$(chave, i, 1))
        Next i

    d = 11 - (d Mod 11)
    If d > 9 Then d = 0
    
    If d <> Val(Mid$(chave, 44, 1)) Then
        nfe_chave_acesso_ok = False
        Exit Function
        End If
    
    nfe_chave_acesso_ok = True
     
End Function

Public Sub executa_download_pdf_danfe_periodo(ByRef f_chamador As Form)

'�CONSTANTES
Const NomeDestaRotina = "executa_download_pdf_danfe_periodo()"
Const CHUNK_SIZE = 1000

Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim strLog As String
Dim strLogNF As String
Dim strLogNfSemDadosPdf As String
Dim strDiretorioPdfDanfe As String
Dim strNomeArqDanfe As String
Dim strNomeArqCompletoDanfe As String
Dim strNumeroNfNormalizado As String
Dim strSerieNfNormalizado As String
Dim strNomeEmitente As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String
Dim id_boleto_cedente As Integer
Dim id_boleto_cedente_anterior As Integer
Dim dtInicioSelecionada As Date
Dim dtTerminoSelecionada As Date
Dim iperc As Integer
Dim intQtdeArqDownload As Long
Dim intQtdeNFeSemDadosPdf As Long
Dim intContadorRegistros As Long
Dim intQtdeTotalRegistros As Long
Dim intNfeRetornoSP As Integer
Dim lFileHandle As Long
Dim lngFileSize As Long
Dim lngOffset As Long
Dim bytFile() As Byte
Dim res As Variant
Dim hwnd As Long

'�BANCO DE DADOS
Dim dbcNFe As ADODB.Connection
Dim t_FIN_BOLETO_CEDENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeDanfe As New ADODB.Command
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPDanfe As ADODB.Recordset

    On Error GoTo TDPDP_TRATA_ERRO

    f_PERIODO.Show vbModal, f_chamador

    If Not f_PERIODO.blnResultadoFormOk Then
        aviso_erro "Opera��o de download dos arquivos de DANFE foi cancelada!!"
        Exit Sub
        End If
        
    dtInicioSelecionada = f_PERIODO.dtInicioSelecionada
    dtTerminoSelecionada = f_PERIODO.dtTerminoSelecionada
    
    If dtTerminoSelecionada > Date Then
        aviso_erro "O t�rmino do per�odo informa uma data futura para download dos arquivos de DANFE!!"
        Exit Sub
        End If
    
'   CONFIRMA��O
    s = "Realiza o download dos arquivos de DANFE das NFe's emitidas entre " & Format$(dtInicioSelecionada, FORMATO_DATA) & " e " & Format$(dtTerminoSelecionada, FORMATO_DATA) & "?"
    If Not confirma(s) Then Exit Sub
    
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    
'   CONEX�O AO BD NFE
    Set dbcNFe = New ADODB.Connection
    dbcNFe.CursorLocation = BD_POLITICA_CURSOR
    dbcNFe.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcNFe.CommandTimeout = BD_COMMAND_TIMEOUT
    
'   T_FIN_BOLETO_CEDENTE
    Set t_FIN_BOLETO_CEDENTE = New ADODB.Recordset
    With t_FIN_BOLETO_CEDENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   t_NFe_EMISSAO
    Set t_NFe_EMISSAO = New ADODB.Recordset
    With t_NFe_EMISSAO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   PREPARA COMMAND'S
    cmdNFeSituacao.CommandType = adCmdStoredProc
    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)
    
    cmdNFeDanfe.CommandType = adCmdStoredProc
    cmdNFeDanfe.CommandText = "Proc_NFe_Danfe"
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("Serie", adChar, adParamInput, 3)
    
    s = "SELECT DISTINCT" & _
            " id_boleto_cedente," & _
            " NFe_serie_NF," & _
            " NFe_numero_NF" & _
        " FROM t_NFe_EMISSAO" & _
        " WHERE" & _
            " (dt_emissao >= " & bd_monta_data(dtInicioSelecionada, False) & ")" & _
            " AND (dt_emissao < " & bd_monta_data(dtTerminoSelecionada + 1, False) & ")" & _
            " AND (st_anulado = 0)" & _
            " AND (id_boleto_cedente = " & usuario.emit_id & ")" & _
        " ORDER BY" & _
            " id_boleto_cedente," & _
            " NFe_serie_NF," & _
            " NFe_numero_NF"
    t_NFe_EMISSAO.Open s, dbc, , , adCmdText
    If t_NFe_EMISSAO.EOF Then
        aviso "N�o h� nenhuma NFe emitida entre " & Format$(dtInicioSelecionada, FORMATO_DATA) & " e " & Format$(dtTerminoSelecionada, FORMATO_DATA) & "!!"
        GoSub TDPDP_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If

    strLogNF = ""
    strLogNfSemDadosPdf = ""
    id_boleto_cedente_anterior = -1
    intQtdeTotalRegistros = t_NFe_EMISSAO.RecordCount
    Do While Not t_NFe_EMISSAO.EOF
        DoEvents
        intContadorRegistros = intContadorRegistros + 1
        id_boleto_cedente = CInt(t_NFe_EMISSAO("id_boleto_cedente"))
        If id_boleto_cedente <> id_boleto_cedente_anterior Then
            s = "SELECT" & _
                    " nome_empresa," & _
                    " NFe_T1_servidor_BD," & _
                    " NFe_T1_nome_BD," & _
                    " NFe_T1_usuario_BD," & _
                    " NFe_T1_senha_BD" & _
                " FROM t_FIN_BOLETO_CEDENTE" & _
                " WHERE" & _
                    " (id = " & CStr(id_boleto_cedente) & ")"
            If t_FIN_BOLETO_CEDENTE.State <> adStateClosed Then t_FIN_BOLETO_CEDENTE.Close
            t_FIN_BOLETO_CEDENTE.Open s, dbc, , , adCmdText
            If t_FIN_BOLETO_CEDENTE.EOF Then
                s = "Falha ao localizar o registro em t_FIN_BOLETO_CEDENTE (id=" & CStr(id_boleto_cedente) & ")!!"
                aviso_erro s
                GoSub TDPDP_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
                
            strNomeEmitente = UCase$(Trim$("" & t_FIN_BOLETO_CEDENTE("nome_empresa")))
            strNfeT1ServidorBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_servidor_BD"))
            strNfeT1NomeBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_nome_BD"))
            strNfeT1UsuarioBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_usuario_BD"))
            strNfeT1SenhaCriptografadaBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_senha_BD"))
            
            decodifica_dado strNfeT1SenhaCriptografadaBd, s_aux
            s = "Provider=" & BD_OLEDB_PROVIDER & _
                ";Data Source=" & strNfeT1ServidorBd & _
                ";Initial Catalog=" & strNfeT1NomeBd & _
                ";User Id=" & strNfeT1UsuarioBd & _
                ";Password=" & s_aux
            If dbcNFe.State <> adStateClosed Then dbcNFe.Close
            dbcNFe.Open s
            
            Set cmdNFeSituacao.ActiveConnection = dbcNFe
            Set cmdNFeDanfe.ActiveConnection = dbcNFe
            
            id_boleto_cedente_anterior = id_boleto_cedente
            End If
        
        strNumeroNfNormalizado = NFeFormataNumeroNF(t_NFe_EMISSAO("NFe_numero_NF"))
        strSerieNfNormalizado = NFeFormataSerieNF(t_NFe_EMISSAO("NFe_serie_NF"))
        
    '   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRA��O C/ O SISTEMA DE NFe DA TARGET ONE
        cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
        cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
        Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
        intNfeRetornoSP = rsNFeRetornoSPSituacao("Retorno")
        
        If intNfeRetornoSP = 1 Then
            cmdNFeDanfe.Parameters("NFe") = strNumeroNfNormalizado
            cmdNFeDanfe.Parameters("Serie") = strSerieNfNormalizado
            Set rsNFeRetornoSPDanfe = cmdNFeDanfe.Execute
            If Not rsNFeRetornoSPDanfe.EOF Then
            '   PROGRESSO
                iperc = Int((intContadorRegistros / intQtdeTotalRegistros) * 100)
                aguarde INFO_EXECUTANDO, "copiando DANFE da NFe n� " & strNumeroNfNormalizado & "  (" & CStr(iperc) & "%)"
                
                lngFileSize = rsNFeRetornoSPDanfe("DanfePDF").ActualSize
                
                If lngFileSize <= 0 Then
                    intQtdeNFeSemDadosPdf = intQtdeNFeSemDadosPdf + 1
                    If strLogNfSemDadosPdf <> "" Then strLogNfSemDadosPdf = strLogNfSemDadosPdf & ", "
                    strLogNfSemDadosPdf = strLogNfSemDadosPdf & CStr(id_boleto_cedente) & "/" & strSerieNfNormalizado & "/" & strNumeroNfNormalizado
                    End If
                
                If lngFileSize > 0 Then
                    intQtdeArqDownload = intQtdeArqDownload + 1
                    
                '   LOG
                    If strLogNF <> "" Then strLogNF = strLogNF & ", "
                    strLogNF = strLogNF & CStr(id_boleto_cedente) & "/" & strSerieNfNormalizado & "/" & strNumeroNfNormalizado
                
                '   ARQUIVO DE DANFE
                    strNomeArqDanfe = "NFe_" & strSerieNfNormalizado & "_" & strNumeroNfNormalizado & ".pdf"
                    strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strNomeEmitente
                    
                    If Not DirectoryExists(strDiretorioPdfDanfe, s_erro) Then
                        If Not ForceDirectories(strDiretorioPdfDanfe, s_erro) Then
                            s = "Falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" & strDiretorioPdfDanfe & "): " & s_erro
                            aviso_erro s
                            GoSub TDPDP_FECHA_TABELAS
                            aguarde INFO_NORMAL, m_id
                            Exit Sub
                            End If
                        End If
                    
                    strNomeArqCompletoDanfe = barra_invertida_add(strDiretorioPdfDanfe) & strNomeArqDanfe
                    If FileExists(strNomeArqCompletoDanfe, s_erro) Then
                        If Not FileDelete(strNomeArqCompletoDanfe, s_erro) Then
                            s = "Falha ao tentar apagar o arquivo PDF do DANFE anterior (" & strNomeArqCompletoDanfe & "): " & s_erro
                            aviso_erro s
                            GoSub TDPDP_FECHA_TABELAS
                            aguarde INFO_NORMAL, m_id
                            Exit Sub
                            End If
                        End If
                
                    lFileHandle = FreeFile
                    Open strNomeArqCompletoDanfe For Binary As #lFileHandle
                    lngOffset = 0
                    Do While lngOffset < lngFileSize
                        DoEvents
                        bytFile = rsNFeRetornoSPDanfe("DanfePDF").GetChunk(CHUNK_SIZE)
                        Put #lFileHandle, , bytFile()
                        lngOffset = lngOffset + CHUNK_SIZE
                        Loop
                    
                    Close #lFileHandle
                    End If
                End If
            End If
        
    '   DESALOCA OS RECORDSETS CRIADOS PELA EXECU��O DA STORED PROCEDURE
        bd_desaloca_recordset rsNFeRetornoSPSituacao, True
        bd_desaloca_recordset rsNFeRetornoSPDanfe, True
        
        t_NFe_EMISSAO.MoveNext
        Loop


    strLog = "Download dos arquivos de DANFE das NFe's emitidas entre " & Format$(dtInicioSelecionada, FORMATO_DATA) & " e " & Format$(dtTerminoSelecionada, FORMATO_DATA) & ": " & CStr(intQtdeArqDownload) & " arquivos copiados" & "; " & CStr(intQtdeNFeSemDadosPdf) & " NFe's n�o foram copiadas por n�o possu�rem o PDF armazenado no BD"
    If strLogNF <> "" Then
        strLog = strLog & "; Rela��o de NFe's copiadas: " & strLogNF
        End If

    If strLogNfSemDadosPdf <> "" Then
        strLog = strLog & "; Rela��o de NFe's N�O copiadas: " & strLogNfSemDadosPdf
        End If
    
    Call grava_log(usuario.id, "", "", "", OP_LOG_DOWNLOAD_DANFE_EM_BATCH, strLog)
    
    GoSub TDPDP_FECHA_TABELAS

    aguarde INFO_NORMAL, m_id

    s = "Download dos arquivos de DANFE do per�odo de " & Format$(dtInicioSelecionada, FORMATO_DATA) & " a " & Format$(dtTerminoSelecionada, FORMATO_DATA) & " foi conclu�do com sucesso!!" & Chr(13) & _
        CStr(intQtdeArqDownload) & " arquivos foram copiados!!"
    If intQtdeNFeSemDadosPdf > 0 Then
        s = s & Chr(13) & _
            CStr(intQtdeNFeSemDadosPdf) & " NFe's n�o foram copiadas por n�o possu�rem o PDF armazenado no BD!!"
        End If
    aviso s
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TDPDP_FECHA_TABELAS:
'===================
'   RECORDSETS
    bd_desaloca_recordset t_FIN_BOLETO_CEDENTE, True
    bd_desaloca_recordset t_NFe_EMISSAO, True
    bd_desaloca_recordset rsNFeRetornoSPSituacao, True
    bd_desaloca_recordset rsNFeRetornoSPDanfe, True
    
'   COMMAND
    bd_desaloca_command cmdNFeSituacao
    bd_desaloca_command cmdNFeDanfe
    
'   CONNECTION
    If Not (dbcNFe Is Nothing) Then
        If dbcNFe.State <> adStateClosed Then dbcNFe.Close
        Set dbcNFe = Nothing
        End If
        
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TDPDP_TRATA_ERRO:
'================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    GoSub TDPDP_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub

Public Sub executa_download_pdf_danfe_periodo_parametro_emitente(ByRef f_chamador As Form)

'�CONSTANTES
Const NomeDestaRotina = "executa_download_pdf_danfe_periodo_parametro_emitente()"
Const CHUNK_SIZE = 1000

Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim strLog As String
Dim strLogNF As String
Dim strLogNfSemDadosPdf As String
Dim strDiretorioPdfDanfe As String
Dim strNomeArqDanfe As String
Dim strNomeArqCompletoDanfe As String
Dim strNumeroNfNormalizado As String
Dim strSerieNfNormalizado As String
Dim strNomeEmitente As String
Dim strPastaEmitente As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String
Dim strUF As String
Dim id_nfe_emitente As Integer
Dim id_nfe_emitente_anterior As Integer
Dim dtInicioSelecionada As Date
Dim dtTerminoSelecionada As Date
Dim iperc As Integer
Dim intQtdeArqDownload As Long
Dim intQtdeNFeSemDadosPdf As Long
Dim intContadorRegistros As Long
Dim intQtdeTotalRegistros As Long
Dim intNfeRetornoSP As Integer
Dim lFileHandle As Long
Dim lngFileSize As Long
Dim lngOffset As Long
Dim bytFile() As Byte
Dim res As Variant
Dim hwnd As Long

'�BANCO DE DADOS
Dim dbcNFe As ADODB.Connection
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeDanfe As New ADODB.Command
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPDanfe As ADODB.Recordset

    On Error GoTo TDPDPPE_TRATA_ERRO

    f_PERIODO.Show vbModal, f_chamador

    If Not f_PERIODO.blnResultadoFormOk Then
        aviso_erro "Opera��o de download dos arquivos de DANFE foi cancelada!!"
        Exit Sub
        End If
        
    dtInicioSelecionada = f_PERIODO.dtInicioSelecionada
    dtTerminoSelecionada = f_PERIODO.dtTerminoSelecionada
    strUF = f_PERIODO.strUFSelecionada
    
    If dtTerminoSelecionada > Date Then
        aviso_erro "O t�rmino do per�odo informa uma data futura para download dos arquivos de DANFE!!"
        Exit Sub
        End If
    
'   CONFIRMA��O
    s = "Realiza o download dos arquivos de DANFE das NFe's emitidas entre " & Format$(dtInicioSelecionada, FORMATO_DATA) & " e " & Format$(dtTerminoSelecionada, FORMATO_DATA) & "?"
    If Not confirma(s) Then Exit Sub
    
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    
'   CONEX�O AO BD NFE
    Set dbcNFe = New ADODB.Connection
    dbcNFe.CursorLocation = BD_POLITICA_CURSOR
    dbcNFe.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcNFe.CommandTimeout = BD_COMMAND_TIMEOUT
    
'   T_NFE_EMITENTE
    Set t_NFE_EMITENTE = New ADODB.Recordset
    With t_NFE_EMITENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   t_NFe_EMISSAO
    Set t_NFe_EMISSAO = New ADODB.Recordset
    With t_NFe_EMISSAO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   PREPARA COMMAND'S
    cmdNFeSituacao.CommandType = adCmdStoredProc
    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)
    
    cmdNFeDanfe.CommandType = adCmdStoredProc
    cmdNFeDanfe.CommandText = "Proc_NFe_Danfe"
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("Serie", adChar, adParamInput, 3)
    
    s_aux = ""
    If strUF <> "" Then
        s_aux = " AND (EXISTS (" & _
            " SELECT 1" & _
            " FROM t_NFe_IMAGEM" & _
            " WHERE (t_NFE_EMISSAO.NFe_numero_NF=t_NFE_IMAGEM.NFe_numero_NF) " & _
            " AND (t_NFE_EMISSAO.NFe_serie_NF=t_NFE_IMAGEM.NFe_serie_NF) " & _
            " AND (t_NFE_IMAGEM.dest__UF = '" & strUF & "')))"
        End If
        
    s = "SELECT DISTINCT" & _
            " id_nfe_emitente," & _
            " NFe_serie_NF," & _
            " NFe_numero_NF" & _
        " FROM t_NFe_EMISSAO" & _
        " WHERE" & _
            " (dt_emissao >= " & bd_monta_data(dtInicioSelecionada, False) & ")" & _
            " AND (dt_emissao < " & bd_monta_data(dtTerminoSelecionada + 1, False) & ")" & _
            " AND (st_anulado = 0)" & _
            " AND (id_nfe_emitente = " & usuario.emit_id & ")" & _
            s_aux & _
        " ORDER BY" & _
            " id_nfe_emitente," & _
            " NFe_serie_NF," & _
            " NFe_numero_NF"
    t_NFe_EMISSAO.Open s, dbc, , , adCmdText
    If t_NFe_EMISSAO.EOF Then
        aviso "N�o h� nenhuma NFe emitida entre " & Format$(dtInicioSelecionada, FORMATO_DATA) & " e " & Format$(dtTerminoSelecionada, FORMATO_DATA) & "!!"
        GoSub TDPDPPE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If

    strLogNF = ""
    strLogNfSemDadosPdf = ""
    id_nfe_emitente_anterior = -1
    intQtdeTotalRegistros = t_NFe_EMISSAO.RecordCount
    Do While Not t_NFe_EMISSAO.EOF
        DoEvents
        intContadorRegistros = intContadorRegistros + 1
        id_nfe_emitente = CInt(t_NFe_EMISSAO("id_nfe_emitente"))
        If id_nfe_emitente <> id_nfe_emitente_anterior Then
            s = "SELECT" & _
                    " razao_social," & _
                    " cnpj," & _
                    " apelido," & _
                    " NFe_T1_servidor_BD," & _
                    " NFe_T1_nome_BD," & _
                    " NFe_T1_usuario_BD," & _
                    " NFe_T1_senha_BD" & _
                " FROM t_NFE_EMITENTE" & _
                " WHERE" & _
                    " (id = " & CStr(id_nfe_emitente) & ")"
            If t_NFE_EMITENTE.State <> adStateClosed Then t_NFE_EMITENTE.Close
            t_NFE_EMITENTE.Open s, dbc, , , adCmdText
            If t_NFE_EMITENTE.EOF Then
                s = "Falha ao localizar o registro em t_NFE_EMITENTE (id=" & CStr(id_nfe_emitente) & ")!!"
                aviso_erro s
                GoSub TDPDPPE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
                
            strNomeEmitente = UCase$(Trim$("" & t_NFE_EMITENTE("razao_social")))
            strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
            strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
            strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
            strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
            'novo padr�o de nome da pasta para DANFEs: <cnpj>-<apelido_com_underlines_substituindo_barras>
            '(ex: 23209013000332-DIS_ES)
            strPastaEmitente = Trim$("" & t_NFE_EMITENTE("cnpj"))
            strPastaEmitente = retorna_so_digitos(strPastaEmitente)
            strPastaEmitente = strPastaEmitente & "-" & Trim$("" & t_NFE_EMITENTE("apelido"))
            strPastaEmitente = substitui_caracteres(strPastaEmitente, "/", "_")
            
            decodifica_dado strNfeT1SenhaCriptografadaBd, s_aux
            s = "Provider=" & BD_OLEDB_PROVIDER & _
                ";Data Source=" & strNfeT1ServidorBd & _
                ";Initial Catalog=" & strNfeT1NomeBd & _
                ";User Id=" & strNfeT1UsuarioBd & _
                ";Password=" & s_aux
            If dbcNFe.State <> adStateClosed Then dbcNFe.Close
            dbcNFe.Open s
            
            Set cmdNFeSituacao.ActiveConnection = dbcNFe
            Set cmdNFeDanfe.ActiveConnection = dbcNFe
            
            id_nfe_emitente_anterior = id_nfe_emitente
            End If
        
        strNumeroNfNormalizado = NFeFormataNumeroNF(t_NFe_EMISSAO("NFe_numero_NF"))
        strSerieNfNormalizado = NFeFormataSerieNF(t_NFe_EMISSAO("NFe_serie_NF"))
        
    '   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRA��O C/ O SISTEMA DE NFe DA TARGET ONE
        cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
        cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
        Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
        intNfeRetornoSP = rsNFeRetornoSPSituacao("Retorno")
        
        If intNfeRetornoSP = 1 Then
            cmdNFeDanfe.Parameters("NFe") = strNumeroNfNormalizado
            cmdNFeDanfe.Parameters("Serie") = strSerieNfNormalizado
            Set rsNFeRetornoSPDanfe = cmdNFeDanfe.Execute
            If Not rsNFeRetornoSPDanfe.EOF Then
            '   PROGRESSO
                iperc = Int((intContadorRegistros / intQtdeTotalRegistros) * 100)
                aguarde INFO_EXECUTANDO, "copiando DANFE da NFe n� " & strNumeroNfNormalizado & "  (" & CStr(iperc) & "%)"
                
                lngFileSize = rsNFeRetornoSPDanfe("DanfePDF").ActualSize
                
                If lngFileSize <= 0 Then
                    intQtdeNFeSemDadosPdf = intQtdeNFeSemDadosPdf + 1
                    If strLogNfSemDadosPdf <> "" Then strLogNfSemDadosPdf = strLogNfSemDadosPdf & ", "
                    strLogNfSemDadosPdf = strLogNfSemDadosPdf & CStr(id_nfe_emitente) & "/" & strSerieNfNormalizado & "/" & strNumeroNfNormalizado
                    End If
                
                If lngFileSize > 0 Then
                    intQtdeArqDownload = intQtdeArqDownload + 1
                    
                '   LOG
                    If strLogNF <> "" Then strLogNF = strLogNF & ", "
                    strLogNF = strLogNF & CStr(id_nfe_emitente) & "/" & strSerieNfNormalizado & "/" & strNumeroNfNormalizado
                
                '   ARQUIVO DE DANFE
                    strNomeArqDanfe = "NFe_" & strSerieNfNormalizado & "_" & strNumeroNfNormalizado & ".pdf"
                    strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strPastaEmitente
                    
                    If Not DirectoryExists(strDiretorioPdfDanfe, s_erro) Then
                        If Not ForceDirectories(strDiretorioPdfDanfe, s_erro) Then
                            s = "Falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" & strDiretorioPdfDanfe & "): " & s_erro
                            aviso_erro s
                            GoSub TDPDPPE_FECHA_TABELAS
                            aguarde INFO_NORMAL, m_id
                            Exit Sub
                            End If
                        End If
                    
                    strNomeArqCompletoDanfe = barra_invertida_add(strDiretorioPdfDanfe) & strNomeArqDanfe
                    If FileExists(strNomeArqCompletoDanfe, s_erro) Then
                        If Not FileDelete(strNomeArqCompletoDanfe, s_erro) Then
                            s = "Falha ao tentar apagar o arquivo PDF do DANFE anterior (" & strNomeArqCompletoDanfe & "): " & s_erro
                            aviso_erro s
                            GoSub TDPDPPE_FECHA_TABELAS
                            aguarde INFO_NORMAL, m_id
                            Exit Sub
                            End If
                        End If
                
                    lFileHandle = FreeFile
                    Open strNomeArqCompletoDanfe For Binary As #lFileHandle
                    lngOffset = 0
                    Do While lngOffset < lngFileSize
                        DoEvents
                        bytFile = rsNFeRetornoSPDanfe("DanfePDF").GetChunk(CHUNK_SIZE)
                        Put #lFileHandle, , bytFile()
                        lngOffset = lngOffset + CHUNK_SIZE
                        Loop
                    
                    Close #lFileHandle
                    End If
                End If
            End If
        
    '   DESALOCA OS RECORDSETS CRIADOS PELA EXECU��O DA STORED PROCEDURE
        bd_desaloca_recordset rsNFeRetornoSPSituacao, True
        bd_desaloca_recordset rsNFeRetornoSPDanfe, True
        
        t_NFe_EMISSAO.MoveNext
        Loop


    strLog = "Download dos arquivos de DANFE das NFe's emitidas entre " & Format$(dtInicioSelecionada, FORMATO_DATA) & " e " & Format$(dtTerminoSelecionada, FORMATO_DATA) & ": " & CStr(intQtdeArqDownload) & " arquivos copiados" & "; " & CStr(intQtdeNFeSemDadosPdf) & " NFe's n�o foram copiadas por n�o possu�rem o PDF armazenado no BD"
    If strLogNF <> "" Then
        strLog = strLog & "; Rela��o de NFe's copiadas: " & strLogNF
        End If

    If strLogNfSemDadosPdf <> "" Then
        strLog = strLog & "; Rela��o de NFe's N�O copiadas: " & strLogNfSemDadosPdf
        End If
    
    Call grava_log(usuario.id, "", "", "", OP_LOG_DOWNLOAD_DANFE_EM_BATCH, strLog)
    
    GoSub TDPDPPE_FECHA_TABELAS

    aguarde INFO_NORMAL, m_id

    s = "Download dos arquivos de DANFE do per�odo de " & Format$(dtInicioSelecionada, FORMATO_DATA) & " a " & Format$(dtTerminoSelecionada, FORMATO_DATA) & " foi conclu�do com sucesso!!" & Chr(13) & _
        CStr(intQtdeArqDownload) & " arquivos foram copiados!!"
    If intQtdeNFeSemDadosPdf > 0 Then
        s = s & Chr(13) & _
            CStr(intQtdeNFeSemDadosPdf) & " NFe's n�o foram copiadas por n�o possu�rem o PDF armazenado no BD!!"
        End If
    aviso s
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TDPDPPE_FECHA_TABELAS:
'===================
'   RECORDSETS
    bd_desaloca_recordset t_NFE_EMITENTE, True
    bd_desaloca_recordset t_NFe_EMISSAO, True
    bd_desaloca_recordset rsNFeRetornoSPSituacao, True
    bd_desaloca_recordset rsNFeRetornoSPDanfe, True
    
'   COMMAND
    bd_desaloca_command cmdNFeSituacao
    bd_desaloca_command cmdNFeDanfe
    
'   CONNECTION
    If Not (dbcNFe Is Nothing) Then
        If dbcNFe.State <> adStateClosed Then dbcNFe.Close
        Set dbcNFe = Nothing
        End If
        
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TDPDPPE_TRATA_ERRO:
'================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    GoSub TDPDPPE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub


Function cst_converte_codigo_entrada_para_saida(ByVal cst_entrada As String) As String
Dim cst_saida As String

    cst_converte_codigo_entrada_para_saida = ""
    
    cst_entrada = Trim$("" & cst_entrada)
    If cst_entrada = "" Then Exit Function
    
    Select Case cst_entrada
        Case "010"
            cst_saida = "060"
        Case "100"
            cst_saida = "200"
        Case "110"
            cst_saida = "260"
        Case Else
            cst_saida = cst_entrada
        End Select
    
    cst_converte_codigo_entrada_para_saida = cst_saida
End Function


Public Sub executa_download_pdf_danfe(ByRef f_chamador As Form)

'�CONSTANTES
Const NomeDestaRotina = "executa_download_pdf_danfe()"
Const CHUNK_SIZE = 1000

Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim strLog As String
Dim strLogNF As String
Dim strLogNfSemDadosPdf As String
Dim strDiretorioPdfDanfe As String
Dim strNomeArqDanfe As String
Dim strNomeArqCompletoDanfe As String
Dim strNumeroNfNormalizado As String
Dim strSerieNfNormalizado As String
Dim strNomeEmitente As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String
Dim id_boleto_cedente As Integer
Dim id_boleto_cedente_anterior As Integer
Dim dtSelecionada As Date
Dim iperc As Integer
Dim intQtdeArqDownload As Long
Dim intQtdeNFeSemDadosPdf As Long
Dim intContadorRegistros As Long
Dim intQtdeTotalRegistros As Long
Dim intNfeRetornoSP As Integer
Dim lFileHandle As Long
Dim lngFileSize As Long
Dim lngOffset As Long
Dim bytFile() As Byte
Dim res As Variant
Dim hwnd As Long

'�BANCO DE DADOS
Dim dbcNFe As ADODB.Connection
Dim t_FIN_BOLETO_CEDENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeDanfe As New ADODB.Command
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPDanfe As ADODB.Recordset

    On Error GoTo TDPD_TRATA_ERRO

    f_DATA.Show vbModal, f_chamador

    If Not f_DATA.blnResultadoFormOk Then
        aviso_erro "Opera��o de download dos arquivos de DANFE foi cancelada!!"
        Exit Sub
        End If
        
    dtSelecionada = f_DATA.dtDataSelecionada
    
    If dtSelecionada > Date Then
        aviso_erro "Foi informada uma data futura para download dos arquivos de DANFE!!"
        Exit Sub
        End If
    
'   CONFIRMA��O
    s = "Realiza o download dos arquivos de DANFE das NFe's emitidas em " & Format$(dtSelecionada, FORMATO_DATA) & "?"
    If Not confirma(s) Then Exit Sub
    
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    
'   CONEX�O AO BD NFE
    Set dbcNFe = New ADODB.Connection
    dbcNFe.CursorLocation = BD_POLITICA_CURSOR
    dbcNFe.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcNFe.CommandTimeout = BD_COMMAND_TIMEOUT
    
'   T_FIN_BOLETO_CEDENTE
    Set t_FIN_BOLETO_CEDENTE = New ADODB.Recordset
    With t_FIN_BOLETO_CEDENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   t_NFe_EMISSAO
    Set t_NFe_EMISSAO = New ADODB.Recordset
    With t_NFe_EMISSAO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   PREPARA COMMAND'S
    cmdNFeSituacao.CommandType = adCmdStoredProc
    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)
    
    cmdNFeDanfe.CommandType = adCmdStoredProc
    cmdNFeDanfe.CommandText = "Proc_NFe_Danfe"
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("Serie", adChar, adParamInput, 3)
    
    s = "SELECT DISTINCT" & _
            " id_boleto_cedente," & _
            " NFe_serie_NF," & _
            " NFe_numero_NF" & _
        " FROM t_NFe_EMISSAO" & _
        " WHERE" & _
            " (dt_emissao = " & bd_monta_data(dtSelecionada, False) & ")" & _
            " AND (st_anulado = 0)" & _
        " ORDER BY" & _
            " id_boleto_cedente," & _
            " NFe_serie_NF," & _
            " NFe_numero_NF"
    t_NFe_EMISSAO.Open s, dbc, , , adCmdText
    If t_NFe_EMISSAO.EOF Then
        aviso "N�o h� nenhuma NFe emitida em " & Format$(dtSelecionada, FORMATO_DATA) & "!!"
        GoSub TDPD_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If

    strLogNF = ""
    strLogNfSemDadosPdf = ""
    id_boleto_cedente_anterior = -1
    intQtdeTotalRegistros = t_NFe_EMISSAO.RecordCount
    Do While Not t_NFe_EMISSAO.EOF
        DoEvents
        intContadorRegistros = intContadorRegistros + 1
        id_boleto_cedente = CInt(t_NFe_EMISSAO("id_boleto_cedente"))
        If id_boleto_cedente <> id_boleto_cedente_anterior Then
            s = "SELECT" & _
                    " nome_empresa," & _
                    " NFe_T1_servidor_BD," & _
                    " NFe_T1_nome_BD," & _
                    " NFe_T1_usuario_BD," & _
                    " NFe_T1_senha_BD" & _
                " FROM t_FIN_BOLETO_CEDENTE" & _
                " WHERE" & _
                    " (id = " & CStr(id_boleto_cedente) & ")"
            If t_FIN_BOLETO_CEDENTE.State <> adStateClosed Then t_FIN_BOLETO_CEDENTE.Close
            t_FIN_BOLETO_CEDENTE.Open s, dbc, , , adCmdText
            If t_FIN_BOLETO_CEDENTE.EOF Then
                s = "Falha ao localizar o registro em t_FIN_BOLETO_CEDENTE (id=" & CStr(id_boleto_cedente) & ")!!"
                aviso_erro s
                GoSub TDPD_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
                
            strNomeEmitente = UCase$(Trim$("" & t_FIN_BOLETO_CEDENTE("nome_empresa")))
            strNfeT1ServidorBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_servidor_BD"))
            strNfeT1NomeBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_nome_BD"))
            strNfeT1UsuarioBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_usuario_BD"))
            strNfeT1SenhaCriptografadaBd = Trim$("" & t_FIN_BOLETO_CEDENTE("NFe_T1_senha_BD"))
            
            decodifica_dado strNfeT1SenhaCriptografadaBd, s_aux
            s = "Provider=" & BD_OLEDB_PROVIDER & _
                ";Data Source=" & strNfeT1ServidorBd & _
                ";Initial Catalog=" & strNfeT1NomeBd & _
                ";User Id=" & strNfeT1UsuarioBd & _
                ";Password=" & s_aux
            If dbcNFe.State <> adStateClosed Then dbcNFe.Close
            dbcNFe.Open s
            
            Set cmdNFeSituacao.ActiveConnection = dbcNFe
            Set cmdNFeDanfe.ActiveConnection = dbcNFe
            
            id_boleto_cedente_anterior = id_boleto_cedente
            End If
        
        strNumeroNfNormalizado = NFeFormataNumeroNF(t_NFe_EMISSAO("NFe_numero_NF"))
        strSerieNfNormalizado = NFeFormataSerieNF(t_NFe_EMISSAO("NFe_serie_NF"))
        
    '   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRA��O C/ O SISTEMA DE NFe DA TARGET ONE
        cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
        cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
        Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
        intNfeRetornoSP = rsNFeRetornoSPSituacao("Retorno")
        
        If intNfeRetornoSP = 1 Then
            cmdNFeDanfe.Parameters("NFe") = strNumeroNfNormalizado
            cmdNFeDanfe.Parameters("Serie") = strSerieNfNormalizado
            Set rsNFeRetornoSPDanfe = cmdNFeDanfe.Execute
            If Not rsNFeRetornoSPDanfe.EOF Then
            '   PROGRESSO
                iperc = Int((intContadorRegistros / intQtdeTotalRegistros) * 100)
                aguarde INFO_EXECUTANDO, "copiando DANFE da NFe n� " & strNumeroNfNormalizado & "  (" & CStr(iperc) & "%)"
                
                lngFileSize = rsNFeRetornoSPDanfe("DanfePDF").ActualSize
                
                If lngFileSize <= 0 Then
                    intQtdeNFeSemDadosPdf = intQtdeNFeSemDadosPdf + 1
                    If strLogNfSemDadosPdf <> "" Then strLogNfSemDadosPdf = strLogNfSemDadosPdf & ", "
                    strLogNfSemDadosPdf = strLogNfSemDadosPdf & CStr(id_boleto_cedente) & "/" & strSerieNfNormalizado & "/" & strNumeroNfNormalizado
                    End If
                
                If lngFileSize > 0 Then
                    intQtdeArqDownload = intQtdeArqDownload + 1
                    
                '   LOG
                    If strLogNF <> "" Then strLogNF = strLogNF & ", "
                    strLogNF = strLogNF & CStr(id_boleto_cedente) & "/" & strSerieNfNormalizado & "/" & strNumeroNfNormalizado
                
                '   ARQUIVO DE DANFE
                    strNomeArqDanfe = "NFe_" & strSerieNfNormalizado & "_" & strNumeroNfNormalizado & ".pdf"
                    strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strNomeEmitente
                    
                    If Not DirectoryExists(strDiretorioPdfDanfe, s_erro) Then
                        If Not ForceDirectories(strDiretorioPdfDanfe, s_erro) Then
                            s = "Falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" & strDiretorioPdfDanfe & "): " & s_erro
                            aviso_erro s
                            GoSub TDPD_FECHA_TABELAS
                            aguarde INFO_NORMAL, m_id
                            Exit Sub
                            End If
                        End If
                    
                    strNomeArqCompletoDanfe = barra_invertida_add(strDiretorioPdfDanfe) & strNomeArqDanfe
                    If FileExists(strNomeArqCompletoDanfe, s_erro) Then
                        If Not FileDelete(strNomeArqCompletoDanfe, s_erro) Then
                            s = "Falha ao tentar apagar o arquivo PDF do DANFE anterior (" & strNomeArqCompletoDanfe & "): " & s_erro
                            aviso_erro s
                            GoSub TDPD_FECHA_TABELAS
                            aguarde INFO_NORMAL, m_id
                            Exit Sub
                            End If
                        End If
                    
                    lFileHandle = FreeFile
                    Open strNomeArqCompletoDanfe For Binary As #lFileHandle
                    lngOffset = 0
                    Do While lngOffset < lngFileSize
                        DoEvents
                        bytFile = rsNFeRetornoSPDanfe("DanfePDF").GetChunk(CHUNK_SIZE)
                        Put #lFileHandle, , bytFile()
                        lngOffset = lngOffset + CHUNK_SIZE
                        Loop
                    
                    Close #lFileHandle
                    End If
                End If
            End If
        
    '   DESALOCA OS RECORDSETS CRIADOS PELA EXECU��O DA STORED PROCEDURE
        bd_desaloca_recordset rsNFeRetornoSPSituacao, True
        bd_desaloca_recordset rsNFeRetornoSPDanfe, True
        
        t_NFe_EMISSAO.MoveNext
        Loop


    strLog = "Download dos arquivos de DANFE das NFe's emitidas em " & Format$(dtSelecionada, FORMATO_DATA) & ": " & CStr(intQtdeArqDownload) & " arquivos copiados" & "; " & CStr(intQtdeNFeSemDadosPdf) & " NFe's n�o foram copiadas por n�o possu�rem o PDF armazenado no BD"
    If strLogNF <> "" Then
        strLog = strLog & "; Rela��o de NFe's copiadas: " & strLogNF
        End If
    
    If strLogNfSemDadosPdf <> "" Then
        strLog = strLog & "; Rela��o de NFe's N�O copiadas: " & strLogNfSemDadosPdf
        End If
    
    Call grava_log(usuario.id, "", "", "", OP_LOG_DOWNLOAD_DANFE_EM_BATCH, strLog)
    
    GoSub TDPD_FECHA_TABELAS

    aguarde INFO_NORMAL, m_id

    s = "Download dos arquivos de DANFE foi conclu�do com sucesso!!" & Chr(13) & _
        CStr(intQtdeArqDownload) & " arquivos foram copiados!!"
    If intQtdeNFeSemDadosPdf > 0 Then
        s = s & Chr(13) & _
            CStr(intQtdeNFeSemDadosPdf) & " NFe's n�o foram copiadas por n�o possu�rem o PDF armazenado no BD!!"
        End If
    aviso s
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TDPD_FECHA_TABELAS:
'==================
'   RECORDSETS
    bd_desaloca_recordset t_FIN_BOLETO_CEDENTE, True
    bd_desaloca_recordset t_NFe_EMISSAO, True
    bd_desaloca_recordset rsNFeRetornoSPSituacao, True
    bd_desaloca_recordset rsNFeRetornoSPDanfe, True
    
'   COMMAND
    bd_desaloca_command cmdNFeSituacao
    bd_desaloca_command cmdNFeDanfe
    
'   CONNECTION
    If Not (dbcNFe Is Nothing) Then
        If dbcNFe.State <> adStateClosed Then dbcNFe.Close
        Set dbcNFe = Nothing
        End If
        
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TDPD_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    GoSub TDPD_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub

Public Sub executa_download_pdf_danfe_parametro_emitente(ByRef f_chamador As Form)

'�CONSTANTES
Const NomeDestaRotina = "executa_download_pdf_danfe_parametro_emitente()"
Const CHUNK_SIZE = 1000

Dim s As String
Dim s_aux As String
Dim s_erro As String
Dim strLog As String
Dim strLogNF As String
Dim strLogNfSemDadosPdf As String
Dim strDiretorioPdfDanfe As String
Dim strNomeArqDanfe As String
Dim strNomeArqCompletoDanfe As String
Dim strNumeroNfNormalizado As String
Dim strSerieNfNormalizado As String
Dim strNomeEmitente As String
Dim strPastaEmitente As String
Dim strNfeT1ServidorBd As String
Dim strNfeT1NomeBd As String
Dim strNfeT1UsuarioBd As String
Dim strNfeT1SenhaCriptografadaBd As String
Dim id_nfe_emitente As Integer
Dim id_nfe_emitente_anterior As Integer
Dim dtSelecionada As Date
Dim iperc As Integer
Dim intQtdeArqDownload As Long
Dim intQtdeNFeSemDadosPdf As Long
Dim intContadorRegistros As Long
Dim intQtdeTotalRegistros As Long
Dim intNfeRetornoSP As Integer
Dim lFileHandle As Long
Dim lngFileSize As Long
Dim lngOffset As Long
Dim bytFile() As Byte
Dim res As Variant
Dim hwnd As Long

'�BANCO DE DADOS
Dim dbcNFe As ADODB.Connection
Dim t_NFE_EMITENTE As ADODB.Recordset
Dim t_NFe_EMISSAO As ADODB.Recordset
Dim cmdNFeSituacao As New ADODB.Command
Dim cmdNFeDanfe As New ADODB.Command
Dim rsNFeRetornoSPSituacao As ADODB.Recordset
Dim rsNFeRetornoSPDanfe As ADODB.Recordset

    On Error GoTo TDPDPE_TRATA_ERRO

    f_DATA.Show vbModal, f_chamador

    If Not f_DATA.blnResultadoFormOk Then
        aviso_erro "Opera��o de download dos arquivos de DANFE foi cancelada!!"
        Exit Sub
        End If
        
    dtSelecionada = f_DATA.dtDataSelecionada
    
    If dtSelecionada > Date Then
        aviso_erro "Foi informada uma data futura para download dos arquivos de DANFE!!"
        Exit Sub
        End If
    
'   CONFIRMA��O
    s = "Realiza o download dos arquivos de DANFE das NFe's emitidas em " & Format$(dtSelecionada, FORMATO_DATA) & "?"
    If Not confirma(s) Then Exit Sub
    
    aguarde INFO_EXECUTANDO, "consultando banco de dados"
    
'   CONEX�O AO BD NFE
    Set dbcNFe = New ADODB.Connection
    dbcNFe.CursorLocation = BD_POLITICA_CURSOR
    dbcNFe.ConnectionTimeout = BD_CONNECTION_TIMEOUT
    dbcNFe.CommandTimeout = BD_COMMAND_TIMEOUT
    
'   t_NFE_EMITENTE
    Set t_NFE_EMITENTE = New ADODB.Recordset
    With t_NFE_EMITENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   t_NFe_EMISSAO
    Set t_NFe_EMISSAO = New ADODB.Recordset
    With t_NFe_EMISSAO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   PREPARA COMMAND'S
    cmdNFeSituacao.CommandType = adCmdStoredProc
    cmdNFeSituacao.CommandText = "Proc_NFe_Integracao_ConsultaEmissao"
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeSituacao.Parameters.Append cmdNFeSituacao.CreateParameter("Serie", adChar, adParamInput, 3)
    
    cmdNFeDanfe.CommandType = adCmdStoredProc
    cmdNFeDanfe.CommandText = "Proc_NFe_Danfe"
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("NFe", adChar, adParamInput, 9)
    cmdNFeDanfe.Parameters.Append cmdNFeDanfe.CreateParameter("Serie", adChar, adParamInput, 3)
    
    s = "SELECT DISTINCT" & _
            " id_nfe_emitente," & _
            " NFe_serie_NF," & _
            " NFe_numero_NF" & _
        " FROM t_NFe_EMISSAO" & _
        " WHERE" & _
            " (dt_emissao = " & bd_monta_data(dtSelecionada, False) & ")" & _
            " AND (st_anulado = 0)" & _
        " ORDER BY" & _
            " id_nfe_emitente," & _
            " NFe_serie_NF," & _
            " NFe_numero_NF"
    t_NFe_EMISSAO.Open s, dbc, , , adCmdText
    If t_NFe_EMISSAO.EOF Then
        aviso "N�o h� nenhuma NFe emitida em " & Format$(dtSelecionada, FORMATO_DATA) & "!!"
        GoSub TDPDPE_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Sub
        End If

    strLogNF = ""
    strLogNfSemDadosPdf = ""
    id_nfe_emitente_anterior = -1
    intQtdeTotalRegistros = t_NFe_EMISSAO.RecordCount
    Do While Not t_NFe_EMISSAO.EOF
        DoEvents
        intContadorRegistros = intContadorRegistros + 1
        id_nfe_emitente = CInt(t_NFe_EMISSAO("id_nfe_emitente"))
        If id_nfe_emitente <> id_nfe_emitente_anterior Then
            s = "SELECT" & _
                    " razao_social," & _
                    " NFe_T1_servidor_BD," & _
                    " NFe_T1_nome_BD," & _
                    " NFe_T1_usuario_BD," & _
                    " NFe_T1_senha_BD" & _
                " FROM t_NFE_EMITENTE" & _
                " WHERE" & _
                    " (id = " & CStr(id_nfe_emitente) & ")"
            If t_NFE_EMITENTE.State <> adStateClosed Then t_NFE_EMITENTE.Close
            t_NFE_EMITENTE.Open s, dbc, , , adCmdText
            If t_NFE_EMITENTE.EOF Then
                s = "Falha ao localizar o registro em t_NFE_EMITENTE (id=" & CStr(id_nfe_emitente) & ")!!"
                aviso_erro s
                GoSub TDPDPE_FECHA_TABELAS
                aguarde INFO_NORMAL, m_id
                Exit Sub
                End If
                
            strNomeEmitente = UCase$(Trim$("" & t_NFE_EMITENTE("razao_social")))
            strNfeT1ServidorBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_servidor_BD"))
            strNfeT1NomeBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_nome_BD"))
            strNfeT1UsuarioBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_usuario_BD"))
            strNfeT1SenhaCriptografadaBd = Trim$("" & t_NFE_EMITENTE("NFe_T1_senha_BD"))
            'novo padr�o de nome da pasta para DANFEs: <cnpj>-<apelido_com_underlines_substituindo_barras>
            '(ex: 23209013000332-DIS_ES)
            strPastaEmitente = Trim$("" & t_NFE_EMITENTE("cnpj"))
            strPastaEmitente = retorna_so_digitos(strPastaEmitente)
            strPastaEmitente = strPastaEmitente & "-" & Trim$("" & t_NFE_EMITENTE("apelido"))
            strPastaEmitente = substitui_caracteres(strPastaEmitente, "/", "_")
            
            decodifica_dado strNfeT1SenhaCriptografadaBd, s_aux
            s = "Provider=" & BD_OLEDB_PROVIDER & _
                ";Data Source=" & strNfeT1ServidorBd & _
                ";Initial Catalog=" & strNfeT1NomeBd & _
                ";User Id=" & strNfeT1UsuarioBd & _
                ";Password=" & s_aux
            If dbcNFe.State <> adStateClosed Then dbcNFe.Close
            dbcNFe.Open s
            
            Set cmdNFeSituacao.ActiveConnection = dbcNFe
            Set cmdNFeDanfe.ActiveConnection = dbcNFe
            
            id_nfe_emitente_anterior = id_nfe_emitente
            End If
        
        strNumeroNfNormalizado = NFeFormataNumeroNF(t_NFe_EMISSAO("NFe_numero_NF"))
        strSerieNfNormalizado = NFeFormataSerieNF(t_NFe_EMISSAO("NFe_serie_NF"))
        
    '   COMMAND PARA CHAMADA DA STORED PROCEDURE DE INTEGRA��O C/ O SISTEMA DE NFe DA TARGET ONE
        cmdNFeSituacao.Parameters("NFe") = strNumeroNfNormalizado
        cmdNFeSituacao.Parameters("Serie") = strSerieNfNormalizado
        Set rsNFeRetornoSPSituacao = cmdNFeSituacao.Execute
        intNfeRetornoSP = rsNFeRetornoSPSituacao("Retorno")
        
        If intNfeRetornoSP = 1 Then
            cmdNFeDanfe.Parameters("NFe") = strNumeroNfNormalizado
            cmdNFeDanfe.Parameters("Serie") = strSerieNfNormalizado
            Set rsNFeRetornoSPDanfe = cmdNFeDanfe.Execute
            If Not rsNFeRetornoSPDanfe.EOF Then
            '   PROGRESSO
                iperc = Int((intContadorRegistros / intQtdeTotalRegistros) * 100)
                aguarde INFO_EXECUTANDO, "copiando DANFE da NFe n� " & strNumeroNfNormalizado & "  (" & CStr(iperc) & "%)"
                
                lngFileSize = rsNFeRetornoSPDanfe("DanfePDF").ActualSize
                
                If lngFileSize <= 0 Then
                    intQtdeNFeSemDadosPdf = intQtdeNFeSemDadosPdf + 1
                    If strLogNfSemDadosPdf <> "" Then strLogNfSemDadosPdf = strLogNfSemDadosPdf & ", "
                    strLogNfSemDadosPdf = strLogNfSemDadosPdf & CStr(id_nfe_emitente) & "/" & strSerieNfNormalizado & "/" & strNumeroNfNormalizado
                    End If
                
                If lngFileSize > 0 Then
                    intQtdeArqDownload = intQtdeArqDownload + 1
                    
                '   LOG
                    If strLogNF <> "" Then strLogNF = strLogNF & ", "
                    strLogNF = strLogNF & CStr(id_nfe_emitente) & "/" & strSerieNfNormalizado & "/" & strNumeroNfNormalizado
                
                '   ARQUIVO DE DANFE
                    strNomeArqDanfe = "NFe_" & strSerieNfNormalizado & "_" & strNumeroNfNormalizado & ".pdf"
                    strDiretorioPdfDanfe = barra_invertida_add(App.Path) & "DANFE\" & strPastaEmitente
                    
                    If Not DirectoryExists(strDiretorioPdfDanfe, s_erro) Then
                        If Not ForceDirectories(strDiretorioPdfDanfe, s_erro) Then
                            s = "Falha ao tentar criar a pasta local para copiar o arquivo PDF do DANFE (" & strDiretorioPdfDanfe & "): " & s_erro
                            aviso_erro s
                            GoSub TDPDPE_FECHA_TABELAS
                            aguarde INFO_NORMAL, m_id
                            Exit Sub
                            End If
                        End If
                    
                    strNomeArqCompletoDanfe = barra_invertida_add(strDiretorioPdfDanfe) & strNomeArqDanfe
                    If FileExists(strNomeArqCompletoDanfe, s_erro) Then
                        If Not FileDelete(strNomeArqCompletoDanfe, s_erro) Then
                            s = "Falha ao tentar apagar o arquivo PDF do DANFE anterior (" & strNomeArqCompletoDanfe & "): " & s_erro
                            aviso_erro s
                            GoSub TDPDPE_FECHA_TABELAS
                            aguarde INFO_NORMAL, m_id
                            Exit Sub
                            End If
                        End If
                    
                    lFileHandle = FreeFile
                    Open strNomeArqCompletoDanfe For Binary As #lFileHandle
                    lngOffset = 0
                    Do While lngOffset < lngFileSize
                        DoEvents
                        bytFile = rsNFeRetornoSPDanfe("DanfePDF").GetChunk(CHUNK_SIZE)
                        Put #lFileHandle, , bytFile()
                        lngOffset = lngOffset + CHUNK_SIZE
                        Loop
                    
                    Close #lFileHandle
                    End If
                End If
            End If
        
    '   DESALOCA OS RECORDSETS CRIADOS PELA EXECU��O DA STORED PROCEDURE
        bd_desaloca_recordset rsNFeRetornoSPSituacao, True
        bd_desaloca_recordset rsNFeRetornoSPDanfe, True
        
        t_NFe_EMISSAO.MoveNext
        Loop


    strLog = "Download dos arquivos de DANFE das NFe's emitidas em " & Format$(dtSelecionada, FORMATO_DATA) & ": " & CStr(intQtdeArqDownload) & " arquivos copiados" & "; " & CStr(intQtdeNFeSemDadosPdf) & " NFe's n�o foram copiadas por n�o possu�rem o PDF armazenado no BD"
    If strLogNF <> "" Then
        strLog = strLog & "; Rela��o de NFe's copiadas: " & strLogNF
        End If
    
    If strLogNfSemDadosPdf <> "" Then
        strLog = strLog & "; Rela��o de NFe's N�O copiadas: " & strLogNfSemDadosPdf
        End If
    
    Call grava_log(usuario.id, "", "", "", OP_LOG_DOWNLOAD_DANFE_EM_BATCH, strLog)
    
    GoSub TDPDPE_FECHA_TABELAS

    aguarde INFO_NORMAL, m_id

    s = "Download dos arquivos de DANFE foi conclu�do com sucesso!!" & Chr(13) & _
        CStr(intQtdeArqDownload) & " arquivos foram copiados!!"
    If intQtdeNFeSemDadosPdf > 0 Then
        s = s & Chr(13) & _
            CStr(intQtdeNFeSemDadosPdf) & " NFe's n�o foram copiadas por n�o possu�rem o PDF armazenado no BD!!"
        End If
    aviso s
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TDPDPE_FECHA_TABELAS:
'==================
'   RECORDSETS
    bd_desaloca_recordset t_NFE_EMITENTE, True
    bd_desaloca_recordset t_NFe_EMISSAO, True
    bd_desaloca_recordset rsNFeRetornoSPSituacao, True
    bd_desaloca_recordset rsNFeRetornoSPDanfe, True
    
'   COMMAND
    bd_desaloca_command cmdNFeSituacao
    bd_desaloca_command cmdNFeDanfe
    
'   CONNECTION
    If Not (dbcNFe Is Nothing) Then
        If dbcNFe.State <> adStateClosed Then dbcNFe.Close
        Set dbcNFe = Nothing
        End If
        
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TDPDPE_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    GoSub TDPDPE_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub



Function calcula_BC_ICMS_ST(ByVal vl_produto As Currency, ByVal perc_MVA_ST As Double) As Currency

    calcula_BC_ICMS_ST = vl_produto + (vl_produto * perc_MVA_ST / 100)
    
End Function


Function calcula_ICMS_ST(ByVal vl_BC_ICMS_ST As Currency, ByVal perc_ICMS_ST As Double, ByVal vl_ICMS As Currency) As Currency

Dim vl_resp As Currency

    vl_resp = vl_BC_ICMS_ST * (perc_ICMS_ST / 100)
    vl_resp = vl_resp - vl_ICMS
    
    calcula_ICMS_ST = vl_resp
    
End Function



Function decodifica_NFe_inutilizacao_status(ByVal strCodStatus As String) As String
Dim strResp As String

    strCodStatus = Trim$("" & strCodStatus)
    
    strResp = ""
    If strCodStatus = "3" Then
        strResp = "Homologado"
    ElseIf strCodStatus = "1" Then
        strResp = "Em Processamento"
    ElseIf strCodStatus = "2" Then
        strResp = "Falha"
        End If
        
    decodifica_NFe_inutilizacao_status = strResp
    
End Function

Function descricao_opcao_forma_pagamento(ByVal codigo As String) As String
Dim s As String
    codigo = Trim$("" & codigo)
    Select Case codigo
        Case ID_FORMA_PAGTO_DINHEIRO
            s = "Dinheiro"
        Case ID_FORMA_PAGTO_DEPOSITO
            s = "Dep�sito"
        Case ID_FORMA_PAGTO_CHEQUE
            s = "Cheque"
        Case ID_FORMA_PAGTO_BOLETO
            s = "Boleto"
        Case ID_FORMA_PAGTO_CARTAO
            s = "Cart�o"
        Case ID_FORMA_PAGTO_BOLETO_AV
            s = "Boleto AV"
        Case ID_FORMA_PAGTO_CARTAO_MAQUINETA
            s = "Cart�o Maquineta"
        Case Else
            s = ""
        End Select
    descricao_opcao_forma_pagamento = s
End Function

Function abreviacao_opcao_forma_pagamento(ByVal codigo As String) As String
Dim s As String
    codigo = Trim$("" & codigo)
    Select Case codigo
        Case ID_FORMA_PAGTO_DINHEIRO
            s = "Dinheiro"
        Case ID_FORMA_PAGTO_DEPOSITO
            s = "Deposito"
        Case ID_FORMA_PAGTO_CHEQUE
            s = "Cheque"
        Case ID_FORMA_PAGTO_BOLETO
            s = "Boleto"
        Case ID_FORMA_PAGTO_CARTAO
            s = "Cartao"
        Case ID_FORMA_PAGTO_BOLETO_AV
            s = "Boleto AV"
        Case ID_FORMA_PAGTO_CARTAO_MAQUINETA
            s = "Cartao M"
        Case Else
            s = ""
        End Select
    abreviacao_opcao_forma_pagamento = s
End Function


Function descricao_tipo_parcelamento(ByVal codigo As String) As String
Dim s As String

    codigo = Trim$("" & codigo)
    
    Select Case codigo
        Case COD_FORMA_PAGTO_A_VISTA
            s = "� Vista"
        Case COD_FORMA_PAGTO_PARCELA_UNICA
            s = "Parcela �nica"
        Case COD_FORMA_PAGTO_PARCELADO_CARTAO
            s = "Parcelado no Cart�o"
        Case COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA
            s = "Parcelado com Entrada"
        Case COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA
            s = "Parcelado sem Entrada"
        Case COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA
            s = "Parcelado no Cart�o Maquineta"
        Case Else
            s = ""
        End Select
        
    descricao_tipo_parcelamento = s
End Function

Function descricao_finalidade_nfe(ByVal codigo As String) As String
Dim s As String

    codigo = Trim$("" & codigo)
    
    Select Case codigo
        Case NFE_FINALIDADE_NFE_NORMAL
            s = "NFe Normal"
        Case NFE_FINALIDADE_NFE_COMPLEMENTAR
            s = "NFe Complementar"
        Case NFE_FINALIDADE_NFE_AJUSTE
            s = "NFe de Ajuste"
        Case NFE_FINALIDADE_NFE_DEVOLUCAO
            s = "Devolu��o de Mercadoria"
        Case Else
            s = ""
        End Select
        
    descricao_finalidade_nfe = s
End Function

Function ibpt_aliquota_aplicavel(ByVal cst As String, ByVal percAliqNac As Single, ByVal percAliqImp As Single) As Single
Dim s_origem As String

    s_origem = Trim$(left$(cst, 1))
    
    If (s_origem = "0") Or (s_origem = "3") Or (s_origem = "4") Or (s_origem = "5") Then
        ibpt_aliquota_aplicavel = percAliqNac
    Else
        ibpt_aliquota_aplicavel = percAliqImp
        End If
    
End Function

Function is_venda_consumidor_final(ByVal cfop As String) As Boolean
Dim s_cfop As String

    is_venda_consumidor_final = False
    
    s_cfop = retorna_so_digitos(Trim("" & cfop))
    
'   A LEI 12.741/2012 EXIGE QUE A INFORMA��O DO TOTAL ESTIMADO DOS TRIBUTOS SEJA
'   INFORMADA NAS NOTAS EMITIDAS P/ CONSUMIDOR FINAL.
'   O CARLOS EM 06/06/2013 INFORMOU QUE SOMENTE CONSEGUE DISTINGUIR VENDAS P/
'   CONSUMIDOR FINAL ATRAV�S DO CFOP NOS CASOS DE VENDAS INTERESTADUAIS, POIS NAS
'   VENDAS DENTRO DO ESTADO UM MESMO CFOP � USADO P/ VENDAS P/ CONSUMIDOR FINAL E
'   P/ REVENDAS.
'   PORTANTO, POR ORA, A INFORMA��O DOS TRIBUTOS SER� EXIBIDA EM TODAS AS NOTAS.
    
    is_venda_consumidor_final = True
    
End Function

Function is_venda_interestadual_de_mercadoria_importada(ByVal cfop As String, ByVal cst As String) As Boolean

    is_venda_interestadual_de_mercadoria_importada = False
    
    cfop = Trim$("" & cfop)
    cst = Trim$("" & cst)
    
    If ((cst = "100") Or (cst = "200")) _
        And _
       ((cfop = "6102") Or (cfop = "6202") Or (cfop = "6949")) Then
        is_venda_interestadual_de_mercadoria_importada = True
        End If
    
End Function

Function existe_divergencia_loc_dest_x_cpof(ByVal cfop As String, ByVal iddest As String) As Boolean

    existe_divergencia_loc_dest_x_cpof = False

    cfop = Trim$("" & cfop)
    iddest = Trim$("" & iddest)
    
    If ((iddest <> "1") And ((left(cfop, 1) = "1") Or (left(cfop, 1) = "5"))) Or _
       ((iddest <> "2") And ((left(cfop, 1) = "2") Or (left(cfop, 1) = "6"))) Then
        existe_divergencia_loc_dest_x_cpof = True
        End If
    
End Function

Function retorna_finalidade_nfe(ByVal cod_natop As String)
Dim strTipoNfe As String

    strTipoNfe = ""
    cod_natop = retorna_so_digitos(cod_natop)
    Select Case cod_natop
        '1.201 = Devolu��o de venda de produ��o do estabelecimento
        '1.202 = Devolu��o de venda de mercadoria adquirida ou recebida de terceiros
        '1.203 = Devolu��o de venda de produ��o do estabelecimento, destinada � Zona Franca de Manaus ou ALC
        '1.204 = Devolu��o de venda de mercadoria adquirida ou recebida de terceiros, destinada � Zona Franca de Manaus ou ALC
        '1.208 = Devolu��o de produ��o do estabelecimento, remetida em transfer�ncia
        '1.209 = Devolu��o de mercadoria adquirida ou recebida de terceiros, remetida em transfer�ncia
        '1.410 = Devolu��o de venda de produ��o do estabelecimento em opera��o com produto sujeito ao regime de substitui��o tribut�ria
        '1.411 = Devolu��o de venda de mercadoria adquirida ou recebida de terceiros em opera��o com mercadoria sujeita ao regime de substitui��o tribut�ria
        '1.503 = Entrada decorrente de devolu��o de produto remetido com fim espec�fico de exporta��o, de produ��o do estabelecimento
        '1.504 = Entrada decorrente de devolu��o de mercadoria remetida com fim espec�fico de exporta��o, adquirida ou recebida de terceiros
        '1.553 = Devolu��o de venda de bem do ativo imobilizado
        '1.660 = Devolu��o de venda de combust�vel ou lubrificante destinado � industrializa��o subsequente
        '1.661 = Devolu��o de venda de combust�vel ou lubrificante destinado � comercializa��o
        '1.662 = Devolu��o de venda de combust�vel ou lubrificante destinado a consumidor ou usu�rio final
        '1.903 = Entrada de mercadoria remetida para industrializa��o e n�o aplicada no referido processo
        '1.918 = Devolu��o de mercadoria remetida em consigna��o mercantil ou industrial
        '2.201 = Devolu��o de venda de produ��o do estabelecimento
        '2.202 = Devolu��o de venda de mercadoria adquirida ou recebida de terceiros
        '2.203 = Devolu��o de venda de produ��o do estabelecimento, destinada � Zona Franca de Manaus ou ALC
        '2.204 = Devolu��o de venda de mercadoria adquirida ou recebida de terceiros, destinada � Zona Franca de Manaus ou ALC
        '2.208 = Devolu��o de produ��o do estabelecimento, remetida em transfer�ncia
        '2.209 = Devolu��o de mercadoria adquirida ou recebida de terceiros, remetida em transfer�ncia
        '2.410 = Devolu��o de venda de produ��o do estabelecimento em opera��o com produto sujeito ao regime de substitui��o tribut�ria
        '2.411 = Devolu��o de venda de mercadoria adquirida ou recebida de terceiros em opera��o com mercadoria sujeita ao regime de substitui��o tribut�ria
        '2.503 = Entrada decorrente de devolu��o de produto remetido com fim espec�fico de exporta��o, de produ��o do estabelecimento
        '2.504 = Entrada decorrente de devolu��o de mercadoria remetida com fim espec�fico de exporta��o, adquirida ou recebida de terceiros
        '2.553 = Devolu��o de venda de bem do ativo imobilizado
        '2.660 = Devolu��o de venda de combust�vel ou lubrificante destinado � industrializa��o subsequente
        '2.661 = Devolu��o de venda de combust�vel ou lubrificante destinado � comercializa��o
        '2.662 = Devolu��o de venda de combust�vel ou lubrificante destinado a consumidor ou usu�rio final
        '2.903 = Entrada de mercadoria remetida para industrializa��o e n�o aplicada no referido processo
        '2.918 = Devolu��o de mercadoria remetida em consigna��o mercantil ou industrial
        '3.201 = Devolu��o de venda de produ��o do estabelecimento
        '3.202 = Devolu��o de venda de mercadoria adquirida ou recebida de terceiros
        '3.211 = Devolu��o de venda de produ��o do estabelecimento sob o regime de "drawback"
        '3.503 = Devolu��o de mercadoria exportada que tenha sido recebida com fim espec�fico de exporta��o
        '3.553 = Devolu��o de venda de bem do ativo imobilizado
        '5.201 = Devolu��o de compra para industrializa��o ou produ��o rural
        '5.202 = Devolu��o de compra para comercializa��o
        '5.208 = Devolu��o de mercadoria recebida em transfer�ncia para industrializa��o ou produ��o rural
        '5.209 = Devolu��o de mercadoria recebida em transfer�ncia para comercializa��o
        '5.210 = Devolu��o de compra para utiliza��o na presta��o de servi�o
        '5.410 = Devolu��o de compra para industrializa��o ou produ��o rural em opera��o com mercadoria sujeita ao regime de substitui��o tribut�ria
        '5.411 = Devolu��o de compra para comercializa��o em opera��o com mercadoria sujeita ao regime de ST
        '5.412 = Devolu��o de bem do ativo imobilizado, em opera��o com mercadoria sujeita ao regime de ST
        '5.413 = Devolu��o de mercadoria destinada ao uso ou consumo, em opera��o com mercadoria sujeita ao regime de substitui��o tribut�ria
        '5.503 = Devolu��o de mercadoria recebida com fim espec�fico de exporta��o
        '5.553 = Devolu��o de compra de bem para o ativo imobilizado
        '5.555 = Devolu��o de bem do ativo imobilizado de terceiro, recebido para uso no estabelecimento
        '5.556 = Devolu��o de compra de material de uso ou consumo
        '5.660 = Devolu��o de compra de combust�vel ou lubrificante adquirido para industrializa��o subsequente
        '5.661 = Devolu��o de compra de combust�vel ou lubrificante adquirido para comercializa��o
        '5.662 = Devolu��o de compra de combust�vel ou lubrificante adquirido por consumidor ou usu�rio final
        '5.918 = Devolu��o de mercadoria recebida em consigna��o mercantil ou industrial
        '6.201 = Devolu��o de compra para industrializa��o ou produ��o rural
        '6.202 = Devolu��o de compra para comercializa��o
        '6.208 = Devolu��o de mercadoria recebida em transfer�ncia para industrializa��o ou produ��o rural
        '6.209 = Devolu��o de mercadoria recebida em transfer�ncia para comercializa��o
        '6.210 = Devolu��o de compra para utiliza��o na presta��o de servi�o
        '6.410 = Devolu��o de compra para industrializa��o ou produ��o rural em opera��o com mercadoria sujeita ao regime de substitui��o tribut�ria
        '6.411 = Devolu��o de compra para comercializa��o em opera��o com mercadoria sujeita ao regime de ST
        '6.412 = Devolu��o de bem do ativo imobilizado, em opera��o com mercadoria sujeita ao regime de ST
        '6.413 = Devolu��o de mercadoria destinada ao uso ou consumo, em opera��o com mercadoria sujeita ao regime de substitui��o tribut�ria
        '6.503 = Devolu��o de mercadoria recebida com fim espec�fico de exporta��o
        '6.553 = Devolu��o de compra de bem para o ativo imobilizado
        '6.555 = Devolu��o de bem do ativo imobilizado de terceiro, recebido para uso no estabelecimento
        '6.556 = Devolu��o de compra de material de uso ou consumo
        '6.660 = Devolu��o de compra de combust�vel ou lubrificante adquirido para industrializa��o subsequente
        '6.661 = Devolu��o de compra de combust�vel ou lubrificante adquirido para comercializa��o
        '6.662 = Devolu��o de compra de combust�vel ou lubrificante adquirido por consumidor ou usu�rio final
        '6.918 = Devolu��o de mercadoria recebida em consigna��o mercantil ou industrial
        '7.201 = Devolu��o de compra para industrializa��o ou produ��o rural
        '7.202 = Devolu��o de compra para comercializa��o
        '7.210 = Devolu��o de compra para utiliza��o na presta��o de servi�o
        '7.211 = Devolu��o de compras para industrializa��o sob o regime de "drawback"
        '7.553 = Devolu��o de compra de bem para o ativo imobilizado
        '7.556 = Devolu��o de compra de material de uso ou consumo
        Case "1201", "1202", "1203", "1204", "1208", "1209", "1410", "1411", "1503", "1504", "1553", "1660", "1661", "1662", "1903", "1918", "2201", "2202", "2203", "2204", _
             "2208", "2209", "2410", "2411", "2503", "2504", "2553", "2660", "2661", "2662", "2903", "2918", "3201", "3202", "3211", "3503", "3553", "5201", "5202", "5208", _
             "5209", "5210", "5410", "5411", "5412", "5413", "5503", "5553", "5555", "5556", "5660", "5661", "5662", "5918", "6201", "6202", "6208", "6209", "6210", "6410", _
             "6411", "6412", "6413", "6503", "6553", "6555", "6556", "6660", "6661", "6662", "6918", "7201", "7202", "7210", "7211", "7553", "7556"
            strTipoNfe = NFE_FINALIDADE_NFE_DEVOLUCAO
        Case Else
            strTipoNfe = NFE_FINALIDADE_NFE_NORMAL
        End Select
        
    retorna_finalidade_nfe = strTipoNfe
End Function

Function cfop_eh_de_remessa(ByVal cod_cfop As String) As Boolean
    
    Dim ok As Boolean
    Dim s_cfop As String
    Dim i As Integer
    
    ok = False
    s_cfop = retorna_so_digitos(cod_cfop)
    For i = LBound(vCFOPsSemPartilha) To UBound(vCFOPsSemPartilha)
        If retorna_so_digitos(vCFOPsSemPartilha(i).codigo) = s_cfop Then
            ok = True
            Exit For
            End If
        Next
    
    cfop_eh_de_remessa = ok
    
End Function

Function tem_instricao_virtual(ByVal id_emitente, s_uf As String) As Boolean
    
    Dim ok As Boolean
    Dim i As Integer
    
    ok = False
    For i = LBound(vCUFsInscricaoVirtual) To UBound(vCUFsInscricaoVirtual)
        If vCUFsInscricaoVirtual(i).c1 = id_emitente And vCUFsInscricaoVirtual(i).c2 = s_uf Then
            ok = True
            Exit For
            End If
        Next
    
    tem_instricao_virtual = ok
    
End Function

Sub limpa_item_TIPO_LINHA_NOTA_FISCAL(ByRef r As TIPO_LINHA_NOTA_FISCAL)
    
    With r
        .fabricante = ""
        .produto = ""
        .descricao = ""
        .EAN = ""
        .ncm = ""
        .NCM_bd = ""
        .NCM_tela = ""
        .cst = ""
        .CST_bd = ""
        .CST_tela = ""
        .qtde_total = 0
        .valor_total = 0
        .qtde_volumes_total = 0
        .peso_total = 0
        .cubagem_total = 0
        .perc_MVA_ST = 0
        .infAdProd = ""
        .vl_outras_despesas_acessorias = 0
        .cfop = ""
        .CFOP_formatado = ""
        .CFOP_tela = ""
        .CFOP_tela_formatado = ""
        .ICMS = ""
        .ICMS_tela = ""
        .tem_dados_IBPT = False
        .percAliqNac = 0
        .percAliqImp = 0
        .xPed = ""
        .nItemPed = ""
        .fcp = 0
        End With
        
End Sub


Sub limpa_item_TIPO_LINHA_NFe_EMISSAO_MANUAL(ByRef r As TIPO_LINHA_NFe_EMISSAO_MANUAL)
    
    With r
        .fabricante = ""
        .produto = ""
        .descricao = ""
        .descricao_bd = ""
        .descricao_tela = ""
        .EAN = ""
        .ncm = ""
        .NCM_bd = ""
        .NCM_tela = ""
        .cst = ""
        .CST_bd = ""
        .CST_tela = ""
        .qtde = 0
        .vl_unitario = 0
        .vl_outras_despesas_acessorias = 0
        .qtde_volumes_total = 0
        .peso_total = 0
        .cubagem_total = 0
        .perc_MVA_ST = 0
        .infAdProd = ""
        .cfop = ""
        .CFOP_formatado = ""
        .CFOP_tela = ""
        .CFOP_tela_formatado = ""
        .ICMS = ""
        .ICMS_tela = ""
        .tem_dados_IBPT = False
        .percAliqNac = 0
        .percAliqImp = 0
        .xPed = ""
        .nItemPed = ""
        .uCom = ""
        .uTrib = ""
        .fcp = ""
        End With
        
End Sub


Sub limpa_TIPO_NFe_IMG(ByRef r As TIPO_NFe_IMG)
Const NomeDestaRotina = "limpa_TIPO_NFe_IMG()"
Dim s As String

    On Error GoTo LTNI_TRATA_ERRO

    With r
        .id = 0
        .id_nfe_emitente = 0
        .versao_layout_NFe = ""
        .pedido = ""
        .operacional__email = ""
        .ide__natOp = ""
        .ide__indPag = ""
        .ide__serie = ""
        .ide__nNF = ""
        .ide__dEmi = ""
        .ide__dEmiUTC = ""
        .ide__dSaiEnt = ""
        .ide__tpNF = ""
        .ide__idDest = ""
        .ide__cMunFG = ""
        .ide__tpAmb = ""
        .ide__finNFe = ""
        .ide__indFinal = ""
        .ide__indPres = ""
        .ide__IEST = ""
        .dest__CNPJ = ""
        .dest__CPF = ""
        .dest__xNome = ""
        .dest__xLgr = ""
        .dest__nro = ""
        .dest__xCpl = ""
        .dest__xBairro = ""
        .dest__cMun = ""
        .dest__xMun = ""
        .dest__UF = ""
        .dest__CEP = ""
        .dest__cPais = ""
        .dest__xPais = ""
        .dest__fone = ""
        .dest__IE = ""
        .dest__ISUF = ""
        .dest__idEstrangeiro = ""
        .dest__indIEDest = ""
        .dest__email = ""
        .entrega__CNPJ = ""
        .entrega__CPF = ""
        .entrega__xLgr = ""
        .entrega__nro = ""
        .entrega__xCpl = ""
        .entrega__xBairro = ""
        .entrega__cMun = ""
        .entrega__xMun = ""
        .entrega__UF = ""
        .total__vBC = ""
        .total__vICMS = ""
        .total__vICMSDeson = ""
        .total__vBCST = ""
        .total__vST = ""
        .total__vProd = ""
        .total__vFrete = ""
        .total__vSeg = ""
        .total__vDesc = ""
        .total__vII = ""
        .total__vIPI = ""
        .total__vPIS = ""
        .total__vCOFINS = ""
        .total__vOutro = ""
        .total__vNF = ""
        .total__vTotTrib = ""
        .transp__modFrete = ""
        .transporta__CNPJ = ""
        .transporta__CPF = ""
        .transporta__xNome = ""
        .transporta__IE = ""
        .transporta__xEnder = ""
        .transporta__xMun = ""
        .transporta__UF = ""
        .vol__qVol = ""
        .vol__esp = ""
        .vol__marca = ""
        .vol__nVol = ""
        .vol__pesoL = ""
        .vol__pesoB = ""
        .vol_nLacre = ""
        .infAdic__infAdFisco = ""
        .infAdic__infCpl = ""
        .codigo_retorno_NFe_T1 = ""
        .msg_retorno_NFe_T1 = ""
        End With
        
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LTNI_TRATA_ERRO:
'===============
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub

Sub limpa_TIPO_NFe_IMG_ITEM(ByRef r() As TIPO_NFe_IMG_ITEM)
Const NomeDestaRotina = "limpa_TIPO_NFe_IMG_ITEM()"
Dim s As String
Dim ic As Integer

    On Error GoTo LTNII_TRATA_ERRO

    For ic = LBound(r) To UBound(r)
        With r(ic)
            .id = 0
            .id_nfe_imagem = 0
            .fabricante = ""
            .produto = ""
            .det__nItem = ""
            .det__cProd = ""
            .det__cEAN = ""
            .det__xProd = ""
            .det__NCM = ""
            .det__CEST = ""
            .det__indEscala = ""
            .det__EXTIPI = ""
            .det__genero = ""
            .det__CFOP = ""
            .det__uCom = ""
            .det__qCom = ""
            .det__vUnCom = ""
            .det__vProd = ""
            .det__cEANTrib = ""
            .det__uTrib = ""
            .det__qTrib = ""
            .det__vUnTrib = ""
            .det__vFrete = ""
            .det__vSeg = ""
            .det__vDesc = ""
            .ICMS__orig = ""
            .ICMS__CST = ""
            .ICMS__modBC = ""
            .ICMS__pRedBC = ""
            .ICMS__vBC = ""
            .ICMS__pICMS = ""
            .ICMS__vICMS = ""
            .ICMS__vICMSDeson = ""
            .ICMS__modBCST = ""
            .ICMS__pMVAST = ""
            .ICMS__pRedBCST = ""
            .ICMS__vBCST = ""
            .ICMS__pICMSST = ""
            .ICMS__vICMSST = ""
            .PIS__CST = ""
            .PIS__vBC = ""
            .PIS__pPIS = ""
            .PIS__vPIS = ""
            .PIS__qBCProd = ""
            .PIS__vAliqProd = ""
            .COFINS__CST = ""
            .COFINS__vBC = ""
            .COFINS__pCOFINS = ""
            .COFINS__vCOFINS = ""
            .COFINS__qBCProd = ""
            .COFINS__vAliqProd = ""
            .IPI__CST = ""
            .IPI__clEnq = ""
            .IPI__CNPJProd = ""
            .IPI__cSelo = ""
            .IPI__qSelo = ""
            .IPI__cEnq = ""
            .IPI__vBC = ""
            .IPI__qUnid = ""
            .IPI__vUnid = ""
            .IPI__pIPI = ""
            .IPI__vIPI = ""
            .det__infAdProd = ""
            .det__vOutro = ""
            .det__indTot = ""
            .det__xPed = ""
            .det__nItemPed = ""
            .det__vTotTrib = ""
            .ICMS__vBCSTRet = ""
            .ICMS__vICMSSTRet = ""
            End With
        Next

Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LTNII_TRATA_ERRO:
'================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
End Sub

Sub limpa_TIPO_NFe_IMG_NFe_REFERENCIADA(ByRef r() As TIPO_NFe_IMG_NFe_REFERENCIADA)
Const NomeDestaRotina = "limpa_TIPO_NFe_IMG_NFe_REFERENCIADA()"
Dim s As String
Dim ic As Integer

    On Error GoTo LTNINR_TRATA_ERRO

    For ic = LBound(r) To UBound(r)
        With r(ic)
            .id = 0
            .id_nfe_imagem = 0
            .refNFe = ""
            .NFe_serie_NF_referenciada = 0
            .NFe_numero_NF_referenciada = 0
            End With
        Next

Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LTNINR_TRATA_ERRO:
'=================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub

Sub limpa_TIPO_NFe_IMG_TAG_DUP(ByRef r() As TIPO_NFe_IMG_TAG_DUP)
Const NomeDestaRotina = "limpa_TIPO_NFe_IMG_TAG_DUP()"
Dim s As String
Dim ic As Integer

    On Error GoTo LTNITD_TRATA_ERRO

    For ic = LBound(r) To UBound(r)
        With r(ic)
            .id = 0
            .id_nfe_imagem = 0
            .nDup = ""
            .dVenc = ""
            .vDup = ""
            End With
        Next

Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LTNITD_TRATA_ERRO:
'=================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub

Sub limpa_TIPO_NFe_IMG_PAG(ByRef r() As TIPO_NFe_IMG_PAG)
Const NomeDestaRotina = "limpa_TIPO_NFe_IMG_PAG()"
Dim s As String
Dim ic As Integer

    On Error GoTo LTNIP_TRATA_ERRO

    For ic = LBound(r) To UBound(r)
        With r(ic)
            .id = 0
            .id_nfe_imagem = 0
            .pag__indPag = ""
            .pag__tPag = ""
            .pag__vPag = ""
            End With
        Next

Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LTNIP_TRATA_ERRO:
'=================
    s = CStr(Err) & ": " & Error$(Err) & _
        vbCrLf & _
        "Rotina: " & NomeDestaRotina
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub

End Sub

Function obtem_aliquota_ICMS_ST(ByVal uf_destino As String, ByRef aliquota_ICMS_ST As Single, ByRef msg_erro As String) As Boolean

    obtem_aliquota_ICMS_ST = False
    
    msg_erro = ""
    aliquota_ICMS_ST = 0
    
    If Trim$(uf_destino) = "" Then
        msg_erro = "Falha ao tentar obter a al�quota do ICMS ST: UF n�o foi informada!!"
        Exit Function
        End If
        
    If Not UF_ok(uf_destino) Then
        msg_erro = "Falha ao tentar obter a al�quota do ICMS ST: UF inv�lida (" & Trim$(uf_destino) & ")!!"
        Exit Function
        End If
        
    ' TODO - DETERMINAR A AL�QUOTA DO ICMS ST EM FUN��O DO ESTADO DE DESTINO
    msg_erro = "Lista de al�quota do ICMS ST n�o dispon�vel!!"
    Exit Function
    
    
    obtem_aliquota_ICMS_ST = True
    
End Function

Function obtem_aliquota_ICMS(ByVal uf_origem As String, ByVal uf_destino As String, ByRef aliquota As Single) As Boolean
Dim s As String
Dim t_FIN_ALIQUOTA_UF As ADODB.Recordset

    On Error GoTo OAI_TRATA_ERRO
    
    obtem_aliquota_ICMS = False
    
'   t_FIN_ALIQUOTA_UF
    Set t_FIN_ALIQUOTA_UF = New ADODB.Recordset
    With t_FIN_ALIQUOTA_UF
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    s = "SELECT *" & _
        " FROM t_FIN_ALIQUOTA_UF" & _
        " WHERE" & _
            " (uf_origem = '" & uf_origem & "')" & _
            " AND (uf_destino = '" & uf_destino & "')"
    t_FIN_ALIQUOTA_UF.Open s, dbc, , , adCmdText
    If t_FIN_ALIQUOTA_UF.EOF Then
        GoSub OAI_FECHA_TABELAS
        aguarde INFO_NORMAL, m_id
        Exit Function
        End If
        
    aliquota = t_FIN_ALIQUOTA_UF("aliquota_icms")
        
    obtem_aliquota_ICMS = True
        
    GoSub OAI_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    
    Exit Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OAI_FECHA_TABELAS:
'===================
'   RECORDSETS
    bd_desaloca_recordset t_FIN_ALIQUOTA_UF, True
        
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
OAI_TRATA_ERRO:
'================
    s = CStr(Err) & "Erro na obten��o da al�quota: " & Error$(Err)
    GoSub OAI_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Function

End Function

Function obtem_aliquota_ICMS_UF_destino(ByVal uf_destino As String, ByRef aliquota_ICMS_UF_destino As Single, ByRef msg_erro As String) As Boolean

    obtem_aliquota_ICMS_UF_destino = False

    msg_erro = ""
    aliquota_ICMS_UF_destino = 0
    
    If Trim$(uf_destino) = "" Then
        msg_erro = "Falha ao tentar obter a al�quota interna: UF n�o foi informada!!"
        Exit Function
        End If
        
    If Not UF_ok(uf_destino) Then
        msg_erro = "Falha ao tentar obter a al�quota interna: UF inv�lida (" & Trim$(uf_destino) & ")!!"
        Exit Function
        End If
        
    If Not obtem_aliquota_ICMS(uf_destino, uf_destino, aliquota_ICMS_UF_destino) Then
        msg_erro = "Falha ao tentar obter a al�quota interna: UF n�o identificada!!"
        End If
        
    If msg_erro <> "" Then Exit Function
    
    obtem_aliquota_ICMS_UF_destino = True

End Function

Function retorna_lista_aliquotas_ICMS() As String
Dim s_lista As String

Dim s As String
Dim t_FIN_ALIQUOTA_UF As ADODB.Recordset

    On Error GoTo RLAI_TRATA_ERRO
    
    s_lista = ""
    
'   t_FIN_ALIQUOTA_UF
    Set t_FIN_ALIQUOTA_UF = New ADODB.Recordset
    With t_FIN_ALIQUOTA_UF
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    s = "SELECT DISTINCT aliquota_icms" & _
        " FROM t_FIN_ALIQUOTA_UF" & _
        " ORDER BY aliquota_icms"
    t_FIN_ALIQUOTA_UF.Open s, dbc, , , adCmdText
    
    Do While Not t_FIN_ALIQUOTA_UF.EOF
        If s_lista <> "" Then s_lista = s_lista & vbCrLf
        s_lista = s_lista & Trim$(CStr(t_FIN_ALIQUOTA_UF("aliquota_icms")))
        t_FIN_ALIQUOTA_UF.MoveNext
        Loop
                
    GoSub RLAI_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
    
    retorna_lista_aliquotas_ICMS = s_lista
    
    Exit Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
RLAI_FECHA_TABELAS:
'===================
'   RECORDSETS
    bd_desaloca_recordset t_FIN_ALIQUOTA_UF, True
        
    Return
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
RLAI_TRATA_ERRO:
'================
    s = CStr(Err) & "Erro na obten��o da al�quota: " & Error$(Err)
    GoSub RLAI_FECHA_TABELAS
    aguarde INFO_NORMAL, m_id
'   DEIXANDO O AVISO COMENTADO, PARA USAR APENAS QUANDO FOR DEBUGAR
'    aviso_erro s
    Exit Function


End Function



Function retorna_CEST(ByVal ncm As String) As String
' O C�digo Especificador da Substitui��o Tribut�ria (CEST) deve ser informado nos casos
' de substitui��o tribut�ria em opera��es interestaduais. Est� relacionado ao NCM do produto.
'
' Como a tabela informada pela Fazenda n�o apresenta todos os NCMs presentes nos produtos da
' base de dados, estamos informando o CEST 01.036.00 (M�quinas e aparelhos de ar condicionado)
' quando o NCM n�o encontrar correspond�ncia.

    Dim s As String
    
    s = ""
      
    If (ncm = "84145910") Then
        s = "0109500"
    ElseIf (ncm = "84131900") Then
        s = "0109300"
    ElseIf (ncm = "84135090") Then
        s = "0109300"
    ElseIf (ncm = "84138100") Then
        s = "0109300"
    ElseIf (ncm = "84212100") Then
        s = "2109600"
    ElseIf (ncm = "84159090") Then
        s = "2110600"
    ElseIf (ncm = "84159020") Then
        s = "2109500"
    ElseIf (ncm = "84159010") Then
        s = "2109400"
    ElseIf (ncm = "84151090") Then
        s = "2109300"
    ElseIf (ncm = "84151019") Then
        s = "2109200"
    ElseIf (ncm = "84151011") Then
        s = "2109100"
    ElseIf (left(ncm, 6) = "841510") Or (left(ncm, 5) = "84158") Then
        s = "2109000"
    ElseIf (left(ncm, 4) = "7608") Then
        s = "1006900"
    ElseIf (ncm = "84213990") Then
        s = "0109600"
    Else
        s = "0103600"
        End If
        
    retorna_CEST = s
    
End Function

Function retorna_num_pedido_base(ByVal numeroPedido As String) As String
Dim i As Integer
Dim numeroPedidoBase As String

    numeroPedido = Trim$("" & numeroPedido)
    numeroPedido = normaliza_num_pedido(numeroPedido)
    
    i = InStr(numeroPedido, COD_SEPARADOR_FILHOTE)
    If i = 0 Then
        numeroPedidoBase = numeroPedido
    Else
        numeroPedidoBase = Mid(numeroPedido, 1, i - 1)
        End If
    
    retorna_num_pedido_base = numeroPedidoBase
End Function

Function sql_monta_criterio_texto_or(ByRef vetor() As String, ByVal parametro As String, ByVal com_aspas As Boolean) As String

Dim s As String
Dim s_resp As String
Dim iv As Integer
Dim s_aspas As String
Dim n_item As Integer


    On Error GoTo SQL_MCTOR_TRATA_ERRO
    
    sql_monta_criterio_texto_or = ""
    
    parametro = Trim$("" & parametro)
    
    s_resp = ""
    
    If com_aspas Then s_aspas = "'" Else s_aspas = ""
    
    n_item = 0
    For iv = LBound(vetor) To UBound(vetor)
        If Trim$(vetor(iv)) <> "" Then
            n_item = n_item + 1
            If s_resp <> "" Then s_resp = s_resp & " OR"
            s_resp = s_resp & " (" & parametro & "=" & s_aspas & Trim$(vetor(iv)) & s_aspas & ")"
            End If
        Next
        
    If n_item > 1 Then s_resp = " (" & s_resp & ")"
    
    sql_monta_criterio_texto_or = s_resp
        
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
SQL_MCTOR_TRATA_ERRO:
'====================
    s = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    aviso_erro s
    
    Exit Function
    
End Function

Function normaliza_lista_pedidos(ByVal lista As String) As String
'�______________________________________________________________________________________________
'|
'|  NORMALIZA A LISTA DE N�MEROS DE PEDIDOS, SEPARADOS POR "ENTER" (CR+LF)
'|

Dim v() As String
Dim v_aux() As String
Dim s As String
Dim s_lista As String
Dim i As Long
            
    s_lista = Trim("" & lista)
    s_lista = Replace$(s_lista, vbLf, "")
    v = Split(s_lista, vbCr)
    ReDim v_aux(0)
    v_aux(UBound(v_aux)) = ""
    For i = LBound(v) To UBound(v)
        If Trim$("" & v(i)) <> "" Then
            s = normaliza_num_pedido(v(i))
            If s <> "" Then
                If Trim$(v_aux(UBound(v_aux))) <> "" Then ReDim Preserve v_aux(UBound(v_aux) + 1)
                v_aux(UBound(v_aux)) = s
                End If
            End If
        Next
                
    s = Join(v_aux(), vbCrLf)
    normaliza_lista_pedidos = s
    
End Function

Function PrintText(ByVal ptext As String, ByVal originX As Integer, ByVal OriginY As Integer, ByVal OffsetX As Integer, ByVal OffsetY As Integer, ByVal fuFormat)
'�   ___________________________________________________________________________
'�  |                                                                           |
'�  |   P R I N T T E X T                                                       |
'�  |___________________________________________________________________________|
'�  |                                                                           |
'�  |                                                                           |
'�  |     ptext - contains the string to be printed                             |
'�  |     pfontname - contains the fontname to be used                          |
'�  |     pfontsize - contains the fontsize to be used                          |
'�  |                                                                           |
'�  |     originx, originy - specifies the x and y coordinates of the           |
'�  |                        upper left origin of the output area               |
'�  |                                                                           |
'�  |     offsetx, offsety - specifies the coordinates of the lower             |
'�  |                  right corner of the DrawText Rectangle                   |
'�  |                  relative to the upper left corner                        |
'�  |                                                                           |
'�  |     fuFormat  -  Accepts four values for formatting the text              |
'�  |                  within the rectangle specified by the previous           |
'�  |                  four parameters                                          |
'�  |                  0 - align left                                           |
'�  |                  1 - align center                                         |
'�  |                  2 - align right                                          |
'�  |                  16 - do word wrapping in rectangle                       |
'�  |                                                                           |
'�  |  Return value is the height of the text in current logical units          |   |
'�  |___________________________________________________________________________|

Dim OldMapMode As Integer
Dim success As Integer
Dim lpDrawTextRect As RECT


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'�  PR�LOGO
'�
    On Error GoTo handler2

    OldMapMode = SetMapMode(Printer.hdc, MM_TEXT)
    
    lpDrawTextRect.left = originX
    lpDrawTextRect.top = OriginY
    lpDrawTextRect.right = OffsetX
    lpDrawTextRect.bottom = OffsetY

  '�DT_NOPREFIX EVITA QUE O CARACTER PRECEDIDO POR & APARE�A SUBLINHADO
    success = DrawText(Printer.hdc, ptext, Len(ptext), lpDrawTextRect, (fuFormat Or DT_NOPREFIX))
    PrintText = success

  '�Reset device context to initial values:
    OldMapMode = SetMapMode(Printer.hdc, OldMapMode)
    PrintText = success

Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
handler2:
'========
    Err.Clear
    PrintText = False

    Exit Function

End Function



Function filtra_pedido(tecla As Integer) As Integer
Dim letra As String
    filtra_pedido = tecla
    letra = UCase$(Chr(tecla))
    If ((Not IsNumeric(letra)) And (Not IsLetra(letra)) And (letra <> COD_SEPARADOR_FILHOTE)) Then filtra_pedido = 0
End Function

' ------------------------------------------------------------------------
'   NORMALIZA_NUM_PEDIDO
Function normaliza_num_pedido(ByVal id_pedido As String) As String
Dim i As Integer
Dim c As String
Dim s As String
Dim s_num As String
Dim s_ano As String
Dim s_filhote As String

    normaliza_num_pedido = ""
    id_pedido = UCase(Trim("" & id_pedido))
    If id_pedido = "" Then Exit Function
    s_num = ""
    For i = 1 To Len(id_pedido)
        If IsNumeric(Mid(id_pedido, i, 1)) Then
            s_num = s_num & Mid(id_pedido, i, 1)
        Else
            Exit For
            End If
        Next
    
    If s_num = "" Then Exit Function
    
    s_ano = ""
    s_filhote = ""
    For i = 1 To Len(id_pedido)
        c = Mid(id_pedido, i, 1)
        If IsLetra(c) Then
            If s_ano = "" Then
                s_ano = c
            ElseIf s_filhote = "" Then
                s_filhote = c
                End If
            End If
        Next

    If s_ano = "" Then Exit Function
    s_num = normaliza_codigo(s_num, TAM_MIN_NUM_PEDIDO)
    s = s_num & s_ano
    If s_filhote <> "" Then s = s & COD_SEPARADOR_FILHOTE & s_filhote
    normaliza_num_pedido = s
End Function

Function configura_registry_client_sql_server(ByRef msg_erro As String) As Boolean
' ------------------------------------------------------------------------
'   CONFIGURA O REGISTRY PARA QUE O CLIENTE DO SQL SERVER FUNCIONE APENAS
'   COM OS ARQUIVOS INSTALADOS PELO MDAC.

Dim s As String
Dim s_chave As String
Dim s_campo As String
Dim s_valor As String
Dim n_valor As Long

    On Error GoTo CRCSQL_TRATA_ERRO
    
    configura_registry_client_sql_server = False
    msg_erro = ""
    
  '�DB-LIB
  '�~~~~~~
    s_chave = "Software\Microsoft\MSSQLServer\Client\DB-Lib"
    s_campo = "UseIntlSettings"
    s_valor = "off"
    If Not registry_grava_string(s_chave, s_campo, s_valor, msg_erro) Then Exit Function
    
  '�SuperSocketNetLib
  '�~~~~~~~~~~~~~~~~~
    s_chave = "Software\Microsoft\MSSQLServer\Client\SuperSocketNetLib"
    s_campo = "ProtocolOrder"
    s_valor = "7463700000"
    If Not registry_grava_binario(s_chave, s_campo, s_valor, msg_erro) Then Exit Function
    
    s_campo = "Encrypt"
    n_valor = 0
    If Not registry_grava_numero(s_chave, s_campo, n_valor, msg_erro) Then Exit Function
    
  '�T C P
  '�~~~~~
    s_chave = "Software\Microsoft\MSSQLServer\Client\SuperSocketNetLib\Tcp"
    s_campo = "DefaultPort"
    n_valor = 1433
    If Not registry_grava_numero(s_chave, s_campo, n_valor, msg_erro) Then Exit Function
    
  '�ConnectTo
  '�~~~~~~~~~
    s_chave = "SOFTWARE\Microsoft\MSSQLServer\Client\ConnectTo"
    s_campo = "DSQUERY"
    s_valor = "DBMSSOCN"
    
  '�PARA CLIENT DO SQL SERVER 2000: N�O ALTERA A CONFIGURA��O DA DLL USADA PARA TCP/IP
    Call registry_le_string(s_chave, s_campo, s)
    s = UCase$(Trim$(s))
    If s <> "DBNETLIB" Then
        If Not registry_grava_string(s_chave, s_campo, s_valor, msg_erro) Then Exit Function
        End If
    
    configura_registry_client_sql_server = True
    
Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CRCSQL_TRATA_ERRO:
'=================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    
End Function


Function le_arquivo_ini(ByRef msg_erro As String) As Boolean
' ------------------------------------------------------------------------
'   L� ARQUIVO DE CONFIGURA��O

Dim s_arq As String
Dim s_linha As String
Dim s_param As String
Dim s_valor As String
Dim s_senha As String
Dim s_senha_at As String
Dim s_senha_cep As String
Dim v() As String

'�ARQUIVO-TEXTO
Dim Fnum As Integer

    On Error GoTo LAI_TRATA_ERRO
    
    le_arquivo_ini = False
    msg_erro = ""
    
    s_arq = barra_invertida_add(App.Path) & ExtractFileName(App.EXEName, True) & ".CFG"
    If Not FileExists(s_arq, msg_erro) Then
        If msg_erro = "" Then msg_erro = "N�O foi encontrado o arquivo " & s_arq
        Exit Function
        End If

    Fnum = FreeFile
    Open s_arq For Input As Fnum
        
    On Error GoTo LAI_TRATA_ERRO_ARQUIVO
        
    s_senha = ""
    s_senha_at = ""
    Do While Not EOF(Fnum)
        
        Line Input #Fnum, s_linha
        
        If Trim$(s_linha) <> "" Then
            v = Split(s_linha, "=", -1)
            
            s_param = UCase$(Trim$(v(LBound(v))))
            
            If UBound(v) <> LBound(v) Then
                s_valor = Trim$(v(UBound(v)))
            Else
                s_valor = ""
                End If
            
            Select Case s_param
                Case "SERVER"
                    bd_selecionado.NOME_SERVIDOR = s_valor
                Case "DATABASE"
                    bd_selecionado.NOME_BD = s_valor
                Case "USER"
                    bd_selecionado.ID_USUARIO = s_valor
                Case "OPTION"
                    s_senha = s_valor
                Case "SERVER_AT"
                    bd_selecionado_at.NOME_SERVIDOR = s_valor
                Case "DATABASE_AT"
                    bd_selecionado_at.NOME_BD = s_valor
                Case "USER_AT"
                    bd_selecionado_at.ID_USUARIO = s_valor
                Case "OPTION_AT"
                    s_senha_at = s_valor
                Case "SERVER_CEP"
                    bd_selecionado_cep.NOME_SERVIDOR = s_valor
                Case "DATABASE_CEP"
                    bd_selecionado_cep.NOME_BD = s_valor
                Case "USER_CEP"
                    bd_selecionado_cep.ID_USUARIO = s_valor
                Case "OPTION_CEP"
                    s_senha_cep = s_valor
                Case "RESPTEC_CNPJ"
                    resptec_emissor.CNPJ = retorna_so_digitos(s_valor)
                Case "RESPTEC_NOME"
                    resptec_emissor.nome = s_valor
                Case "RESPTEC_EMAIL"
                    resptec_emissor.EMAIL = s_valor
                Case "RESPTEC_TELEFONE"
                    resptec_emissor.telefone = retorna_so_digitos(s_valor)
                End Select
            End If
        Loop
        
    Close Fnum
        
    'If Not decriptografa(s_senha, bd_selecionado.SENHA_USUARIO) Then
    If Not decodifica_dado(s_senha, bd_selecionado.SENHA_USUARIO) Then
        msg_erro = "Senha inv�lida !!"
        Exit Function
        End If
    
    If Not decodifica_dado(s_senha_cep, bd_selecionado_cep.SENHA_USUARIO) Then
        If Trim$(bd_selecionado_cep.NOME_BD) <> "" Then
            msg_erro = "Senha inv�lida (BD: " & bd_selecionado_cep.NOME_BD & ")!!"
        Else
            msg_erro = "Senha inv�lida para o banco de dados de CEP!!"
            End If
        Exit Function
        End If
        
    If Not decodifica_dado(s_senha_at, bd_selecionado_at.SENHA_USUARIO) Then
        'n�o faz nada
        End If
        
    le_arquivo_ini = True
    
Exit Function







'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LAI_TRATA_ERRO_ARQUIVO:
'==============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    On Error Resume Next
    Close Fnum
    
    Exit Function
    
    

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LAI_TRATA_ERRO:
'==============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    
    
End Function


Function le_arquivo_CFOP(ByRef v_CFOP() As TIPO_LISTA_CFOP, ByRef msg_erro As String) As Boolean
' ------------------------------------------------------------------------
'   L� ARQUIVO COM A LISTA DE OP��ES PARA "NATUREZA DA OPERA��O"

Dim s_arq As String
Dim s_linha As String
Dim s_codigo As String
Dim s_descricao As String
Dim v() As String

'�ARQUIVO-TEXTO
Dim Fnum As Integer

    On Error GoTo LACFOP_TRATA_ERRO
    
    le_arquivo_CFOP = False
    msg_erro = ""
    ReDim v_CFOP(0)
    v_CFOP(UBound(v_CFOP)).codigo = ""
    v_CFOP(UBound(v_CFOP)).descricao = ""
    
    s_arq = barra_invertida_add(App.Path) & "CFOP.TXT"
    If Not FileExists(s_arq, msg_erro) Then
        If msg_erro = "" Then msg_erro = "N�O foi encontrado o arquivo " & s_arq
        Exit Function
        End If

    Fnum = FreeFile
    Open s_arq For Input As Fnum
        
    On Error GoTo LACFOP_TRATA_ERRO_ARQUIVO
        
    Do While Not EOF(Fnum)
        
        Line Input #Fnum, s_linha
        
        If Trim$(s_linha) <> "" Then
            v = Split(s_linha, "=", -1)
            
            s_codigo = Trim$(v(LBound(v)))
            
            If UBound(v) <> LBound(v) Then
                s_descricao = Trim$(v(UBound(v)))
            Else
                s_descricao = ""
                End If
                        
            If (s_codigo <> "") And (s_descricao <> "") Then
                If Trim$(v_CFOP(UBound(v_CFOP)).codigo) <> "" Then
                    ReDim Preserve v_CFOP(UBound(v_CFOP) + 1)
                    End If
                
                v_CFOP(UBound(v_CFOP)).codigo = s_codigo
                v_CFOP(UBound(v_CFOP)).descricao = s_descricao
                End If
            End If
        Loop
        
    Close Fnum
        
    le_arquivo_CFOP = True
    
Exit Function







'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LACFOP_TRATA_ERRO_ARQUIVO:
'=========================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    On Error Resume Next
    Close Fnum
    
    Exit Function
    
    

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LACFOP_TRATA_ERRO:
'=================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    
    
End Function


Function le_arquivo_REMESSA_CFOP(ByRef v_REMESSA_CFOP() As TIPO_LISTA_CFOP, ByRef msg_erro As String) As Boolean
' ------------------------------------------------------------------------
'   L� ARQUIVO COM A LISTA DE CFOP's PARA OS QUAIS A INFORMA��O DE PARTILHA N�O SER� ENVIADA

Dim s_arq As String
Dim s_linha As String
Dim s_codigo As String
Dim s_descricao As String
Dim v() As String
Dim i As Integer

'�ARQUIVO-TEXTO
Dim Fnum As Integer

    On Error GoTo LARCFOP_TRATA_ERRO
    
    le_arquivo_REMESSA_CFOP = False
    msg_erro = ""
    ReDim v_REMESSA_CFOP(0)
    v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = ""
    v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = ""
    
    s_arq = barra_invertida_add(App.Path) & "REMESSA_CFOP.TXT"
    If Not FileExists(s_arq, msg_erro) Then
        
        ' SE O ARQUIVO N�O EXISTIR, TENTAR CRIAR EM TEMPO DE EXECU��O, COM OS C�DIGOS
        ' FORNECIDOS PELA SEFAZ + C�DIGOS DE ENTRADA FORNECIDOS PELO HENRIQUE (16/09/2016)

        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "1.912"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "ENTRADA DE MERCADORIA DEMONSTRA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "2.912"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "ENTRADA DE MERCADORIA DEMONSTRA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.414"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE PRODU��O DO ESTABELECIMENTO PARA VENDA FORA DO ESTABELECIMENTO EM OPERA��O COM PRODUTO SUJEITO AO REGIME DE SUBSTITUI��O TRIBUT�RIA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.415"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA ADQUIRIDA OU RECEBIDA DE TERCEIROS PARA VENDA FORA DO ESTABELECIMENTO, EM OPERA��O COM MERCADORIA SUJEITA AO REGIME DE SUBSTITUI��O TRIBUT�RIA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.451"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE ANIMAL E DE INSUMO PARA ESTABELECIMENTO PRODUTOR"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.501"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE PRODU��O DO ESTABELECIMENTO, COM FIM ESPEC�FICO DE EXPORTA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.502"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA ADQUIRIDA OU RECEBIDA DE TERCEIROS, COM FIM ESPEC�FICO DE EXPORTA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.504"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIAS PARA FORMA��O DE LOTE DE EXPORTA��O, DE PRODUTOS INDUSTRIALIZADOS OU PRODUZIDOS PELO PR�PRIO ESTABELECIMENTO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.505"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIAS, ADQUIRIDAS OU RECEBIDAS DE TERCEIROS, PARA FORMA��O DE LOTE DE EXPORTA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.554"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE BEM DO ATIVO IMOBILIZADO PARA USO FORA DO ESTABELECIMENTO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.657"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE COMBUST�VEL OU LUBRIFICANTE ADQUIRIDO OU RECEBIDO DE TERCEIROS PARA VENDA FORA DO ESTABELECIMENTO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.663"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA ARMAZENAGEM DE COMBUST�VEL OU LUBRIFICANTE"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.666"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA POR CONTA E ORDEM DE TERCEIROS DE COMBUST�VEL OU LUBRIFICANTE RECEBIDO PARA ARMAZENAGEM"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.901"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA INDUSTRIALIZA��O POR ENCOMENDA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.904"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA VENDA FORA DO ESTABELECIMENTO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.905"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA DEP�SITO FECHADO OU ARMAZ�M GERAL"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.908"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE BEM POR CONTA DE CONTRATO DE COMODATO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.910"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA EM BONIFICA��O, DOA��O OU BRINDE"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.911"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE AMOSTRA GR�TIS"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.912"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA OU BEM PARA DEMONSTRA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.914"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA OU BEM PARA EXPOSI��O OU FEIRA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.915"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA OU BEM PARA CONSERTO OU REPARO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.917"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA EM CONSIGNA��O MERCANTIL OU INDUSTRIAL"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.920"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE VASILHAME OU SACARIA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.923"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA POR CONTA E ORDEM DE TERCEIROS, EM VENDA � ORDEM OU EM OPERA��ES COM ARMAZ�M GERAL OU DEP�SITO FECHADO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.924"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA INDUSTRIALIZA��O POR CONTA E ORDEM DO ADQUIRENTE DA MERCADORIA, QUANDO ESTA N�O TRANSITAR PELO ESTABELECIMENTO DO ADQUIRENTE"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "5.934"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA SIMB�LICA DE MERCADORIA DEPOSITADA EM ARMAZ�M GERAL OU DEP�SITO FECHADO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.414"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE PRODU��O DO ESTABELECIMENTO PARA VENDA FORA DO ESTABELECIMENTO EM OPERA��O COM PRODUTO SUJEITO AO REGIME DE SUBSTITUI��O TRIBUT�RIA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.415"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA ADQUIRIDA OU RECEBIDA DE TERCEIROS PARA VENDA FORA DO ESTABELECIMENTO, EM OPERA��O COM MERCADORIA SUJEITA AO REGIME DE SUBSTITUI��O TRIBUT�RIA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.501"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE PRODU��O DO ESTABELECIMENTO, COM FIM ESPEC�FICO DE EXPORTA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.502"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA ADQUIRIDA OU RECEBIDA DE TERCEIROS, COM FIM ESPEC�FICO DE EXPORTA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.504"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIAS PARA FORMA��O DE LOTE DE EXPORTA��O, DE PRODUTOS INDUSTRIALIZADOS OU PRODUZIDOS PELO PR�PRIO ESTABELECIMENTO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.505"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIAS, ADQUIRIDAS OU RECEBIDAS DE TERCEIROS, PARA FORMA��O DE LOTE DE EXPORTA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.554"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE BEM DO ATIVO IMOBILIZADO PARA USO FORA DO ESTABELECIMENTO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.657"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE COMBUST�VEL OU LUBRIFICANTE ADQUIRIDO OU RECEBIDO DE TERCEIROS PARA VENDA FORA DO ESTABELECIMENTO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.663"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA ARMAZENAGEM DE COMBUST�VEL OU LUBRIFICANTE"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.666"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA POR CONTA E ORDEM DE TERCEIROS DE COMBUST�VEL OU LUBRIFICANTE RECEBIDO PARA ARMAZENAGEM"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.901"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA INDUSTRIALIZA��O POR ENCOMENDA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.904"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA VENDA FORA DO ESTABELECIMENTO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.905"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA DEP�SITO FECHADO OU ARMAZ�M GERAL"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.908"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE BEM POR CONTA DE CONTRATO DE COMODATO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.910"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA EM BONIFICA��O, DOA��O OU BRINDE"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.911"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE AMOSTRA GR�TIS"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.912"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA OU BEM PARA DEMONSTRA��O"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.914"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA OU BEM PARA EXPOSI��O OU FEIRA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.915"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA OU BEM PARA CONSERTO OU REPARO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.917"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA EM CONSIGNA��O MERCANTIL OU INDUSTRIAL"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.920"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE VASILHAME OU SACARIA"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.923"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA DE MERCADORIA POR CONTA E ORDEM DE TERCEIROS, EM VENDA � ORDEM OU EM OPERA��ES COM ARMAZ�M GERAL OU DEP�SITO FECHADO"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.924"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA PARA INDUSTRIALIZA��O POR CONTA E ORDEM DO ADQUIRENTE DA MERCADORIA, QUANDO ESTA N�O TRANSITAR PELO ESTABELECIMENTO DO ADQUIRENTE"
        
        If v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo <> "" Then ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = "6.934"
        v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = "REMESSA SIMB�LICA DE MERCADORIA DEPOSITADA EM ARMAZ�M GERAL OU DEP�SITO FECHADO"
        
        Fnum = FreeFile
        Open s_arq For Output As Fnum
            
        On Error GoTo LARCFOP_TRATA_ERRO_CRIA_ARQUIVO
        
        For i = LBound(v_REMESSA_CFOP) To UBound(v_REMESSA_CFOP)
            s_linha = v_REMESSA_CFOP(i).codigo & " = " & v_REMESSA_CFOP(i).descricao
            Print #Fnum, s_linha
            Next
        
        Close #Fnum
        
    Else
        Fnum = FreeFile
        Open s_arq For Input As Fnum
            
        On Error GoTo LARCFOP_TRATA_ERRO_ARQUIVO
            
        Do While Not EOF(Fnum)
            
            Line Input #Fnum, s_linha
            
            If Trim$(s_linha) <> "" Then
                v = Split(s_linha, "=", -1)
                
                s_codigo = Trim$(v(LBound(v)))
                
                If UBound(v) <> LBound(v) Then
                    s_descricao = Trim$(v(UBound(v)))
                Else
                    s_descricao = ""
                    End If
                            
                If (s_codigo <> "") And (s_descricao <> "") Then
                    If Trim$(v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo) <> "" Then
                        ReDim Preserve v_REMESSA_CFOP(UBound(v_REMESSA_CFOP) + 1)
                        End If
                    
                    v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).codigo = s_codigo
                    v_REMESSA_CFOP(UBound(v_REMESSA_CFOP)).descricao = s_descricao
                    End If
                End If
            Loop
            
        Close Fnum
        
        End If
        
    le_arquivo_REMESSA_CFOP = True
    
Exit Function





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LARCFOP_TRATA_ERRO_ARQUIVO:
'=========================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    On Error Resume Next
    Close Fnum
    
    Exit Function
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LARCFOP_TRATA_ERRO_CRIA_ARQUIVO:
'===============================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    On Error Resume Next
    Close Fnum
    
    Return

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LARCFOP_TRATA_ERRO:
'=================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    Exit Function
    
    
End Function

Function le_UFs_INSCRICAO_VIRTUAL(ByRef v_ID_UF() As TIPO_DUAS_COLUNAS, ByRef msg_erro As String) As Boolean
' ------------------------------------------------------------------------
'   L� TABELA COM A LISTA DE UFs PARA OS QUAIS N�O SER� EMITIDA A INFORMA��O SOBRE PARTILHA DO ICMS

Dim t_INSCRICAO_VIRTUAL_EMITENTE As ADODB.Recordset
Dim s As String

    On Error GoTo LUIV_TRATA_ERRO
    
'   t_INSCRICAO_VIRTUAL_EMITENTE
    Set t_INSCRICAO_VIRTUAL_EMITENTE = New ADODB.Recordset
    With t_INSCRICAO_VIRTUAL_EMITENTE
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With

    
    
    le_UFs_INSCRICAO_VIRTUAL = False
    msg_erro = ""
    ReDim v_ID_UF(0)
    v_ID_UF(UBound(v_ID_UF)).c1 = 0
    v_ID_UF(UBound(v_ID_UF)).c2 = ""
    
    On Error GoTo LUIV_TRATA_ERRO
        
    s = "SELECT *" & _
        " FROM t_INSCRICAO_VIRTUAL_EMITENTE"
    If t_INSCRICAO_VIRTUAL_EMITENTE.State <> adStateClosed Then t_INSCRICAO_VIRTUAL_EMITENTE.Close
    t_INSCRICAO_VIRTUAL_EMITENTE.Open s, dbc, , , adCmdText
    
    Do While Not t_INSCRICAO_VIRTUAL_EMITENTE.EOF
        
        If Trim$(v_ID_UF(UBound(v_ID_UF)).c1) <> 0 Then
            ReDim Preserve v_ID_UF(UBound(v_ID_UF) + 1)
            End If
        v_ID_UF(UBound(v_ID_UF)).c1 = t_INSCRICAO_VIRTUAL_EMITENTE("id_nfe_emitente")
        v_ID_UF(UBound(v_ID_UF)).c2 = Trim("" & t_INSCRICAO_VIRTUAL_EMITENTE("uf"))
        
        t_INSCRICAO_VIRTUAL_EMITENTE.MoveNext
        
        Loop
        
    GoSub LUIV_FECHA_TABELAS
    
    le_UFs_INSCRICAO_VIRTUAL = True
    
    
    Exit Function
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LUIV_FECHA_TABELAS:
'===================
'   RECORDSETS
    bd_desaloca_recordset t_INSCRICAO_VIRTUAL_EMITENTE, True
        
    Return


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LUIV_TRATA_ERRO:
'=================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    
    GoSub LUIV_FECHA_TABELAS
    
    Exit Function
    
    
End Function


Function normaliza_codigo(ByVal codigo As String, ByVal tamanho_default As Long) As String
' ------------------------------------------------------------------------
'   NORMALIZA_CODIGO
Dim s As String
    normaliza_codigo = ""
    s = Trim$("" & codigo)
    If s = "" Then Exit Function
    Do While Len(s) < tamanho_default: s = "0" & s: Loop
    normaliza_codigo = s
End Function

Function filtra_numerico(tecla As Integer) As Integer
'�_____________________________________
'|                                     |
'|  PERMITE SOMENTE C�DIGOS NUM�RICOS  |
'|_____________________________________|

    filtra_numerico = tecla
 
  '�PERMITE A PASSAGEM DOS C�DIGOS DE CONTROLE
    If tecla < Asc(" ") Then Exit Function
    
  '�FILTRA C�DIGOS DIFERENTES DE ['0'..'9']
    If (tecla < Asc("0") Or tecla > Asc("9")) Then
        filtra_numerico = 0
        Beep
        End If

End Function

Function filtra_letra(tecla As Integer) As Integer
'�_____________________________________
'|                                     |
'|  PERMITE SOMENTE LETRAS             |
'|_____________________________________|

    filtra_letra = tecla
 
  '�PERMITE A PASSAGEM DOS C�DIGOS DE CONTROLE
    If tecla < Asc(" ") Then Exit Function
    
  '�FILTRA C�DIGOS DIFERENTES DE ['A'..'Z']
    If (Asc(UCase(Chr(tecla))) < Asc("A") Or Asc(UCase(Chr(tecla))) > Asc("Z")) Then
        filtra_letra = 0
        Beep
        End If

End Function


Function configura_registry_usuario_horario_verao(hv_ini As Integer, hv_fim As Integer) As Boolean
' ------------------------------------------------------------------------
'   CONFIGURA O REGISTRY PARA QUE O PROGRAMA MEMORIZE SE A OP��O
'   "HOR�RIO DE VER�O" EST� MARCADA OU N�O.
'   TAMB�M SER� MEMORIZADA A DATA EM QUE A MEMORIZA��O FOI GRAVADA,
'   PARA PERGUNTAR PERIODICAMENTE SE A MARCA��O DEVE SER MANTIDA.

Dim s As String
Dim s_chave As String
Dim s_campo As String
Dim s_valor As String
Dim n_valor As Long

    configura_registry_usuario_horario_verao = False
    
    If (hv_ini <> hv_fim) Then
        s_chave = REG_CHAVE_USUARIO_HORARIO_VERAO
        
        s_campo = "HVAtivo"
        n_valor = hv_fim
        If Not registry_usuario_grava_numero(s_chave, s_campo, n_valor) Then Exit Function
        
        s_campo = "HVData"
        s_valor = CStr(Now)
        If Not registry_usuario_grava_string(s_chave, s_campo, s_valor) Then Exit Function

        End If
        
    configura_registry_usuario_horario_verao = True
    
End Function

Function le_registry_usuario_horario_verao(hv_valor As Integer, hv_data As String) As Boolean
' ------------------------------------------------------------------------
'   LE O REGISTRY PARA QUE O PROGRAMA VERIFIQUE SE A OP��O
'   "HOR�RIO DE VER�O" EST� MARCADA OU N�O.
'   TAMB�M SER� LIDA A DATA EM QUE A MEMORIZA��O FOI GRAVADA,
'   PARA PERGUNTAR PERIODICAMENTE SE A MARCA��O DEVE SER MANTIDA.

Dim s As String
Dim s_chave As String
Dim s_campo As String
Dim s_valor As String
Dim n_valor As Long

    le_registry_usuario_horario_verao = False
    
    s_chave = REG_CHAVE_USUARIO_HORARIO_VERAO
    
    s_campo = "HVAtivo"
    n_valor = 0
    If Not registry_usuario_le_numero(s_chave, s_campo, n_valor) Then Exit Function
    hv_valor = n_valor
    
    s_campo = "HVData"
    s_valor = ""
    If Not registry_usuario_le_string(s_chave, s_campo, s_valor) Then Exit Function
    hv_data = s_valor
        
    le_registry_usuario_horario_verao = True
    
End Function

Function calculaDataPrimeiroBoleto(ByVal intPrazoEmissaoPrimeiroBoleto As Integer) As Date

Dim dtResposta As Date

    If intPrazoEmissaoPrimeiroBoleto <= 29 Then
        dtResposta = Date + 30
    Else
        'dtResposta = Date + intPrazoEmissaoPrimeiroBoleto + 7
        'REMO��O DOS 07 DIAS ADICIONAIS, A PEDIDO DO CARLOS
        dtResposta = Date + intPrazoEmissaoPrimeiroBoleto
        End If

    calculaDataPrimeiroBoleto = dtResposta
    
End Function


Function configura_registry_usuario_info_parcelas(parc_ini As Integer, parc_fim As Integer) As Boolean
' ------------------------------------------------------------------------
'   CONFIGURA O REGISTRY PARA QUE O PROGRAMA MEMORIZE SE A OP��O
'   SOBRE A INCLUS�O DOS DADOS DE PARCELAS NO CAMPO 'INFORMA��ES ADICIONAIS'
'   EST� MARCADA OU N�O.

Dim s As String
Dim s_chave As String
Dim s_campo As String
Dim s_valor As String
Dim n_valor As Long

    configura_registry_usuario_info_parcelas = False
    
    If (parc_ini <> parc_fim) Then
        s_chave = REG_CHAVE_USUARIO_PARCELAS_INFO
        
        s_campo = "InfoParcAtivo"
        n_valor = parc_fim
        If Not registry_usuario_grava_numero(s_chave, s_campo, n_valor) Then Exit Function
        
        End If
        
    configura_registry_usuario_info_parcelas = True
    
End Function

Function le_registry_usuario_info_parcelas(parc_valor As Integer) As Boolean
' ------------------------------------------------------------------------
'   LE O REGISTRY PARA QUE O PROGRAMA VERIFIQUE SE A OP��O
'   SOBRE A INCLUS�O DOS DADOS DE PARCELAS NO CAMPO 'INFORMA��ES ADICIONAIS'
'   EST� MARCADA OU N�O.

Dim s As String
Dim s_chave As String
Dim s_campo As String
Dim s_valor As String
Dim n_valor As Long

    le_registry_usuario_info_parcelas = False
    
    s_chave = REG_CHAVE_USUARIO_PARCELAS_INFO
    
    s_campo = "InfoParcAtivo"
    n_valor = 0
    If Not registry_usuario_le_numero(s_chave, s_campo, n_valor) Then Exit Function
    parc_valor = n_valor
    
    le_registry_usuario_info_parcelas = True
    
End Function

Function geraDadosParcelasPagto(v_pedido() As String, v_parcela_pagto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO, ByRef strMsgErro As String) As Boolean
'�__________________________________________________________________________________________
'|
'|  ANALISA O(S) PEDIDO(S) PARA VERIFICAR SE H� ALGUM QUE ESPECIFICA PAGAMENTO VIA BOLETO.
'|  EM CASO AFIRMATIVO, CALCULA A QUANTIDADE DE PARCELAS, DATAS E VALORES.
'|

Dim s As String
Dim s_where As String
Dim i As Integer
Dim j As Integer
Dim intQtdeTotalPedidos As Integer
Dim intQtdePedidosPagtoBoleto As Integer
Dim intQtdeTotalParcelas As Integer
Dim intQtdePlanoContas As Integer
Dim vlTotalPedido As Currency
Dim vlTotalFormaPagto As Currency
Dim vlDiferencaArredondamento As Currency
Dim vlDiferencaArredondamentoRestante As Currency
Dim vlRateio As Currency
Dim dtUltimoPagtoCalculado As Date
Dim blnPagtoPorBoleto As Boolean
Dim strTipoParcelamento As String
Dim strListaPedidosPagtoBoleto As String
Dim strListaPedidosPagtoNaoBoleto As String
Dim vPedidoCalculoParcelas() As TIPO_PEDIDO_CALCULO_PARCELAS_BOLETO

'�BANCO DE DADOS
Dim t_PEDIDO As ADODB.Recordset
Dim t_PEDIDO_ITEM As ADODB.Recordset
Dim tAux As ADODB.Recordset

    On Error GoTo GDPP_TRATA_ERRO

    geraDadosParcelasPagto = False
    
    strMsgErro = ""
    ReDim v_parcela_pagto(0)
    
    ReDim vPedidoCalculoParcelas(0)
    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pedido = ""
  
  '�T_PEDIDO
    Set t_PEDIDO = New ADODB.Recordset
    With t_PEDIDO
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
  '�T_PEDIDO_ITEM
    Set t_PEDIDO_ITEM = New ADODB.Recordset
    With t_PEDIDO_ITEM
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
  '�tAux
    Set tAux = New ADODB.Recordset
    With tAux
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    For i = LBound(v_pedido) To UBound(v_pedido)
        If Trim$(v_pedido(i)) <> "" Then
            s = "SELECT" & _
                    " t_PEDIDO__BASE.tipo_parcelamento," & _
                    " t_PEDIDO__BASE.av_forma_pagto," & _
                    " t_PEDIDO__BASE.pc_qtde_parcelas," & _
                    " t_PEDIDO__BASE.pc_valor_parcela," & _
                    " t_PEDIDO__BASE.pce_forma_pagto_entrada," & _
                    " t_PEDIDO__BASE.pce_forma_pagto_prestacao," & _
                    " t_PEDIDO__BASE.pce_entrada_valor," & _
                    " t_PEDIDO__BASE.pce_prestacao_qtde," & _
                    " t_PEDIDO__BASE.pce_prestacao_valor," & _
                    " t_PEDIDO__BASE.pce_prestacao_periodo," & _
                    " t_PEDIDO__BASE.pse_forma_pagto_prim_prest," & _
                    " t_PEDIDO__BASE.pse_forma_pagto_demais_prest," & _
                    " t_PEDIDO__BASE.pse_prim_prest_valor," & _
                    " t_PEDIDO__BASE.pse_prim_prest_apos," & _
                    " t_PEDIDO__BASE.pse_demais_prest_qtde," & _
                    " t_PEDIDO__BASE.pse_demais_prest_valor," & _
                    " t_PEDIDO__BASE.pse_demais_prest_periodo," & _
                    " t_PEDIDO__BASE.pu_forma_pagto," & _
                    " t_PEDIDO__BASE.pu_valor," & _
                    " t_PEDIDO__BASE.pu_vencto_apos" & _
                " FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
                    " ON (SUBSTRING(t_PEDIDO.pedido,1," & CStr(TAM_MIN_ID_PEDIDO) & ")=t_PEDIDO__BASE.pedido)" & _
                " WHERE" & _
                    " (t_PEDIDO.pedido='" & Trim$(v_pedido(i)) & "')"
            If t_PEDIDO.State <> adStateClosed Then t_PEDIDO.Close
            t_PEDIDO.Open s, dbc, , , adCmdText
            If t_PEDIDO.EOF Then
                If strMsgErro <> "" Then strMsgErro = strMsgErro & vbCrLf
                strMsgErro = strMsgErro & "Pedido " & Trim$(v_pedido(i)) & " n�o est� cadastrado!!"
            Else
                intQtdeTotalPedidos = intQtdeTotalPedidos + 1
                
                strTipoParcelamento = Trim$("" & t_PEDIDO("tipo_parcelamento"))
                blnPagtoPorBoleto = False
                If strTipoParcelamento = CStr(COD_FORMA_PAGTO_A_VISTA) Then
                    If Trim$("" & t_PEDIDO("av_forma_pagto")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
                    If Trim$("" & t_PEDIDO("pce_forma_pagto_entrada")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
                    If Trim$("" & t_PEDIDO("pce_forma_pagto_prestacao")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
                    If Trim$("" & t_PEDIDO("pse_forma_pagto_prim_prest")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
                    If Trim$("" & t_PEDIDO("pse_forma_pagto_demais_prest")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
                ElseIf strTipoParcelamento = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
                    If Trim$("" & t_PEDIDO("pu_forma_pagto")) = CStr(ID_FORMA_PAGTO_BOLETO) Then blnPagtoPorBoleto = True
                    End If
                    
                If (Trim$(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pedido) <> "") Then
                    ReDim Preserve vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas) + 1)
                    End If
            
                With vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas))
                    .pedido = Trim$(v_pedido(i))
                    .tipo_parcelamento = t_PEDIDO("tipo_parcelamento")
                    .av_forma_pagto = t_PEDIDO("av_forma_pagto")
                    .pu_forma_pagto = t_PEDIDO("pu_forma_pagto")
                    .pu_valor = t_PEDIDO("pu_valor")
                    .pu_vencto_apos = t_PEDIDO("pu_vencto_apos")
                    .pc_qtde_parcelas = t_PEDIDO("pc_qtde_parcelas")
                    .pc_valor_parcela = t_PEDIDO("pc_valor_parcela")
                    .pce_forma_pagto_entrada = t_PEDIDO("pce_forma_pagto_entrada")
                    .pce_forma_pagto_prestacao = t_PEDIDO("pce_forma_pagto_prestacao")
                    .pce_entrada_valor = t_PEDIDO("pce_entrada_valor")
                    .pce_prestacao_qtde = t_PEDIDO("pce_prestacao_qtde")
                    .pce_prestacao_valor = t_PEDIDO("pce_prestacao_valor")
                    .pce_prestacao_periodo = t_PEDIDO("pce_prestacao_periodo")
                    .pse_forma_pagto_prim_prest = t_PEDIDO("pse_forma_pagto_prim_prest")
                    .pse_forma_pagto_demais_prest = t_PEDIDO("pse_forma_pagto_demais_prest")
                    .pse_prim_prest_valor = t_PEDIDO("pse_prim_prest_valor")
                    .pse_prim_prest_apos = t_PEDIDO("pse_prim_prest_apos")
                    .pse_demais_prest_qtde = t_PEDIDO("pse_demais_prest_qtde")
                    .pse_demais_prest_valor = t_PEDIDO("pse_demais_prest_valor")
                    .pse_demais_prest_periodo = t_PEDIDO("pse_demais_prest_periodo")
                    End With
                    
            '   CALCULA O VALOR TOTAL DESTE PEDIDO
                s = "SELECT" & _
                        " p.pedido," & _
                        " Coalesce(Sum(qtde*preco_NF),0) AS vl_total" & _
                    " FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM i ON (p.pedido=i.pedido)" & _
                    " WHERE" & _
                        " (p.pedido = '" & Trim$(v_pedido(i)) & "')" & _
                    " GROUP BY" & _
                        " p.pedido" & _
                    " UNION " & _
                    " SELECT" & _
                        " p.pedido," & _
                        " -1*Coalesce(Sum(qtde*preco_NF),0) AS vl_total" & _
                    " FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO id ON (p.pedido=id.pedido)" & _
                    " WHERE" & _
                        " (p.pedido = '" & Trim$(v_pedido(i)) & "')" & _
                    " GROUP BY" & _
                        " p.pedido"

                s = "SELECT" & _
                        " pedido," & _
                        " Sum(vl_total) AS vl_total" & _
                    " FROM" & _
                        "(" & _
                            s & _
                        ") t" & _
                    " GROUP BY" & _
                        " pedido"
                
                If tAux.State <> adStateClosed Then tAux.Close
                tAux.Open s, dbc, , , adCmdText
                If tAux.EOF Then
                    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).vlTotalDestePedido = 0
                Else
                    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).vlTotalDestePedido = tAux("vl_total")
                    End If
                
            '   CALCULA O VALOR TOTAL DA FAM�LIA DE PEDIDOS
                s = "SELECT" & _
                        " Coalesce(Sum(qtde*preco_NF),0) AS vl_total" & _
                    " FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM i ON (p.pedido=i.pedido)" & _
                    " WHERE" & _
                        " (p.pedido LIKE '" & retorna_num_pedido_base(Trim$(v_pedido(i))) & BD_CURINGA_TODOS & "')" & _
                        " AND (st_entrega <> '" & ST_ENTREGA_CANCELADO & "')" & _
                    " UNION " & _
                    " SELECT" & _
                        " -1*Coalesce(Sum(qtde*preco_NF),0) AS vl_total" & _
                    " FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO id ON (p.pedido=id.pedido)" & _
                    " WHERE" & _
                        " (p.pedido LIKE '" & retorna_num_pedido_base(Trim$(v_pedido(i))) & BD_CURINGA_TODOS & "')"

                s = "SELECT" & _
                        " Sum(vl_total) AS vl_total" & _
                    " FROM" & _
                        "(" & _
                            s & _
                        ") t"
                
                If tAux.State <> adStateClosed Then tAux.Close
                tAux.Open s, dbc, , , adCmdText
                If tAux.EOF Then
                    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).vlTotalFamiliaPedidos = 0
                Else
                    vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).vlTotalFamiliaPedidos = tAux("vl_total")
                    End If
                
            '   CALCULA A RAZ�O ENTRE OS VALORES DESTE PEDIDO E A FAM�LIA DE PEDIDOS
                With vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas))
                    If .vlTotalFamiliaPedidos = 0 Then
                        .razaoValorPedidoFilhote = 0
                    Else
                        .razaoValorPedidoFilhote = .vlTotalDestePedido / .vlTotalFamiliaPedidos
                        End If
                    End With
                
                If blnPagtoPorBoleto Then
                    intQtdePedidosPagtoBoleto = intQtdePedidosPagtoBoleto + 1
                    If strListaPedidosPagtoBoleto <> "" Then strListaPedidosPagtoBoleto = strListaPedidosPagtoBoleto & ", "
                    strListaPedidosPagtoBoleto = strListaPedidosPagtoBoleto & Trim$(v_pedido(i))
                Else
                    If strListaPedidosPagtoNaoBoleto <> "" Then strListaPedidosPagtoNaoBoleto = strListaPedidosPagtoNaoBoleto & ", "
                    strListaPedidosPagtoNaoBoleto = strListaPedidosPagtoNaoBoleto & Trim$(v_pedido(i))
                    End If
                End If
            End If
        Next
    
  
  
'   SE HOUVER ALGUM PEDIDO QUE DEFINA PAGAMENTO POR BOLETO, OS DADOS DE PAGAMENTO SER�O IMPRESSOS NA NF.
'   ENTRETANTO, QUANDO H� MAIS DE 2 PEDIDOS, A FORMA DE PAGAMENTO DEVE SER ID�NTICA P/ QUE SE POSSA SOMAR
'   OS VALORES DE CADA PARCELA, CASO CONTR�RIO SER� RETORNADA UMA MENSAGEM DE ERRO PARA EXIBI��O.

'�  N�O H� PEDIDOS POR BOLETOS!
    If intQtdePedidosPagtoBoleto = 0 Then
        geraDadosParcelasPagto = True
        GoSub GDPP_FECHA_TABELAS
        Exit Function
        End If
    
    
  '�H� PEDIDOS QUE S�O POR BOLETO E OUTROS QUE N�O
    If intQtdePedidosPagtoBoleto <> intQtdeTotalPedidos Then
        strMsgErro = "H� pedido(s) que especifica(m) pagamento via boleto banc�rio e h� pedido(s) que especifica(m) outro(s) meio(s) de pagamento:" & Chr(13) & _
                     "Pagamento via boleto banc�rio: " & strListaPedidosPagtoBoleto & Chr(13) & _
                     "Pagamento via outros meios: " & strListaPedidosPagtoNaoBoleto & Chr(13) & _
                     Chr(13) & _
                     "N�o � poss�vel gerar os dados de pagamento para impress�o na NFe!!"
        GoSub GDPP_FECHA_TABELAS
        Exit Function
        End If
    
    
  '�H� MAIS DO QUE 1 PEDIDO P/ SER PAGO POR BOLETO
    If intQtdePedidosPagtoBoleto > 1 Then
      '�H� PEDIDOS QUE ESPECIFICAM DIFERENTES FORMAS DE PAGAMENTO
        For i = LBound(vPedidoCalculoParcelas) To (UBound(vPedidoCalculoParcelas) - 1)
            If vPedidoCalculoParcelas(i).tipo_parcelamento <> vPedidoCalculoParcelas(i + 1).tipo_parcelamento Then
                If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                strMsgErro = strMsgErro & "Pedido " & vPedidoCalculoParcelas(i).pedido & "=" & descricao_tipo_parcelamento(vPedidoCalculoParcelas(i).tipo_parcelamento) & _
                             " e pedido " & vPedidoCalculoParcelas(i + 1).pedido & "=" & descricao_tipo_parcelamento(vPedidoCalculoParcelas(i + 1).tipo_parcelamento)
                End If
            Next
            
        If strMsgErro <> "" Then
            strMsgErro = "Os pedidos especificam diferentes formas de pagamento!!" & _
                        Chr(13) & _
                        strMsgErro & _
                        Chr(13) & _
                        Chr(13) & _
                        "N�o � poss�vel gerar os dados de pagamento para impress�o na NFe!!"
            GoSub GDPP_FECHA_TABELAS
            Exit Function
            End If
        
      '�H� PEDIDOS QUE P/ UMA FORMA DE PAGAMENTO DEFINEM DIFERENTES PRAZOS DE PAGAMENTO
        For i = LBound(vPedidoCalculoParcelas) To (UBound(vPedidoCalculoParcelas) - 1)
        '   PARCELADO COM ENTRADA
        '   ~~~~~~~~~~~~~~~~~~~~~
            If CStr(vPedidoCalculoParcelas(i).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
                If vPedidoCalculoParcelas(i).pce_forma_pagto_entrada <> vPedidoCalculoParcelas(i + 1).pce_forma_pagto_entrada Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia na forma de pagamento da entrada: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i).pce_forma_pagto_entrada) & ") e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i + 1).pce_forma_pagto_entrada) & ")"
                    End If
                
                If vPedidoCalculoParcelas(i).pce_forma_pagto_prestacao <> vPedidoCalculoParcelas(i + 1).pce_forma_pagto_prestacao Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia na forma de pagamento das presta��es: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i).pce_forma_pagto_prestacao) & ") e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i + 1).pce_forma_pagto_prestacao) & ")"
                    End If
                    
                If vPedidoCalculoParcelas(i).pce_prestacao_qtde <> vPedidoCalculoParcelas(i + 1).pce_prestacao_qtde Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia na quantidade de presta��es: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pce_prestacao_qtde) & " " & IIf(vPedidoCalculoParcelas(i).pce_prestacao_qtde > 1, "presta��es", "presta��o") & ") e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pce_prestacao_qtde) & " " & IIf(vPedidoCalculoParcelas(i + 1).pce_prestacao_qtde > 1, "presta��es", "presta��o") & ")"
                    End If
                    
                If vPedidoCalculoParcelas(i).pce_prestacao_periodo <> vPedidoCalculoParcelas(i + 1).pce_prestacao_periodo Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia no per�odo de vencimento das presta��es: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pce_prestacao_periodo) & " dias) e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pce_prestacao_periodo) & " dias)"
                    End If
                    
        '   PARCELADO SEM ENTRADA
        '   ~~~~~~~~~~~~~~~~~~~~~
            ElseIf CStr(vPedidoCalculoParcelas(i).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
                If vPedidoCalculoParcelas(i).pse_forma_pagto_prim_prest <> vPedidoCalculoParcelas(i + 1).pse_forma_pagto_prim_prest Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia na forma de pagamento da 1� presta��o: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i).pse_forma_pagto_prim_prest) & ") e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i + 1).pse_forma_pagto_prim_prest) & ")"
                    End If
                    
                If vPedidoCalculoParcelas(i).pse_forma_pagto_demais_prest <> vPedidoCalculoParcelas(i + 1).pse_forma_pagto_demais_prest Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia na forma de pagamento das demais presta��es: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i).pse_forma_pagto_demais_prest) & ") e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & descricao_opcao_forma_pagamento(vPedidoCalculoParcelas(i + 1).pse_forma_pagto_demais_prest) & ")"
                    End If
                    
                If vPedidoCalculoParcelas(i).pse_prim_prest_apos <> vPedidoCalculoParcelas(i + 1).pse_prim_prest_apos Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia no prazo de pagamento da 1� presta��o: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pse_prim_prest_apos) & ") e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pse_prim_prest_apos) & ")"
                    End If
                    
                If vPedidoCalculoParcelas(i).pse_demais_prest_qtde <> vPedidoCalculoParcelas(i + 1).pse_demais_prest_qtde Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia na quantidade de presta��es: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pse_demais_prest_qtde) & ") e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pse_demais_prest_qtde) & ")"
                    End If
                    
                If vPedidoCalculoParcelas(i).pse_demais_prest_periodo <> vPedidoCalculoParcelas(i + 1).pse_demais_prest_periodo Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia no per�odo de vencimento das presta��es: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pse_demais_prest_periodo) & " dias) e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pse_demais_prest_periodo) & " dias)"
                    End If
            
        '   PARCELA �NICA
        '   ~~~~~~~~~~~~~
            ElseIf CStr(vPedidoCalculoParcelas(i).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
                If vPedidoCalculoParcelas(i).pu_vencto_apos <> vPedidoCalculoParcelas(i + 1).pu_vencto_apos Then
                    If strMsgErro <> "" Then strMsgErro = strMsgErro & Chr(13)
                    strMsgErro = strMsgErro & "Diverg�ncia no prazo de vencimento da parcela �nica: " & _
                                 vPedidoCalculoParcelas(i).pedido & " (" & CStr(vPedidoCalculoParcelas(i).pu_vencto_apos) & " dia(s)) e " & _
                                 vPedidoCalculoParcelas(i + 1).pedido & " (" & CStr(vPedidoCalculoParcelas(i + 1).pu_vencto_apos) & " dia(s))"
                    End If
                End If
            Next
            
        If strMsgErro <> "" Then
            strMsgErro = "Os pedidos especificam diferentes prazos e/ou condi��es de pagamento para a mesma forma de pagamento: " & descricao_tipo_parcelamento(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) & "!!" & _
                        Chr(13) & _
                        Chr(13) & _
                        strMsgErro & _
                        Chr(13) & _
                        Chr(13) & _
                        "N�o � poss�vel gerar os dados de pagamento para impress�o na NFe!!"
            GoSub GDPP_FECHA_TABELAS
            Exit Function
            End If
        End If
        
       
  '�H� MAIS DO QUE 1 PEDIDO P/ SER PAGO POR BOLETO
    If intQtdePedidosPagtoBoleto > 1 Then
        s_where = ""
        For i = LBound(v_pedido) To UBound(v_pedido)
            If Trim$(v_pedido(i)) <> "" Then
                If s_where <> "" Then s_where = s_where & " OR"
                s_where = s_where & " (pedido='" & Trim$(v_pedido(i)) & "')"
                End If
            Next
        
        s = "SELECT DISTINCT" & _
                " id_plano_contas_empresa," & _
                " id_plano_contas_grupo," & _
                " id_plano_contas_conta," & _
                " natureza" & _
            " FROM t_PEDIDO tP" & _
                " INNER JOIN t_LOJA tL ON (tP.loja=tL.loja)" & _
            " WHERE" & _
                s_where
                
        If tAux.State <> adStateClosed Then tAux.Close
        tAux.Open s, dbc, , , adCmdText
        intQtdePlanoContas = 0
        Do While Not tAux.EOF
            intQtdePlanoContas = intQtdePlanoContas + 1
            tAux.MoveNext
            Loop
            
        If intQtdePlanoContas > 1 Then
            strMsgErro = "Os pedidos s�o de lojas que especificam diferentes planos de conta!!" & _
                        Chr(13) & _
                        Chr(13) & _
                        "N�o � poss�vel gerar os dados de pagamento para impress�o na NFe!!"
            GoSub GDPP_FECHA_TABELAS
            Exit Function
            End If
        End If
        
       
  '�HOUVE ALGUM ERRO?
    If strMsgErro <> "" Then
        GoSub GDPP_FECHA_TABELAS
        Exit Function
        End If
    
    
  '�OBT�M O VALOR TOTAL
  '�~~~~~~~~~~~~~~~~~~~
    For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
        With vPedidoCalculoParcelas(i)
            If Trim$(.pedido) <> "" Then
                vlTotalPedido = vlTotalPedido + .vlTotalDestePedido
            '   DADOS DO RATEIO NO CASO DE PAGAMENTO � VISTA
                If CStr(.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_A_VISTA) Then
                    If Trim$("" & v_parcela_pagto(0).strDadosRateio) <> "" Then v_parcela_pagto(0).strDadosRateio = v_parcela_pagto(0).strDadosRateio & "|"
                    v_parcela_pagto(0).strDadosRateio = v_parcela_pagto(0).strDadosRateio & .pedido & "=" & CStr(.vlTotalDestePedido)
                    End If
                End If
            End With
        Next
             

  '�CONSISTE VALOR TOTAL C/ A SOMA DOS VALORES DEFINIDOS NA FORMA DE PAGTO
  '�~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
        With vPedidoCalculoParcelas(i)
            If CStr(.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
                vlTotalFormaPagto = vlTotalFormaPagto + arredonda_para_monetario(.pu_valor * .razaoValorPedidoFilhote)
            ElseIf CStr(.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
                vlTotalFormaPagto = vlTotalFormaPagto + arredonda_para_monetario(.pce_entrada_valor * .razaoValorPedidoFilhote)
                vlTotalFormaPagto = vlTotalFormaPagto + CInt(.pce_prestacao_qtde) * arredonda_para_monetario(.pce_prestacao_valor * .razaoValorPedidoFilhote)
            ElseIf CStr(.tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
                vlTotalFormaPagto = vlTotalFormaPagto + arredonda_para_monetario(.pse_prim_prest_valor * .razaoValorPedidoFilhote)
                vlTotalFormaPagto = vlTotalFormaPagto + CInt(.pse_demais_prest_qtde) * arredonda_para_monetario(.pse_demais_prest_valor * .razaoValorPedidoFilhote)
                End If
            End With
        Next
        
    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_A_VISTA) Then
        vlTotalFormaPagto = vlTotalPedido
        End If
                
    vlDiferencaArredondamento = vlTotalPedido - vlTotalFormaPagto
    vlDiferencaArredondamentoRestante = vlDiferencaArredondamento
    
    If Abs(vlDiferencaArredondamento) > 1 Then
        strMsgErro = "A soma dos valores definidos na forma de pagamento (" & Format$(vlTotalFormaPagto, FORMATO_MOEDA) & ") n�o coincide com o valor total do(s) pedido(s) (" & Format$(vlTotalPedido, FORMATO_MOEDA) & ")!!" & _
                     Chr(13) & _
                     "N�o � poss�vel gerar os dados de pagamento para impress�o na NFe!!"
        GoSub GDPP_FECHA_TABELAS
        Exit Function
        End If
  
  '�CALCULA OS DADOS DAS PARCELAS DOS BOLETOS
  '�~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  '�LEMBRANDO QUE:
  '�    SE O PRAZO DEFINIDO PARA O 1� BOLETO FOR AT� 29 DIAS ENT�O:
  '�        VENCIMENTO = DATA EM QUE A NF EST� SENDO EMITIDA + 30 DIAS
  '�    SEN�O
  '�        VENCIMENTO = DATA EM QUE A NF EST� SENDO EMITIDA + PRAZO DEFINIDO PELO CLIENTE + 7 DIAS
  
'�  � VISTA
'   ~~~~~~~
    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_A_VISTA) Then
        With v_parcela_pagto(0)
            .intNumDestaParcela = 1
            .intNumTotalParcelas = 1
            .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).av_forma_pagto
            .vlValor = vlTotalPedido
            .dtVencto = Date + 30
            End With
        End If
        
        
'�  PARCELA �NICA
'   ~~~~~~~~~~~~~
    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELA_UNICA) Then
        With v_parcela_pagto(0)
            .intNumDestaParcela = 1
            .intNumTotalParcelas = 1
            .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pu_forma_pagto
            .dtVencto = calculaDataPrimeiroBoleto(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pu_vencto_apos)
            For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
                .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(i).pu_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
                If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
                .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(i).pedido & "=" & CStr(arredonda_para_monetario(vPedidoCalculoParcelas(i).pu_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote))
                Next
            End With
        End If
        
        
'   PARCELADO COM ENTRADA
'   ~~~~~~~~~~~~~~~~~~~~~
    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) Then
      '�ENTRADA
        With v_parcela_pagto(0)
            .intNumDestaParcela = 1
            intQtdeTotalParcelas = 1
            .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_entrada
            End With
        
      '�ENTRADA � POR BOLETO?
        If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_entrada) = CStr(ID_FORMA_PAGTO_BOLETO) Then
            dtUltimoPagtoCalculado = Date + 30
        Else
            dtUltimoPagtoCalculado = Date
            End If
        
        With v_parcela_pagto(0)
            .dtVencto = dtUltimoPagtoCalculado
            For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
                .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(i).pce_entrada_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
                vlRateio = arredonda_para_monetario(vPedidoCalculoParcelas(i).pce_entrada_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
                If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_entrada) = CStr(ID_FORMA_PAGTO_BOLETO) Then
                    If vlDiferencaArredondamentoRestante <> 0 Then
                        .vlValor = .vlValor + vlDiferencaArredondamentoRestante
                        vlRateio = vlRateio + vlDiferencaArredondamentoRestante
                        vlDiferencaArredondamentoRestante = 0
                        End If
                    End If
                If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
                .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(i).pedido & "=" & CStr(vlRateio)
                Next
            End With
        
      '�PRESTA��ES
        For i = 1 To vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_qtde
            intQtdeTotalParcelas = intQtdeTotalParcelas + 1
            If v_parcela_pagto(UBound(v_parcela_pagto)).intNumDestaParcela <> 0 Then
                ReDim Preserve v_parcela_pagto(UBound(v_parcela_pagto) + 1)
                End If
            
            With v_parcela_pagto(UBound(v_parcela_pagto))
                .intNumDestaParcela = intQtdeTotalParcelas
                .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_prestacao
                End With
            
        '   PRESTA��ES S�O POR BOLETO?
            If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_prestacao) = CStr(ID_FORMA_PAGTO_BOLETO) Then
            '   A ENTRADA N�O FOI PAGA POR BOLETO!
                If intQtdeTotalParcelas = 1 Then
                '   ESTA PRESTA��O SER� O 1� BOLETO DA S�RIE
                    If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo) = CInt(30) Then
                        dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
                    ElseIf CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo) <= 29 Then
                        dtUltimoPagtoCalculado = DateAdd("d", 30, dtUltimoPagtoCalculado)
                    Else
                        dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo), dtUltimoPagtoCalculado)
                        End If
                Else
                  '�CALCULA A DATA DOS DEMAIS BOLETOS
                    If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo) = CInt(30) Then
                        dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
                    Else
                        dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo), dtUltimoPagtoCalculado)
                        End If
                    End If
            Else
            '   C�LCULO P/ PRESTA��ES QUE N�O S�O POR BOLETO
                If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo) = CInt(30) Then
                    dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
                Else
                    dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_prestacao_periodo), dtUltimoPagtoCalculado)
                    End If
                End If
            
            With v_parcela_pagto(UBound(v_parcela_pagto))
                .dtVencto = dtUltimoPagtoCalculado
                
                For j = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
                    .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(j).pce_prestacao_valor * vPedidoCalculoParcelas(j).razaoValorPedidoFilhote)
                    vlRateio = arredonda_para_monetario(vPedidoCalculoParcelas(j).pce_prestacao_valor * vPedidoCalculoParcelas(j).razaoValorPedidoFilhote)
                    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pce_forma_pagto_prestacao) = CStr(ID_FORMA_PAGTO_BOLETO) Then
                        If vlDiferencaArredondamentoRestante <> 0 Then
                            .vlValor = .vlValor + vlDiferencaArredondamentoRestante
                            vlRateio = vlRateio + vlDiferencaArredondamentoRestante
                            vlDiferencaArredondamentoRestante = 0
                            End If
                        End If
                    If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
                    .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(j).pedido & "=" & CStr(vlRateio)
                    Next
                End With
            Next
        
        For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
            v_parcela_pagto(i).intNumTotalParcelas = intQtdeTotalParcelas
            Next
        End If
        
        
'   PARCELADO SEM ENTRADA
'   ~~~~~~~~~~~~~~~~~~~~~
    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).tipo_parcelamento) = CStr(COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) Then
    '   1� PRESTA��O
        With v_parcela_pagto(0)
            .intNumDestaParcela = 1
            intQtdeTotalParcelas = 1
            .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_prim_prest
            End With
      
    '�  1� PRESTA��O � POR BOLETO?
        If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_prim_prest) = CStr(ID_FORMA_PAGTO_BOLETO) Then
            With v_parcela_pagto(0)
                dtUltimoPagtoCalculado = calculaDataPrimeiroBoleto(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_prim_prest_apos)
                End With
        Else
            dtUltimoPagtoCalculado = DateAdd("d", vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_prim_prest_apos, Date)
            End If
            
        With v_parcela_pagto(0)
            .dtVencto = dtUltimoPagtoCalculado
            For i = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
                .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(i).pse_prim_prest_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
                vlRateio = arredonda_para_monetario(vPedidoCalculoParcelas(i).pse_prim_prest_valor * vPedidoCalculoParcelas(i).razaoValorPedidoFilhote)
                If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_prim_prest) = CStr(ID_FORMA_PAGTO_BOLETO) Then
                    If vlDiferencaArredondamentoRestante <> 0 Then
                        .vlValor = .vlValor + vlDiferencaArredondamentoRestante
                        vlRateio = vlRateio + vlDiferencaArredondamentoRestante
                        vlDiferencaArredondamentoRestante = 0
                        End If
                    End If
                If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
                .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(i).pedido & "=" & CStr(vlRateio)
                Next
            End With
            
    '�  DEMAIS PRESTA��ES
        For i = 1 To vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_qtde
            intQtdeTotalParcelas = intQtdeTotalParcelas + 1
            If v_parcela_pagto(UBound(v_parcela_pagto)).intNumDestaParcela <> 0 Then
                ReDim Preserve v_parcela_pagto(UBound(v_parcela_pagto) + 1)
                End If
        
            With v_parcela_pagto(UBound(v_parcela_pagto))
                .intNumDestaParcela = intQtdeTotalParcelas
                .id_forma_pagto = vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_demais_prest
                End With

        '�  DEMAIS PRESTA��ES S�O POR BOLETO?
            If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_demais_prest) = CStr(ID_FORMA_PAGTO_BOLETO) Then
            '   A 1� PRESTA��O N�O FOI PAGA POR BOLETO!
                If intQtdeTotalParcelas = 1 Then
                '�  ESTA PRESTA��O SER� O 1� BOLETO DA S�RIE
                    If (CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_prim_prest_apos) + _
                        CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo)) >= 30 Then
                        
                        If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo) = CInt(30) Then
                            dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
                        Else
                            dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo), dtUltimoPagtoCalculado)
                            End If
                    Else
                        dtUltimoPagtoCalculado = DateAdd("d", 30, Date)
                        End If
                Else
                  '�CALCULA A DATA DOS DEMAIS BOLETOS
                    If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo) = CInt(30) Then
                        dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
                    Else
                        dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo), dtUltimoPagtoCalculado)
                        End If
                    End If
            Else
            '   C�LCULO P/ PRESTA��ES QUE N�O S�O POR BOLETO
                If CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo) = CInt(30) Then
                    dtUltimoPagtoCalculado = DateAdd("m", 1, dtUltimoPagtoCalculado)
                Else
                    dtUltimoPagtoCalculado = DateAdd("d", CInt(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_demais_prest_periodo), dtUltimoPagtoCalculado)
                    End If
                End If
                
            With v_parcela_pagto(UBound(v_parcela_pagto))
                .dtVencto = dtUltimoPagtoCalculado
                For j = LBound(vPedidoCalculoParcelas) To UBound(vPedidoCalculoParcelas)
                    .vlValor = .vlValor + arredonda_para_monetario(vPedidoCalculoParcelas(j).pse_demais_prest_valor * vPedidoCalculoParcelas(j).razaoValorPedidoFilhote)
                    vlRateio = arredonda_para_monetario(vPedidoCalculoParcelas(j).pse_demais_prest_valor * vPedidoCalculoParcelas(j).razaoValorPedidoFilhote)
                    If CStr(vPedidoCalculoParcelas(UBound(vPedidoCalculoParcelas)).pse_forma_pagto_demais_prest) = CStr(ID_FORMA_PAGTO_BOLETO) Then
                        If vlDiferencaArredondamentoRestante <> 0 Then
                            .vlValor = .vlValor + vlDiferencaArredondamentoRestante
                            vlRateio = vlRateio + vlDiferencaArredondamentoRestante
                            vlDiferencaArredondamentoRestante = 0
                            End If
                        End If
                    If Trim$("" & .strDadosRateio) <> "" Then .strDadosRateio = .strDadosRateio & "|"
                    .strDadosRateio = .strDadosRateio & vPedidoCalculoParcelas(j).pedido & "=" & CStr(vlRateio)
                    Next
                End With
            Next
            
        For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
            v_parcela_pagto(i).intNumTotalParcelas = intQtdeTotalParcelas
            Next
        End If
    
        
    geraDadosParcelasPagto = True
    
    GoSub GDPP_FECHA_TABELAS
    
Exit Function
    
    
    
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GDPP_TRATA_ERRO:
'===============
    strMsgErro = CStr(Err) & ": " & Error$(Err)
    GoSub GDPP_FECHA_TABELAS
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GDPP_FECHA_TABELAS:
'==================
  '�RECORDSETS
    bd_desaloca_recordset t_PEDIDO, True
    bd_desaloca_recordset t_PEDIDO_ITEM, True
    bd_desaloca_recordset tAux, True
    Return
    
    
End Function

Function gravaDadosParcelaPagto(ByVal numNF As Long, v_parcela_pagto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO, ByRef strMsgErro As String) As Boolean
'�__________________________________________________________________________________________
'|
'|  GRAVA AS INFORMA��ES DOS BOLETOS NO BANCO DE DADOS
'|

Dim s As String
Dim s_where As String
Dim s_pedido_aux As String
Dim i As Integer
Dim j As Integer
Dim intNsuNfParcelaPagto As Long
Dim intNsuNfParcelaPagtoItem As Long
Dim intQtdeParcelas As Integer
Dim intQtdeParcelasBoleto As Integer
Dim intRecordsAffected As Long
Dim strIdCliente As String
Dim v_pedido() As String
Dim v_pedido_aux() As String

'�BANCO DE DADOS
Dim t As ADODB.Recordset

    On Error GoTo GDPP_TRATA_ERRO

    gravaDadosParcelaPagto = False
    
    strMsgErro = ""
    
'   TEM DADOS P/ GRAVAR?
    intQtdeParcelas = 0
    For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
        If v_parcela_pagto(i).intNumDestaParcela > 0 Then
            intQtdeParcelas = intQtdeParcelas + 1
            End If
        
        If CStr(v_parcela_pagto(i).id_forma_pagto) = CStr(ID_FORMA_PAGTO_BOLETO) Then
            intQtdeParcelasBoleto = intQtdeParcelasBoleto + 1
            End If
        Next
        
    If (intQtdeParcelas = 0) Then
        gravaDadosParcelaPagto = True
        Exit Function
        End If
        
'   RECORDSET
    Set t = New ADODB.Recordset
    With t
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
'   OBT�M IDENTIFICA��O DO CLIENTE
'   LEMBRANDO QUE GARANTIDAMENTE TODOS OS PEDIDOS S�O DO MESMO CLIENTE
    v_pedido = Split(v_parcela_pagto(UBound(v_parcela_pagto)).strDadosRateio, "|")
    v_pedido_aux = Split(v_pedido(LBound(v_pedido)), "=")
    s_pedido_aux = Trim$(v_pedido_aux(LBound(v_pedido_aux)))
    
    s = "SELECT" & _
            " c.id" & _
        " FROM t_PEDIDO p" & _
            " INNER JOIN t_CLIENTE c" & _
                " ON p.id_cliente=c.id" & _
        " WHERE" & _
            " p.pedido = '" & s_pedido_aux & "'"
    If t.State <> adStateClosed Then t.Close
    t.Open s, dbc, , , adCmdText
    If Not t.EOF Then
        strIdCliente = Trim$("" & t("id"))
    Else
        strMsgErro = "Falha ao tentar localizar a identifica��o do cliente!!"
        GoSub GDPP_FECHA_TABELAS
        Exit Function
        End If
            
            
'   GRAVA REGISTRO PRINCIPAL
'   ~~~~~~~~~~~~~~~
    dbc.BeginTrans
'   ~~~~~~~~~~~~~~~
'   SE HOUVER DADOS DE PARCELAS CADASTRADOS ANTERIORMENTE NO STATUS INICIAL P/ ESTE(S) PEDIDO(S),
'   CANCELA-OS ANTES DE CADASTRAR OS NOVOS DADOS
    s_where = ""
    For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
        With v_parcela_pagto(i)
            If .intNumDestaParcela <> 0 Then
                v_pedido = Split(.strDadosRateio, "|")
                For j = LBound(v_pedido) To UBound(v_pedido)
                    If Trim$(v_pedido(j)) <> "" Then
                        v_pedido_aux = Split(v_pedido(j), "=")
                        s_pedido_aux = Trim$(v_pedido_aux(LBound(v_pedido_aux)))
                        If s_pedido_aux <> "" Then
                            If InStr(s_where, s_pedido_aux) = 0 Then
                                If s_where <> "" Then s_where = s_where & " OR"
                                s_where = s_where & " (pedido='" & Trim$(v_pedido_aux(LBound(v_pedido_aux))) & "')"
                                End If
                            End If
                        End If
                    Next
                End If
            End With
        Next
        
    If s_where <> "" Then
        s = "SELECT DISTINCT" & _
                " tpp.id" & _
            " FROM t_FIN_NF_PARCELA_PAGTO tpp" & _
                " INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM tppi" & _
                    " ON (tpp.id=tppi.id_nf_parcela_pagto)" & _
                " INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO tppir" & _
                    " ON (tppi.id=tppir.id_nf_parcela_pagto_item)" & _
            " WHERE" & _
                " (tpp.status = " & NF_PARCELA_PAGTO__STATUS_INICIAL & ")" & _
                " AND (" & s_where & ")"
        If t.State <> adStateClosed Then t.Close
        t.Open s, dbc, , , adCmdText
        Do While Not t.EOF
            s = "UPDATE" & _
                    " t_FIN_NF_PARCELA_PAGTO" & _
                " SET" & _
                    " status = " & NF_PARCELA_PAGTO__STATUS_CANCELADO & _
                " WHERE" & _
                    " (id = " & t("id") & ")" & _
                    " AND (status = " & NF_PARCELA_PAGTO__STATUS_INICIAL & ")"
            Call dbc.Execute(s, intRecordsAffected)
            If intRecordsAffected = 0 Then
                strMsgErro = "Falha ao tentar cancelar registros anteriores dos dados de pagamento do(s) pedido(s) especificado(s)!!"
            '   ~~~~~~~~~~~~~~~~~
                dbc.RollbackTrans
            '   ~~~~~~~~~~~~~~~~~
                GoSub GDPP_FECHA_TABELAS
                Exit Function
                End If
            t.MoveNext
            Loop
        End If
        
'   OBT�M NSU
    If Not geraNsu(NSU_T_FIN_NF_PARCELA_PAGTO, intNsuNfParcelaPagto, strMsgErro) Then
        If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
        strMsgErro = "Falha ao gravar os dados de pagamento!!" & strMsgErro
    '   ~~~~~~~~~~~~~~~~~
        dbc.RollbackTrans
    '   ~~~~~~~~~~~~~~~~~
        GoSub GDPP_FECHA_TABELAS
        Exit Function
        End If

    On Error GoTo GDPP_TRATA_ERRO_TRANSACAO
'   LEMBRANDO QUE DT_CADASTRO, DT_ULT_ATUALIZACAO E STATUS S�O INSERIDOS VIA DEFAULT DAS COLUNAS
    s = "INSERT INTO t_FIN_NF_PARCELA_PAGTO (" & _
            "id," & _
            "id_cliente," & _
            "numero_NF," & _
            "qtde_parcelas," & _
            "qtde_parcelas_boleto," & _
            "usuario_cadastro," & _
            "usuario_ult_atualizacao" & _
        ") VALUES (" & _
            CStr(intNsuNfParcelaPagto) & "," & _
            "'" & strIdCliente & "'," & _
            CStr(numNF) & "," & _
            CStr(intQtdeParcelas) & "," & _
            CStr(intQtdeParcelasBoleto) & "," & _
            "'" & Trim$(usuario.id) & "'," & _
            "'" & Trim$(usuario.id) & "'" & _
        ")"
    Call dbc.Execute(s, intRecordsAffected)
    If intRecordsAffected = 0 Then
        strMsgErro = "Falha ao tentar inserir registro principal dos dados de pagamento!!"
    '   ~~~~~~~~~~~~~~~~~
        dbc.RollbackTrans
    '   ~~~~~~~~~~~~~~~~~
        GoSub GDPP_FECHA_TABELAS
        Exit Function
        End If

'   GRAVA REGISTRO DAS PARCELAS
    For i = LBound(v_parcela_pagto) To UBound(v_parcela_pagto)
    '   OBT�M NSU
        If Not geraNsu(NSU_T_FIN_NF_PARCELA_PAGTO_ITEM, intNsuNfParcelaPagtoItem, strMsgErro) Then
            If strMsgErro <> "" Then strMsgErro = Chr(13) & Chr(13) & strMsgErro
            strMsgErro = "Falha ao gravar os dados de pagamento!!" & strMsgErro
        '   ~~~~~~~~~~~~~~~~~
            dbc.RollbackTrans
        '   ~~~~~~~~~~~~~~~~~
            GoSub GDPP_FECHA_TABELAS
            Exit Function
            End If
    
        With v_parcela_pagto(i)
            If .intNumDestaParcela <> 0 Then
                s = "INSERT INTO t_FIN_NF_PARCELA_PAGTO_ITEM (" & _
                        "id," & _
                        "id_nf_parcela_pagto," & _
                        "num_parcela," & _
                        "forma_pagto," & _
                        "dt_vencto," & _
                        "valor" & _
                    ") VALUES (" & _
                        CStr(intNsuNfParcelaPagtoItem) & "," & _
                        CStr(intNsuNfParcelaPagto) & "," & _
                        CStr(.intNumDestaParcela) & "," & _
                        CStr(.id_forma_pagto) & "," & _
                        sqlMontaDateParaSqlDateTime(.dtVencto) & "," & _
                        sqlFormataDecimal(.vlValor) & _
                    ")"
                Call dbc.Execute(s, intRecordsAffected)
                If intRecordsAffected = 0 Then
                    strMsgErro = "Falha ao tentar inserir registro da parcela " & .intNumDestaParcela & "!!"
                '   ~~~~~~~~~~~~~~~~~
                    dbc.RollbackTrans
                '   ~~~~~~~~~~~~~~~~~
                    GoSub GDPP_FECHA_TABELAS
                    Exit Function
                    End If
                
                v_pedido = Split(.strDadosRateio, "|")
                For j = LBound(v_pedido) To UBound(v_pedido)
                    If Trim$(v_pedido(j)) <> "" Then
                        v_pedido_aux = Split(v_pedido(j), "=")
                        s = "INSERT INTO t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO (" & _
                                "id_nf_parcela_pagto_item," & _
                                "pedido," & _
                                "id_nf_parcela_pagto," & _
                                "valor" & _
                            ") VALUES (" & _
                                CStr(intNsuNfParcelaPagtoItem) & "," & _
                                "'" & Trim$(v_pedido_aux(LBound(v_pedido_aux))) & "'," & _
                                CStr(intNsuNfParcelaPagto) & "," & _
                                sqlFormataDecimal(CCur(Trim$(v_pedido_aux(UBound(v_pedido_aux))))) & _
                            ")"
                        Call dbc.Execute(s, intRecordsAffected)
                        If intRecordsAffected = 0 Then
                            strMsgErro = "Falha ao tentar inserir registro do rateio da parcela " & .intNumDestaParcela & "!!"
                        '   ~~~~~~~~~~~~~~~~~
                            dbc.RollbackTrans
                        '   ~~~~~~~~~~~~~~~~~
                            GoSub GDPP_FECHA_TABELAS
                            Exit Function
                            End If
                        End If
                    Next
                End If
            End With
        Next
    
'   ~~~~~~~~~~~~~~~
    dbc.CommitTrans
'   ~~~~~~~~~~~~~~~
    On Error GoTo GDPP_TRATA_ERRO
    
    gravaDadosParcelaPagto = True

    GoSub GDPP_FECHA_TABELAS
    
Exit Function
    
    
    
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GDPP_TRATA_ERRO:
'===============
    strMsgErro = CStr(Err) & ": " & Error$(Err)
    GoSub GDPP_FECHA_TABELAS
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GDPP_TRATA_ERRO_TRANSACAO:
'=========================
    strMsgErro = CStr(Err) & ": " & Error$(Err)
    On Error Resume Next
    dbc.RollbackTrans
    GoSub GDPP_FECHA_TABELAS
    Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GDPP_FECHA_TABELAS:
'==================
  '�RECORDSETS
    bd_desaloca_recordset t, True
    Return

End Function


Function consultaDadosParcelasPagto(v_pedido() As String, ByRef v_parcela_pagto() As TIPO_NF_LINHA_DADOS_PARCELA_PAGTO, ByRef strMsgErro As String) As Boolean
'�__________________________________________________________________________________________
'|
'|  CONSULTA PEDIDOS COM PARCELAMENTO VIA BOLETO PARA EXIBI��O.
'|
'|

Dim s As String
Dim i As Integer

'�BANCO DE DADOS
Dim rs As ADODB.Recordset

    On Error GoTo CDPP_TRATA_ERRO

    consultaDadosParcelasPagto = False
    
    strMsgErro = ""
    ReDim v_parcela_pagto(0)
    
    ReDim v_parcela_pagto(0)
    v_parcela_pagto(UBound(v_parcela_pagto)).intNumDestaParcela = 0
  
  '�rs
    Set rs = New ADODB.Recordset
    With rs
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
    For i = LBound(v_pedido) To UBound(v_pedido)
        If Trim$(v_pedido(i)) <> "" Then
            s = "select i.num_parcela, i.forma_pagto, i.valor, i.dt_vencto " & _
                "from t_FIN_NF_PARCELA_PAGTO_ITEM i " & _
                "inner join t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO ir on i.id = ir.id_nf_parcela_pagto_item " & _
                "where ir.pedido = '" & v_pedido(i) & "' "
            If rs.State <> adStateClosed Then rs.Close
            rs.Open s, dbc, , , adCmdText
            If rs.EOF Then
                If strMsgErro <> "" Then strMsgErro = strMsgErro & vbCrLf
                strMsgErro = strMsgErro & "Pedido " & Trim$(v_pedido(i)) & " n�o est� cadastrado!!"
            Else
                Do While Not rs.EOF
                    If v_parcela_pagto(UBound(v_parcela_pagto)).intNumDestaParcela <> 0 Then
                        ReDim Preserve v_parcela_pagto(UBound(v_parcela_pagto) + 1)
                        End If
                    With v_parcela_pagto(UBound(v_parcela_pagto))
                        .intNumDestaParcela = rs("num_parcela")
                        .id_forma_pagto = rs("num_parcela")
                        .vlValor = rs("valor")
                        .dtVencto = rs("dt_vencto")
                        End With
                    rs.MoveNext
                    Loop
                End If
            End If
        Next
        
    consultaDadosParcelasPagto = True
    
    GoSub CDPP_FECHA_TABELAS
    
Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CDPP_TRATA_ERRO:
'===============
    strMsgErro = strMsgErro & vbCrLf & CStr(Err) & ": " & Error$(Err)
    GoSub CDPP_FECHA_TABELAS
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CDPP_FECHA_TABELAS:
'==================
  '�RECORDSETS
    bd_desaloca_recordset rs, True
    Return
    
    
End Function



Function ExisteDadosParcelasPagto(pedido As String, ByRef strMsgErro As String) As Boolean
'�__________________________________________________________________________________________
'|
'|  VERIFICA SE EXISTEM PEDIDOS COM PARCELAMENTO VIA BOLETO PARA EXIBI��O.
'|
'|

Dim s As String
Dim i As Integer
Dim ok As Boolean

'�BANCO DE DADOS
Dim rs As ADODB.Recordset

    On Error GoTo EDPP_TRATA_ERRO

    ok = False
    
    strMsgErro = ""
  '�rs
    Set rs = New ADODB.Recordset
    With rs
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
    If Trim$(pedido) <> "" Then
        s = "select i.num_parcela, i.forma_pagto, i.valor, i.dt_vencto " & _
            "from t_FIN_NF_PARCELA_PAGTO_ITEM i " & _
            "inner join t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO ir on i.id = ir.id_nf_parcela_pagto_item " & _
            "where ir.pedido = '" & pedido & "' "
        If rs.State <> adStateClosed Then rs.Close
        rs.Open s, dbc, , , adCmdText
        If rs.RecordCount > 0 Then
            ok = True
            End If
        End If
        
    ExisteDadosParcelasPagto = ok
    
    GoSub EDPP_FECHA_TABELAS
    
Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EDPP_TRATA_ERRO:
'===============
    strMsgErro = strMsgErro & vbCrLf & CStr(Err) & ": " & Error$(Err)
    GoSub EDPP_FECHA_TABELAS
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EDPP_FECHA_TABELAS:
'==================
  '�RECORDSETS
    bd_desaloca_recordset rs, True
    Return
    
    
End Function



Function ExisteDadosParcelasPagtobkp(v_pedido() As String, ByRef strMsgErro As String) As Boolean
'�__________________________________________________________________________________________
'|
'|  VERIFICA SE EXISTEM PEDIDOS COM PARCELAMENTO VIA BOLETO PARA EXIBI��O.
'|
'|

Dim s As String
Dim i As Integer
Dim ok As Boolean

'�BANCO DE DADOS
Dim rs As ADODB.Recordset

    On Error GoTo EDPP_TRATA_ERRO

    ok = False
    
    strMsgErro = ""
    ReDim v_parcela_pagto(0)
    
    ReDim v_parcela_pagto(0)
    v_parcela_pagto(UBound(v_parcela_pagto)).intNumDestaParcela = 0
  
  '�rs
    Set rs = New ADODB.Recordset
    With rs
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
        
    For i = LBound(v_pedido) To UBound(v_pedido)
            If Trim$(v_pedido(i)) <> "" Then
            s = "select i.num_parcela, i.forma_pagto, i.valor, i.dt_vencto " & _
                "from t_FIN_NF_PARCELA_PAGTO_ITEM i " & _
                "inner join t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO ir on i.id = ir.id_nf_parcela_pagto_item " & _
                "where ir.pedido = '" & v_pedido(i) & "' "
            If rs.State <> adStateClosed Then rs.Close
            rs.Open s, dbc, , , adCmdText
            If Not rs.EOF Then
                ok = True
                End If
            End If
        Next
        
    ExisteDadosParcelasPagtobkp = ok
    
    GoSub EDPP_FECHA_TABELAS
    
Exit Function
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EDPP_TRATA_ERRO:
'===============
    strMsgErro = strMsgErro & vbCrLf & CStr(Err) & ": " & Error$(Err)
    GoSub EDPP_FECHA_TABELAS
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EDPP_FECHA_TABELAS:
'==================
  '�RECORDSETS
    bd_desaloca_recordset rs, True
    Return
    
    
End Function







