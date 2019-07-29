Attribute VB_Name = "mod_ADM"
Option Explicit

'===============================================================================
'
'                                 MÓDULO ADM
'                          ADMINISTRAÇÃO E MANUTENÇÃO
'                       _______________________________
' 
'                           EDIÇÃO = 029
'                           DATA   = 10.OUT.2018
'                       _______________________________
' 
' 
'===============================================================================
'
'                Módulo ADM foi projetado pela
' 
'                   PRAGMÁTICA Engenheiros Consultores Associados,
'                              Serviços e Comércio Ltda.
'                   Rua Leandro Dupré, nº 204 - cj. 153/154
'                   04025-010  -  São Paulo  -  SP
'                   Telefone (11)5084-6772
'                   Fax      (11)5084-6772
'                   CNPJ nº 65.486.524/0001-80
' 
' _____________________________________________________________________________
'|                                                                             |
'|  H I S T Ó R I C O    D A S    A L T E R A Ç Õ E S                          |
'|_____________________________________________________________________________|
'|          |      |                                                           |
'|  DATA    | RESP | ALTERAÇÃO                                                 |
'|__________|______|___________________________________________________________|
'|05.08.2003| HHO  | Início v 1.00                                             |
'|__________|______|___________________________________________________________|
'|08.08.2003| HHO  | V 1.01 Alteração do modo de ler os dados da planilha, que |
'|          |      | antes era célula a célula e agora é feito através de uma  |
'|          |      | matriz que recebe toda a linha de uma só vez.  O objeto   |
'|          |      | usado para ler os dados continua sendo o "range".         |
'|          |      | Isso foi feito para evitar o problema em que o Excel trava|
'|          |      | durante o processamento.  Essa alteração também melhorou  |
'|          |      | o desempenho do programa.                                 |
'|          |      | O tempo do parâmetro "App.OleRequestPendingTimeout" foi   |
'|          |      | aumentado p/ evitar a mensagem de erro avisando que o co- |
'|          |      | mando não pode ser completado porque o Excel está ocupado.|
'|__________|______|___________________________________________________________|
'|29.08.2003| HHO  | V 1.02 Implementação da operação de limpeza do banco de   |
'|          |      | dados.                                                    |
'|__________|______|___________________________________________________________|
'|08.09.2003| HHO  | V 1.03 Liberação do acesso ao módulo para o usuário de ní-|
'|          |      | vel "administrador".                                      |
'|__________|______|___________________________________________________________|
'|05.11.2003| HHO  | V 1.04 Implementação de tratamento para carregar a tabela |
'|          |      | de produtos para uma loja ou um grupo de lojas.           |
'|          |      | Além disso, foi introduzido uma painel para permitir a    |
'|          |      | seleção de quais planilhas devem ser carregadas.          |
'|__________|______|___________________________________________________________|
'|04.12.2003| HHO  | V 1.05 Implementação de tratamento para eliminar orçamen- |
'|          |      | tos antigos na rotina de limpeza do banco de dados.       |
'|          |      | Implementação da operação de backup do banco de dados em  |
'|          |      | arquivos gravados na máquina local.                       |
'|          |      | Implementação da operação de restauração do banco de dados|
'|          |      | a partir dos arquivos de dados gerados pela rotina de     |
'|          |      | backup.  Para que a operação de restauração fique disponí-|
'|          |      | vel no painel do módulo, é necessário que as tabelas      |
'|          |      | T_CONTROLE, T_VERSAO e T_USUARIO estejam vazias. Para que |
'|          |      | a restauração do backup seja processada, é necessário que |
'|          |      | todas as tabelas que são salvas no backup estejam vazias. |
'|          |      | A rotina não faz a limpeza das tabelas automaticamente por|
'|          |      | uma questão de segurança.                                 |
'|__________|______|___________________________________________________________|
'|05.12.2003| HHO  | V 1.06 Correção no tratamento do evento form activate do  |
'|          |      | painel principal para que não seja processado 2 vezes.    |
'|          |      | Isso pode ocorrer se, após o fechamento do painel de lo-  |
'|          |      | gin, for exibido um aviso através de message box em algu- |
'|          |      | ma das verificações feitas em seguida.                    |
'|          |      | Neste caso, o 2º evento começa a ser processado no momento|
'|          |      | em que a message box é exibida, ou seja, o processamento  |
'|          |      | do 1º evento é interrompido nesse ponto.                  |
'|          |      | Após concluir o processamento do 2º evento, o 1º evento   |
'|          |      | volta a ser processado.                                   |
'|          |      | O efeito causado neste caso é que a tela de login e as    |
'|          |      | mensagens de alerta são exibidas 2 vezes.                 |
'|          |      | Isso ocorre somente no executável (lembre-se de que exis- |
'|          |      | tem diferenças entre executar dentro do Visual Basic e    |
'|          |      | através do executável).                                   |
'|__________|______|___________________________________________________________|
'|23.09.2004| HHO  |V 1.07 Implementação de tratamento para eliminar registros |
'|          |      |antigos de transações de pagamento pela Visanet.           |
'|          |      |Execução automática da limpeza de registros de log antigos |
'|          |      |ao acionar o módulo.                                       |
'|          |      |Aumento do prazo de permanência dos pedidos, orçamentos,   |
'|          |      |etc ao realizar a limpeza do banco de dados.               |
'|__________|______|___________________________________________________________|
'|21.09.2007| HHO  |V 1.08 Implementação de variável global (DESENVOLVIMENTO)  |
'|          |      |para diferenciar a execução no "ambiente de desenvolvimen- |
'|          |      |to" e no "ambiente de produção".                           |
'|          |      |Com isso, ao se fazer o backup dos dados de produção para  |
'|          |      |atualizar a base de desenvolvimento, não se executa rotinas|
'|          |      |como a limpeza automática da tabela de log.                |
'|__________|______|___________________________________________________________|
'|08.01.2008| HHO  |V 1.09 Implementação da limpeza automática das tabelas:    |
'|          |      |    t_SESSAO_ABANDONADA                                    |
'|          |      |    t_SESSAO_HISTORICO                                     |
'|          |      |    t_SESSAO_RESTAURADA                                    |
'|__________|______|___________________________________________________________|
'|07.02.2008| HHO  |V 1.10 Implementação da limpeza automática das tabelas:    |
'|          |      |    t_ESTOQUE_LOG                                          |
'|          |      |    t_ESTOQUE_SALDO_DIARIO                                 |
'|__________|______|___________________________________________________________|
'|23.06.2008| HHO  |V 1.11 Implementação da nova coluna de 'custo2' na planilha|
'|          |      |de produtos.                                               |
'|__________|______|___________________________________________________________|
'|06.08.2008| HHO  |V 1.11B Como a alteração do Produto Composto ainda não foi |
'|          |      |publicada, as modificações feitas devido ao novo site      |
'|          |      |Artven (Fabricante) foram introduzidas na versão anterior  |
'|          |      |ao do Produto Composto (v1.12).                            |
'|          |      |Portanto, foi implementada a exibição da informação sobre  |
'|          |      |onde está conectado:                                       |
'|          |      |   A) Artven Bonshop (Artven3)                             |
'|          |      |   B) Artven Fabricante (Artven)                           |
'|__________|______|___________________________________________________________|
'|14.07.2008| HHO  |V 1.12 Implementação de atualização automática dos valores |
'|          |      |dos produtos compostos com base nos valores dos seus produ-|
'|          |      |tos componentes (t_PRODUTO e t_PRODUTO_LOJA).              |
'|          |      |Este recurso não foi implantado e foi cancelado.           |
'|__________|______|___________________________________________________________|
'|04.02.2009| HHO  |V 1.13 Implementação das seguintes melhorias:              |
'|          |      |   1) Nova coluna na planilha básica de produtos p/ exibir |
'|          |      |      mensagens de alerta. Nesta nova coluna serão cadas-  |
'|          |      |      trados os códigos das mensagens de alerta que devem  |
'|          |      |      ser exibidos no sistema Artven3. Os códigos devem    |
'|          |      |      ser separados por vírgula ou ponto e vírgula.        |
'|          |      |   2) Armazenamento da descricação do produto com texto    |
'|          |      |      formatado (negrito, itálico, sublinhado) no novo     |
'|          |      |      campo t_PRODUTO.descricao_html                       |
'|__________|______|___________________________________________________________|
'|03.08.2009| HHO  |V 1.14 Ampliação do prazo de permanência dos dados.        |
'|__________|______|___________________________________________________________|
'|12.03.2010| HHO  |V 1.15 Criação de novas colunas na planilha de produtos.   |
'|          |      |   A) Cubagem: informação sobre as dimensões do produto p/ |
'|          |      |      facilitar a administração dos valores de frete.      |
'|          |      |   B) NCM: código NCM do produto (NFe)                     |
'|          |      |   C) CST: código CST do produto (NFe)                     |
'|__________|______|___________________________________________________________|
'|28.04.2010| HHO  |V 1.16 Alteração da rotina de carga dos dados da planilha  |
'|          |      |de produtos para minimizar o tempo em que os registros fi- |
'|          |      |cam bloqueados para os demais usuários do sistema.         |
'|          |      |Para isso, os dados das tabelas de produção são copiados p/|
'|          |      |tabelas temporárias e estas é que são utilizadas durante o |
'|          |      |processamento. Ao final, os dados das tabelas de produção  |
'|          |      |são substituídos pelos dados já processados contidos nas   |
'|          |      |tabelas temporárias, dentro de uma sessão de transação.    |
'|__________|______|___________________________________________________________|
'|08.06.2010| HHO  |V 1.17 Criação de novas coluna na planilha de produtos.    |
'|          |      |   A) Percentual de margem de valor adicionado do ICMS ST  |
'|          |      |                                                           |
'|          |      |A nova informação é necessária para os produtos de fabri-  |
'|          |      |cação própria, pois neste caso é necessário calcular na    |
'|          |      |NFe a base de cálculo da substituição tributária e o valor |
'|          |      |do imposto a ser recolhido por ST.                         |
'|__________|______|___________________________________________________________|
'|16.09.2010| HHO  |V 1.18 Retirada das rotinas que realizam a limpeza das ta- |
'|          |      |belas:                                                     |
'|          |      |   t_LOG                                                   |
'|          |      |   t_ESTOQUE_LOG                                           |
'|          |      |   t_ESTOQUE_SALDO_DIARIO                                  |
'|          |      |   t_SESSAO_ABANDONADA                                     |
'|          |      |   t_SESSAO_HISTORICO                                      |
'|          |      |   t_SESSAO_RESTAURADA                                     |
'|          |      |Esta alteração foi feita porque foi criado um serviço para |
'|          |      |ser instalado no servidor que será responsável por executar|
'|          |      |tarefas dessa natureza automaticamente.                    |
'|          |      |Além disso, também foram retiradas as operações de backup  |
'|          |      |e restauração do banco de dados porque estavam obsoletas   |
'|          |      |no atual contexto: servidor próprio, tamanho grande do BD, |
'|          |      |esquema de backup automático no servidor, etc.             |
'|          |      |Outra função retirada é a de limpeza dos dados antigos de  |
'|          |      |pedidos, orçamentos, estoque, etc porque há interesse em   |
'|          |      |se manter todo o histórico dos pedidos e as rotinas estavam|
'|          |      |defasadas com relação às evoluções ocorridas no sistema.   |
'|          |      |Portanto, para evitar acionamentos acidentais de tais ope- |
'|          |      |rações e para deixar o projeto deste módulo mais "limpo",  |
'|          |      |as funções desnecessárias foram removidas.                 |
'|__________|______|___________________________________________________________|
'|08.09.2011| HHO  |V1.19 Ajuste na rotina de transferência de produtos devido |
'|          |      |aos novos campos criados na tabela t_PRODUTO em decorrência|
'|          |      |da melhoria no processo de separação de mercadorias no es- |
'|          |      |toque. Um desses novos campos (deposito_zona_id) está defi-|
'|          |      |nido como "NOT NULL" e possui uma "CONSTRAINT DEFAULT(0)". |
'|          |      |Como a "CONSTRAINT" não é criada automaticamente ao criar a|
'|          |      |tabela temporária, ao criar o registro de um novo produto  |
'|          |      |estava ocorrendo o seguinte erro: "-2147217873: Cannot     |
'|          |      |insert the value NULL into column 'deposito_zona_id', table|
'|          |      |'artven2Lab.artven2.tmpAdm__t_PRODUTO'; column does not    |
'|          |      |allow nulls. INSERT fails."                                |
'|          |      |Para solucionar isso, está sendo criada a "CONSTRAINT      |
'|          |      |DEFAULT(0)" na tabela temporária.                          |
'|__________|______|___________________________________________________________|
'|15.01.2013| HHO  |V1.20 Ajuste na rotina de transferência de produtos devido |
'|          |      |aos novos campos criados na tabela t_PRODUTO em decorrência|
'|          |      |do relatório Farol Resumido.                               |
'|          |      |Um desses novos campos (farol_qtde_comprada) está definido |
'|          |      |como "NOT NULL"  e possui uma "CONSTRAINT DEFAULT(0)".     |
'|          |      |Como a "CONSTRAINT" não é criada automaticamente ao criar a|
'|          |      |tabela temporária, ao criar o registro de um novo produto  |
'|          |      |estava ocorrendo um erro.                                  |
'|          |      |Para solucionar isso, está sendo criada a "CONSTRAINT      |
'|          |      |DEFAULT(0)" na tabela temporária.                          |
'|__________|______|___________________________________________________________|
'|11.02.2013| HHO  |V1.21 Inclusão de tratamento para uma nova coluna na plani-|
'|          |      |lha de produtos referente ao novo campo "descontinuado" da |
'|          |      |tabela t_PRODUTO                                           |
'|__________|______|___________________________________________________________|
'|26.11.2014| HHO  |V1.22 Inclusão de novas colunas:                           |
'|          |      |   1) Potência (Btu/h)                                     |
'|          |      |   2) Ciclo (Frio/Quente Frio)                             |
'|          |      |   3) Posição Mercado (Básico/Premium)                     |
'|          |      | Além disso, várias colunas foram movidas de posição p/    |
'|          |      | melhorar o layout da planilha.                            |
'|__________|______|___________________________________________________________|
'|05.05.2016|LHGX  |Alteração para funcionamento em diversos ambientes         |
'|          |      |(entrada da DIS)                                           |
'|__________|______|___________________________________________________________|
'|21.09.2016| HHO  |V1.24 Alteração da rotina de carga da planilha de produtos |
'|          |      |para aceitar um tamanho maior no campo de descrição do     |
'|          |      |produto (de 40 para 120 caracteres).                       |
'|__________|______|___________________________________________________________|
'|19.10.2016| HHO  |V1.25 Alteração da rotina de carga da planilha de produtos |
'|          |      |para tratar o valor -1,00 na tabela de preços das lojas,   |
'|          |      |pois esse valor significa que o preço não deve ser atuali- |
'|          |      |zado.                                                      |
'|__________|______|___________________________________________________________|
'|06.03.2017| HHO  |V1.26 Alteração da rotina de carga da planilha de produtos |
'|          |      |para excluir o vínculo entre o produto e a regra de múlti- |
'|          |      |plos CD's para os produtos excluídos.                      |
'|__________|______|___________________________________________________________|
'|20.07.2018| HHO  |V1.27 Ajustes na rotina de carga da planilha de produtos   |
'|          |      |p/ permitir código EAN em duplicidade, desde que sejam     |
'|          |      |produtos do mesmo fabricante. A justificativa está no      |
'|          |      |uso de diferentes códigos internos de produto para um mes- |
'|          |      |mo produto quando se trata de um item usado em equipamen-  |
'|          |      |tos diferentes. A diferenciação através de diferentes có-  |
'|          |      |digos internos facilita a administração do estoque.        |
'|          |      |Inclusão de consistência para o tamanho do código NCM.     |
'|__________|______|___________________________________________________________|
'|08.08.2018| HHO  |V1.27(B) Alteração da extensão aceita para arquivos de     |
'|          |      |Excel de 'XLS' para 'XLSX'.                                |
'|__________|______|___________________________________________________________|
'|10.10.2018| HHO  |V1.28 Ajustes na rotina de carga da planilha de produtos   |
'|          |      |para aceitar códigos com tamanhos de 8, 12, 13 e 14 carac- |
'|          |      |teres no campo EAN (GTIN-8, GTIN-12, GTIN-13 ou GTIN-14,   |
'|          |      |antigos códigos EAN, UPC e DUN-14).                        |
'|__________|______|___________________________________________________________|
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'|          |      |                                                           |
'|__________|______|___________________________________________________________|
'


Global Const m_id_versao = "1.28"
Global Const m_id = "Módulo de Administração e Manutenção - v" & m_id_versao & " - 10.OUT.2018"




' PERÍODO MÍNIMO DE PERMANÊNCIA DOS REGISTROS NO BANCO DE DADOS
Global Const CORTE_PEDIDO_EM_DIAS = (20 * 366)
Global Const CORTE_ESTOQUE_EM_DIAS = (20 * 366)
Global Const CORTE_ORCAMENTO_EM_DIAS = (20 * 366)
Global Const CORTE_SENHA_DESCONTO_EM_DIAS = (20 * 366)


Global Const CORTE_PEDIDO_EM_REGISTROS = 5000
Global Const CORTE_ESTOQUE_EM_REGISTROS = 5000
Global Const CORTE_ORCAMENTO_EM_REGISTROS = 5000

Global Const PERIODO_LIMPEZA_BD_EM_DIAS = 60
Global Const PERIODO_BACKUP_BD_EM_DIAS = 3


' PASTA PARA GRAVAR O BACKUP DO BANCO DE DADOS
Global Const PASTA_BACKUP_BD = "BACKUP_BANCO_DADOS"
Global Const EXTENSAO_ARQ_DADO_BACKUP_BD = "DAT"
Global LISTA_TABELAS_EXCLUIDAS_DO_BACKUP As String
Global Const LISTA_TABELAS_EXCLUIDAS_DO_BACKUP_PRODUCAO = "|DTPROPERTIES|T_GPR_VOTACAO|T_GPR_VOTACAO_BCP|T_LOG|T_PRODUTO_LOJA|"
Global Const LISTA_TABELAS_EXCLUIDAS_DO_BACKUP_DESENVOLVIMENTO = "|DTPROPERTIES|T_GPR_VOTACAO|T_GPR_VOTACAO_BCP|T_LOG|"


' CARACTERES SEPARADORES
Global Const COD_SEPARADOR_COLUNA = "¦"
Global Const COD_SUBSTITUICAO_SEPARADOR_COLUNA = "|"

Global Const COD_SEPARADOR_REGISTRO = "¬"
Global Const COD_SUBSTITUICAO_SEPARADOR_REGISTRO = "-"

Global Const COD_VALOR_NULL = "°"
Global Const COD_SUBSTITUICAO_VALOR_NULL = "º"



' CONTROLE DE ACESSO
Global Const OP_CEN_MODULO_ADM = 17300


' CÓDIGOS PARA REGISTRO DE OPERAÇÕES NO LOG
Global Const OP_LOG_CARREGA_TABELA_PRODUTOS = "CARREGA PRODUTOS"
Global Const OP_LOG_ELIMINA_PEDIDO_ANTIGO = "APAGA PEDIDO ANTIGO"
Global Const OP_LOG_ELIMINA_ESTOQUE_ANTIGO = "APAGA ESTOQUE ANTIGO"
Global Const OP_LOG_ELIMINA_ORCAMENTO_ANTIGO = "APAGA ORÇAMTO ANTIGO"
Global Const OP_LOG_ELIMINA_SENHA_DESCONTO_ANTIGA = "APAGA AUT DESC ANTIG"
Global Const OP_LOG_BACKUP_BD = "BACKUP BD"
Global Const OP_LOG_RESTAURA_BACKUP_BD = "RESTAURA BACKUP BD"


Global Const TAM_MIN_NUM_PEDIDO = 6     ' SOMENTE PARTE NUMÉRICA DO NÚMERO DO PEDIDO
Global Const TAM_MIN_ID_PEDIDO = 7      ' PARTE NUMÉRICA DO NÚMERO DO PEDIDO + LETRA REFERENTE AO ANO
Global Const TAM_MIN_FABRICANTE = 3
Global Const TAM_MAX_FABRICANTE = 4
Global Const TAM_MIN_PRODUTO = 6
Global Const TAM_MAX_PRODUTO = 8
Global Const TAM_MIN_LOJA = 2
Global Const TAM_MAX_LOJA = 3
Global Const TAM_MIN_GRUPO_LOJAS = 2
Global Const TAM_MAX_GRUPO_LOJAS = 3

Global Const REG_CHAVE_ART3 = "SOFTWARE\PRAGMATICA\Artven3"
Global Const REG_CHAVE_ADM_OPCOES = REG_CHAVE_ART3 & "\ADM\Opcoes"
Global Const REG_CHAVE_ADM_TABELA_PRODUTOS = REG_CHAVE_ADM_OPCOES & "\Tabela Produtos"
Global Const REG_CAMPO_ADM_TABELA_PRODUTOS_PLANILHA = "Planilha"

Global Const COR_TAB_FUNDO = &HDFF9F9
Global Const COR_TAB_ATIVO = &H8ACAC7
Global Const COR_TAB_INATIVO = &H39807C
Global Const COR_TEXTO_TAB_INATIVO = &H0&
Global Const COR_TEXTO_TAB_ATIVO = &HF40000
Global Const COR_SUBLINHADO_TAB = &H6139FD
Global Const COR_BARRA_PROGRESSO = &H40C840
Global Const COR_BARRA_PROGRESSO_FUNDO = &HE0E0E0
Global Const COR_MENU_ATIVO = &HC00000
Global Const COR_SOMBRA_PAINEL = &HE0E0E0
Global Const COR_FRENTE_PAINEL = &HDFF9F9
Global Const COR_BORDA_PAINEL = &H0&


'   LOJA OU GRUPO DE LOJAS
Global Const PREFIXO_NUMERO_LOJA = "L"
Global Const PREFIXO_NUMERO_GRUPO_LOJAS = "G"


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
        
'   OPERAÇÕES (MOVIMENTOS) DO ESTOQUE
Global Const OP_ESTOQUE_ENTRADA = "CAD"
Global Const OP_ESTOQUE_VENDA = "VDA"
Global Const OP_ESTOQUE_CONVERSAO_KIT = "KIT"
Global Const OP_ESTOQUE_TRANSFERENCIA = "TRF"
Global Const OP_ESTOQUE_ENTREGA = "ETG"
Global Const OP_ESTOQUE_DEVOLUCAO = "DEV"
    
'   STATUS DE ENTREGA DO PEDIDO
Global Const ST_ENTREGA_ESPERAR = "ESP"         ' NENHUMA MERCADORIA SOLICITADA ESTÁ DISPONÍVEL
Global Const ST_ENTREGA_SPLIT_POSSIVEL = "SPL"  ' PARTE DA MERCADORIA ESTÁ DISPONÍVEL PARA ENTREGA
Global Const ST_ENTREGA_SEPARAR = "SEP"         ' TODA A MERCADORIA ESTÁ DISPONÍVEL E JÁ PODE SER SEPARADA PARA ENTREGA
Global Const ST_ENTREGA_A_ENTREGAR = "AET"      ' A TRANSPORTADORA JÁ SEPAROU A MERCADORIA PARA ENTREGA
Global Const ST_ENTREGA_ENTREGUE = "ETG"        ' MERCADORIA FOI ENTREGUE
Global Const ST_ENTREGA_CANCELADO = "CAN"       ' VENDA FOI CANCELADA
    
'   STATUS DO ORÇAMENTO
Global Const ST_ORCAMENTO_CANCELADO = "CAN"           ' ORÇAMENTO FOI CANCELADO
    
'   STATUS "RECEBIDO" DO PEDIDO (EXISTE APENAS PARA SATISFAZER AO CLIENTE QUANDO O PEDIDO É IMPRESSO)
Global Const ST_RECEBIDO_SIM = "S"
Global Const ST_RECEBIDO_NAO = "N"
Global Const ST_RECEBIDO_PARCIAL = "P"
    
'   STATUS DE PAGAMENTO DO PEDIDO (CONTROLA DE FATO O ANDAMENTO DOS PAGAMENTOS)
Global Const ST_PAGTO_PAGO = "S"
Global Const ST_PAGTO_NAO_PAGO = "N"
Global Const ST_PAGTO_PARCIAL = "P"
    
' CÓDIGOS PARA OPERAÇÕES
Global Const OP_CONSULTA = "C"
Global Const OP_INCLUI = "I"
Global Const OP_EXCLUI = "E"
Global Const OP_ALTERA = "A"
    
' CÓDIGOS PARA NÍVEL DOS USUÁRIOS
Global Const ID_VENDEDOR = "V"
Global Const ID_SEPARADOR = "S"
Global Const ID_ADMINISTRADOR = "A"
Global Const ID_GERENCIAL = "G"

' CÓDIGOS QUE IDENTIFICAM SE É PESSOA FÍSICA OU JURÍDICA
Global Const ID_PF = "PF"
Global Const ID_PJ = "PJ"

' CONSTANTES QUE IDENTIFICAM O NSU NA T_CONTROLE (MÁXIMO 20 CARACTERES)
Global Const NSU_QUADRO_AVISO = "QUADRO_DE_AVISOS"
Global Const NSU_CADASTRO_CLIENTES = "CADASTRO_CLIENTES"
Global Const NSU_PEDIDO = "PEDIDO"
Global Const NSU_PEDIDO_TEMPORARIO = "PEDIDO_TEMPORARIO"
Global Const NSU_ID_ESTOQUE_MOVTO = "ESTOQUE_MOVTO"
Global Const NSU_ID_ESTOQUE = "ESTOQUE"
Global Const NSU_ID_ESTOQUE_TEMP = "ESTOQUE_TEMPORARIO"
Global Const NSU_PEDIDO_PAGAMENTO = "PEDIDO_PAGAMENTO"
Global Const NSU_DESC_SUP_AUTORIZACAO = "DESC_SUP_AUTORIZACAO"
Global Const NSU_PEDIDO_ITEM_DEVOLVIDO = "PEDIDO_ITEM_DEVOLVID"
Global Const NSU_ULTIMA_LIMPEZA_BD = "ULTIMA_LIMPEZA_BD"
Global Const NSU_ULTIMO_BACKUP_BD = "ULTIMO_BACKUP_BD"




Type TIPO_LISTA_CFOP
    codigo As String
    descricao As String
    End Type


'------------------------------------------------------------------------------------
'   PAINEL P/ SELECIONAR DIRETÓRIO
'------------------------------------------------------------------------------------
    Type TIPO_OPCAO_SELECIONA_DIRETORIO
        titulo_principal As String
        titulo_secundario As String
        diretorio_selecionado As String
        diretorio_inicial_default As String
        cancelou_operacao As Boolean
        End Type

    Global opcao_seleciona_diretorio As TIPO_OPCAO_SELECIONA_DIRETORIO
    

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
    
  ' DB-LIB
  ' ~~~~~~
    s_chave = "Software\Microsoft\MSSQLServer\Client\DB-Lib"
    s_campo = "UseIntlSettings"
    s_valor = "off"
    If Not registry_grava_string(s_chave, s_campo, s_valor, msg_erro) Then Exit Function
    
  ' SuperSocketNetLib
  ' ~~~~~~~~~~~~~~~~~
    s_chave = "Software\Microsoft\MSSQLServer\Client\SuperSocketNetLib"
    s_campo = "ProtocolOrder"
    s_valor = "7463700000"
    If Not registry_grava_binario(s_chave, s_campo, s_valor, msg_erro) Then Exit Function
    
    s_campo = "Encrypt"
    n_valor = 0
    If Not registry_grava_numero(s_chave, s_campo, n_valor, msg_erro) Then Exit Function
    
  ' T C P
  ' ~~~~~
    s_chave = "Software\Microsoft\MSSQLServer\Client\SuperSocketNetLib\Tcp"
    s_campo = "DefaultPort"
    n_valor = 1433
    If Not registry_grava_numero(s_chave, s_campo, n_valor, msg_erro) Then Exit Function
    
  ' ConnectTo
  ' ~~~~~~~~~
    s_chave = "SOFTWARE\Microsoft\MSSQLServer\Client\ConnectTo"
    s_campo = "DSQUERY"
    s_valor = "DBMSSOCN"
    
  ' PARA CLIENT DO SQL SERVER 2000: NÃO ALTERA A CONFIGURAÇÃO DA DLL USADA PARA TCP/IP
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


Function IsTabelaBasicaProdutos(ByVal codigo As String) As Boolean
' _____________________________________________________________________________________
'|
'|  IDENTIFICA SE O CÓDIGO ESPECIFICADO É DA PLANILHA
'|  QUE CONTÉM A TABELA BÁSICA DE PRODUTOS
'|
'|  SUBSÍDIOS:
'|  ==========
'|      TABELA BÁSICA: NOME DA PLANILHA IGUAL A "00" (OU "0", OU "000")
'|      TABELA DE PRODUTOS DE UMA LOJA: "01" OU "001" A "999"
'|                                      "L01" OU "L001" A "L999"
'|      TABELA DE PRODUTOS DE UM GRUPO DE LOJAS: "G01" A "G99"
'|

    IsTabelaBasicaProdutos = False
    
    codigo = Trim$("" & codigo)
    
    If codigo = "" Then Exit Function
    If codigo <> String$(Len(codigo), "0") Then Exit Function
    
    IsTabelaBasicaProdutos = True
    
End Function

Function separa_campo(ByVal Texto As String, ByVal separador_campo As String) As String
' ___________________________________________________________________________________________________________
'|
'|  RETORNA A PARTE DO TEXTO QUE ESTÁ ANTES DO CARACTER SEPARADOR.
'|

Dim i As Long
Dim c As String
Dim s_resp As String
    
    separa_campo = ""
    
    s_resp = ""
    For i = 1 To Len(Texto)
        c = Mid$(Texto, i, 1)
        If c = separador_campo Then Exit For
        s_resp = s_resp & c
        Next
        
    separa_campo = s_resp
    
End Function

Function IsNumeroLoja(ByVal numero As String) As Boolean
' __________________________________________________________________________________________________
'|
'|  IDENTIFICA SE É UM NÚMERO DE LOJA
'|

Dim s_numero As String
Dim s_prefixo As String

    On Error GoTo INL_TRATA_ERRO
    
    IsNumeroLoja = False
    
    numero = UCase$(Trim$("" & numero))
    
    s_prefixo = ""
    If Left$(numero, Len(PREFIXO_NUMERO_LOJA)) = PREFIXO_NUMERO_LOJA Then s_prefixo = Left$(numero, Len(PREFIXO_NUMERO_LOJA))
    
    s_numero = Mid$(numero, Len(s_prefixo) + 1)
    
    If Not IsNumeric(s_numero) Then Exit Function
    If CLng(s_numero) <= 0 Then Exit Function
    If Len(s_numero) > TAM_MAX_LOJA Then Exit Function
        
    IsNumeroLoja = True
    
Exit Function




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
INL_TRATA_ERRO:
'==============
    Err.Clear
    Exit Function
    
End Function

Function IsNumeroGrupodeLojas(ByVal numero As String) As Boolean
' __________________________________________________________________________________________________
'|
'|  IDENTIFICA SE É UM NÚMERO DE GRUPO DE LOJAS
'|

Dim s_numero As String
Dim s_prefixo As String

    On Error GoTo INGL_TRATA_ERRO
    
    IsNumeroGrupodeLojas = False
    
    numero = UCase$(Trim$("" & numero))
    
    s_prefixo = ""
    If Left$(numero, Len(PREFIXO_NUMERO_GRUPO_LOJAS)) = PREFIXO_NUMERO_GRUPO_LOJAS Then s_prefixo = Left$(numero, Len(PREFIXO_NUMERO_GRUPO_LOJAS))
    
    s_numero = Mid$(numero, Len(s_prefixo) + 1)
    
    If s_prefixo = "" Then Exit Function
    
    If Not IsNumeric(s_numero) Then Exit Function
    If CLng(s_numero) <= 0 Then Exit Function
    If Len(s_numero) > TAM_MAX_GRUPO_LOJAS Then Exit Function
        
    IsNumeroGrupodeLojas = True
    
Exit Function




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
INGL_TRATA_ERRO:
'===============
    Err.Clear
    Exit Function
    
End Function


Function le_arquivo_ini(ByRef msg_erro As String) As Boolean
' ------------------------------------------------------------------------
'   LÊ ARQUIVO DE CONFIGURAÇÃO

Dim s_arq As String
Dim s_linha As String
Dim s_param As String
Dim s_valor As String
Dim s_senha As String
Dim v() As String

' ARQUIVO-TEXTO
Dim Fnum As Integer

    On Error GoTo LAI_TRATA_ERRO
    
    le_arquivo_ini = False
    msg_erro = ""
    
    s_arq = barra_invertida_add(App.Path) & ExtractFileName(App.EXEName, True) & ".INI"
    If Not FileExists(s_arq, msg_erro) Then
        If msg_erro = "" Then msg_erro = "NÃO foi encontrado o arquivo " & s_arq
        Exit Function
        End If

    Fnum = FreeFile
    Open s_arq For Input As Fnum
        
    On Error GoTo LAI_TRATA_ERRO_ARQUIVO
        
    s_senha = ""
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
                Case "SERVIDOR_BD"
                    bd_selecionado.NOME_SERVIDOR = s_valor
                Case "NOME_BD"
                    bd_selecionado.NOME_BD = s_valor
                Case "NOME_USUARIO_BD"
                    bd_selecionado.ID_USUARIO = s_valor
                Case "SENHA_USUARIO_BD"
                    s_senha = s_valor
                End Select
            End If
        Loop
        
    Close Fnum
        
    If Not decodifica_dado(s_senha, bd_selecionado.SENHA_USUARIO) Then
        msg_erro = "Senha inválida !!"
        Exit Function
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


Function normaliza_codigo(ByRef codigo As String, ByVal tamanho_default As Long)
' ------------------------------------------------------------------------
'   NORMALIZA_CODIGO
Dim s As String
    normaliza_codigo = ""
    s = Trim$("" & codigo)
    If s = "" Then Exit Function
    Do While Len(s) < tamanho_default: s = "0" & s: Loop
    normaliza_codigo = s
End Function

Function remove_prefixo_do_numero(ByVal numero As String, ByVal prefixo As String) As String
Dim s_resp As String
    
    numero = Trim$("" & numero)
    prefixo = Trim$("" & prefixo)
        
    If Left$(numero, Len(prefixo)) = prefixo Then
        s_resp = Mid$(numero, Len(prefixo) + 1)
    Else
        s_resp = numero
        End If
        
    remove_prefixo_do_numero = s_resp
    
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


