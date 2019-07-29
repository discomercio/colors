VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form f_MAIN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F8F8FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Administração e Manutenção"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "f_MAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton b_dummy 
      Appearance      =   0  'Flat
      Caption         =   "b_dummy"
      Height          =   345
      Left            =   6000
      TabIndex        =   4
      Top             =   2775
      Width           =   1350
   End
   Begin VB.Timer relogio 
      Interval        =   1000
      Left            =   6840
      Top             =   3210
   End
   Begin VB.PictureBox base_status 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2100
      Width           =   7335
      Begin VB.Label info 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "info"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   2415
         TabIndex        =   3
         Top             =   75
         Width           =   870
      End
      Begin VB.Label hoje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "00.MMM.0000"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   75
         Width           =   1080
      End
      Begin VB.Label agora 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "00:00"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   75
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog fcmd 
      Left            =   6780
      Top             =   690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "c:\"
   End
   Begin VB.Label lblTitAmb 
      AutoSize        =   -1  'True
      BackColor       =   &H00F8F8FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ambiente: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1455
      TabIndex        =   7
      Top             =   840
      Width           =   1110
   End
   Begin VB.Label lblTitSite 
      AutoSize        =   -1  'True
      BackColor       =   &H00F8F8FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Site: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1455
      TabIndex        =   6
      Top             =   480
      Width           =   540
   End
   Begin VB.Label barra 
      Caption         =   "    Carregar Tabela de Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1455
      TabIndex        =   5
      Top             =   1320
      Width           =   4500
   End
   Begin VB.Image seta 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   240
      Index           =   0
      Left            =   315
      Picture         =   "f_MAIN.frx":030A
      Top             =   495
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "f_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************************************************************************
'    =====================================
'    A U T O M A Ç Ã O    D O    E X C E L
'    =====================================
' 
' 1) DEVE SER ADICIONADA A SEGUINTE REFERÊNCIA NO PROJETO
'    [Project] -> [References...] -> [Microsoft Excel 8.0 Object Library]

' 
' 
' 2) DEVE-SE EVITAR A INSERÇÃO DE FIGURAS NA PLANILHA DEVIDO À LENTIDÃO DESTA
'    OPERAÇÃO.  UM BITMAP DE 175 KB LOCALIZADO NA UNIDADE LOCAL DEMORA CERCA DE
'    35s PARA SER INSERIDO.
' 
' 
' 
' 3) A CONFIGURAÇÃO DA FONTE PADRÃO NÃO É RECOMENDÁVEL PELOS SEGUINTES MOTIVOS:
'    a) AS ALTERAÇÕES SÓ TEM EFEITO NA PRÓXIMA EXECUÇÃO DO EXCEL.
'    b) AS ALTERAÇÕES FEITAS VIA AUTOMAÇÃO FAZEM COM QUE AS CONFIGURAÇÕES
'       ORIGINAIS DO USUÁRIO SEJAM PERDIDAS.
' 
'    UMA ALTERNATIVA É A UTILIZAÇÃO DO ESTILO "NORMAL" (PROPRIEDADE STYLE), QUE
'    PODE SER CONFIGURADO DA MANEIRA NECESSÁRIA, COM A VANTAGEM DE QUE ESSAS
'    CONFIGURAÇÕES SÃO GRAVADAS JUNTO COM O WORKBOOK.
' 
' 
' 
' 4) AO PROGRAMAR ROTINAS DE AUTOMAÇÃO DO EXCEL, DEVE-SE TER SEMPRE EM MENTE:
'    a) FATO: O EXCEL NÃO É FECHADO ENQUANTO HOUVER REFERÊNCIAS SENDO FEITAS
'       A ELE (VEJA MSDN Q187745).  NESTA SITUAÇÃO, QUANDO É ACIONADO O COMANDO
'       PARA FECHAR, ELE SIMPLESMENTE FICA INVISÍVEL, MAS AINDA CONTINUA
'       CARREGADO NA MEMÓRIA E ATIVO.
' 
'    b) FATO: PARA FECHAR UM OBJETO EXCEL CRIADO PARA AUTOMAÇÃO, DEVEM SER
'       EXECUTADOS OS SEGUINTES COMANDOS:
'           b1) oEXCEL.Quit
'           b2) Set oEXCEL = Nothing
'       APÓS A EXECUÇÃO DO COMANDO "Set oEXCEL = Nothing", O EXCEL DEVE FECHAR
'       E SUMIR DA LISTA DE PROCESSOS.  CASO ISSO NÃO OCORRA, ENTÃO É PORQUE
'       AINDA EXISTEM REFERÊNCIAS PENDENTES SENDO FEITAS A ELE.  NESTE CASO,
'       O EXCEL SOMENTE IRÁ SE FECHAR APÓS AS REFERÊNCIAS SEREM ELIMINADAS.
'       SE O OBJETO EXCEL FOI CRIADO E MANIPULADO DENTRO DE UMA ROTINA, ENTÃO
'       AS REFERÊNCIAS SERÃO ELIMINADAS QUANDO A ROTINA TERMINAR.  OU SEJA,
'       NO CASO DE HAVER REFERÊNCIAS PENDENTES, O EXCEL SERÁ FECHADO SOMENTE
'       APÓS A EXECUÇÃO DA INSTRUÇÃO "EXIT SUB/EXIT FUNCTION".
'       SEMPRE QUE UMA ROTINA FOR IMPLEMENTADA OU ALTERADA, DEVE-SE VERIFICAR
'       SE ESTÃO FICANDO REFERÊNCIAS PENDENTES, POIS ISSO PODE GERAR O
'       SEGUINTE ERRO LOGO APÓS A EXECUÇÃO DA INSTRUÇÃO "EXIT SUB/EXIT FUNCTION":
'       "EXCEL causou uma falha de página inválida no módulo OLE32.DLL em 0187:7ff38ae5."
' 
'    c) UMA DAS CAUSAS DAS REFERÊNCIAS AO EXCEL QUE FICAM PENDENTES É O USO
'       DO COMANDO "WITH" DO VISUAL BASIC EM UMA DETERMINADA SITUAÇÃO.
'       VEJA OS EXEMPLOS A SEGUIR:
' 
'       EXEMPLO 1:
'       With oXL.Workbooks(1).Worksheets(1)
'           .Range("A1:C1").Value = Array("X1", "X2", "X3")
'           End With
' 
'       EXEMPLO 2:
'       Set oRANGE = oWB.Worksheets(1).Range("A1:D3")
'       With oRANGE
'           With .Font
'               .SIZE = 9
'               End With
'           End With
' 
'       EXPLICAÇÕES:
'       ============
'       NO EXEMPLO 1, A CHAMADA AO RANGE("A1:C1") RETORNA UM OBJETO RANGE QUE FICA
'       PENDENTE DEVIDO AO USO DO COMANDO "WITH".
'       NO EXEMPLO 2, A CHAMADA À PROPRIEDADE "FONT" RETORNA UM OBJETO "FONT" QUE
'       TAMBÉM FICA PENDENTE DEVIDO AO USO DO COMANDO "WITH".
' 
'       CORREÇÕES:
'       ==========
'       NO EXEMPLO 1, DEVE-SE USAR:
'           oWB.Worksheets(1).Range("A1:C1").Value = Array("X1", "X2", "X3")
'               OU ENTÃO
'           Set oRANGE = oWB.Worksheets(1).Range("A1:C1")
'           oRANGE.Value = Array("X1", "X2", "X3")
' 
'       NO EXEMPLO 2, DEVE-SE USAR:
'           Set oFONT = oWB.Worksheets(1).Range("A1:D3").Font
'           With oFONT
'               .SIZE = 9
'               End With
' 
'       TANTO OS OBJETOS oRANGE QUANTO oFONT DEVEM SER FINALIZADOS NO FINAL DA
'       DA ROTINA COM "Set oRANGE = Nothing" E "Set oFONT = Nothing", CASO
'       CONTRÁRIO FICARÃO COMO REFERÊNCIAS PENDENTES.
'       NOTE QUE ELAS PRECISAM SER FINALIZADAS COM "Nothing" SOMENTE NO FINAL
'       DA ROTINA, OU SEJA, DURANTE O PROCESSAMENTO ELAS PODEM RECEBER NOVAS
'       ATRIBUIÇÕES DE OBJETOS DIRETAMENTE.
' 
'       CONCLUSÃO:
'       ==========
'       NÃO SE DEVE USAR O COMANDO "WITH" COM NENHUMA PROPRIEDADE OU MÉTODO
'       QUE RETORNE UM OBJETO.
' 
' 
' 
' 5) O TAMANHO MÁXIMO PARA O CONTEÚDO DE CADA CÉLULA É DE 255 CARACTERES.
'    SE O CONTEÚDO DA CÉLULA EXCEDER ESTE VALOR, IRÃO OCORRER PROBLEMAS
'    NA EXIBIÇÃO DOS DADOS, POIS A CÉLULA SERÁ EXIBIDA COMO "#########".
'    O CONTEÚDO SERÁ ARMAZENADO CORRETAMENTE, SENDO QUE O PROBLEMA ESTÁ
'    APENAS NA EXIBIÇÃO (TELA E IMPRESSORA).
' 
' 
' 6) PROPRIEDADE "PAGESETUP":
'    A CONFIGURAÇÃO DE QUALQUER PROPRIEDADE DE "PAGESETUP" CAUSA UMA EXCESSIVA
'    LENTIDÃO NA TRANSFERÊNCIA DE DADOS E TORNA O EXCEL INSTÁVEL.
'    A LENTIDÃO PODE SER MAIS DE 10 VEZES O TEMPO CONSUMIDO EM UM PROCES-
'    SAMENTO NORMAL.
'    A INSTABILIDADE É NOTADA PRINCIPALMENTE QUANDO SE GERA PLANILHAS COM
'    MUITAS LINHAS OU QUE TENHAS CÉLULAS COM CONTEÚDO EXTENSO (EM TORNO
'    DE 200 CARACTERES).  A INSTABILIDADE SE TRADUZ EM ERROS DE "INVALID
'    PAGE FAULT" DO EXCEL.
'    AS RECOMENDAÇÕES DO ARTIGO MSDN Q145598 NÃO SURTIRAM NENHUM EFEITO.
'    ALÉM DISSO, A PROPRIEDADE "DisplayPageBreaks" DA PLANILHA DEVE SER
'    MANTIDA COMO "FALSE" EM VIRTUDE DOS MESMOS PROBLEMAS DO "PAGESETUP".
'    CONCLUSÃO: ENQUANTO NÃO SE ENCONTRA UMA SOLUÇÃO EFICAZ, DEVE-SE
'    EVITAR O USO DESSAS PROPRIEDADES.
' 
' 
' 
' 
' 
' 
' 
' 


Dim modulo_executou_inicializacao As Boolean


' CORES DO PAINEL
  Const COR_BARRA = &HE8C8BF
  Const FATOR_COR_SOMBREAMENTO = 25  ' EM PERCENTUAL

' BARRA: OPERAÇÕES
  Const o_CARREGA_PRODUTOS = 0
  
  

Private Type TIPO_T_PRODUTO
    fabricante As String
    produto As String
    descricao As String
    descricao_html As String
    ean As String
    grupo As String
    preco_fabricante As Currency
    vl_custo2 As Currency
    estoque_critico As Long
    peso As Double
    qtde_volumes As Long
    alertas As String
    cubagem As Double
    ncm As String
    cst As String
    perc_MVA_ST As Double
    descontinuado As String
    potencia_BTU As Long
    ciclo As String
    posicao_mercado As String
    dt_cadastro As Variant
    dt_ult_atualizacao As Variant
    End Type

Private Type TIPO_T_PRODUTO_LOJA
    fabricante As String
    produto As String
    preco_lista As Currency
    margem As Double
    desc_max As Double
    comissao As Double
    vendavel As String
    qtde_max_venda As Long
    cor As String
    dt_cadastro As Variant
    dt_ult_atualizacao As Variant
    End Type

Private Function elimina_pedido_antigo(ByRef msg_erro As String) As Boolean

Dim s As String
Dim s_log As String
Dim hora_inicio As Date
Dim hora_termino As Date
Dim dt_servidor As Date
Dim dt_corte As Date
Dim s_sql As String
Dim t As ADODB.Recordset
Dim sx As ADODB.Recordset
Dim n_reg As Long
Dim n_contador As Long
Dim pagamento_quitado As Boolean
Dim todos_cancelados As Boolean
Dim pode_apagar As Boolean
Dim dt_ult_entrega As Date
Dim dt_ult_cancelamento As Date
Dim n_t_pedido_pagto_visanet As Long
Dim n_t_pedido_pagamento As Long
Dim n_t_pedido_item_devolvido As Long
Dim n_t_pedido_item As Long
Dim n_t_pedido As Long

    On Error GoTo EPA_TRATA_ERRO

    elimina_pedido_antigo = False
    msg_erro = ""
    
    hora_inicio = Now
    
    aguarde INFO_EXECUTANDO, "eliminando registros de pedidos antigos"
    
    If Not obtem_data_servidor(dt_servidor, msg_erro) Then
        s = "Erro ao consultar a data/hora do servidor!!"
        If msg_erro <> "" Then s = s & Chr(13) & Chr(13) & msg_erro
        aviso_erro s
        aguarde INFO_NORMAL
        Exit Function
        End If
        
    dt_corte = dt_servidor - CORTE_PEDIDO_EM_DIAS
    
  ' RECORDSET
    Set t = New ADODB.Recordset
    t.CursorType = BD_CURSOR_SOMENTE_LEITURA
    t.LockType = BD_POLITICA_LOCKING
    t.CacheSize = BD_CACHE_CONSULTA
    
    Set sx = New ADODB.Recordset
    sx.CursorType = BD_CURSOR_SOMENTE_LEITURA
    sx.LockType = BD_POLITICA_LOCKING
    sx.CacheSize = BD_CACHE_CONSULTA
    
    
  ' EXECUTA CORTE NO BD POR DATA SOMENTE SE HOUVER MAIS REGISTROS QUE O LIMITE MÍNIMO
    s_sql = "SELECT COUNT(*) AS total FROM t_PEDIDO WHERE (data >= " & bd_monta_data(dt_corte, False) & ")"
    t.Open s_sql, dbc, , , adCmdText
    n_reg = 0
    If Not t.EOF Then If IsNumeric(t("total")) Then n_reg = CLng(t("total"))
    If n_reg <= CORTE_PEDIDO_EM_REGISTROS Then
        elimina_pedido_antigo = True
        s_log = "Eliminação de dados antigos não foi feita porque restariam apenas " & _
                Format$(n_reg, FORMATO_NUMERO) & " registros em T_PEDIDO posteriores à data de corte " & _
                Format$(dt_corte, FORMATO_DATA) & " (limite mínimo: " & _
                Format$(CORTE_PEDIDO_EM_REGISTROS, FORMATO_NUMERO) & ")"
        Call grava_log(usuario.id, "", "", "", OP_LOG_ELIMINA_PEDIDO_ANTIGO, s_log, msg_erro)
        GoSub EPA_FECHA_TABELAS
        aguarde INFO_NORMAL
        Exit Function
        End If
            
            
  ' IMPORTANTE: COMO A TRANSAÇÃO PODE BLOQUEAR O ACESSO DE OUTROS USUÁRIOS AO BD, NUNCA
  ' =========== DEVE HAVER INTERAÇÃO COM O USUÁRIO NAS ROTINAS CONTIDAS NA TRANSAÇÃO !!!
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dbc.BeginTrans
    On Error GoTo EPA_TRATA_ERRO_TRANSACAO
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    n_t_pedido_pagto_visanet = 0
    n_t_pedido_pagamento = 0
    n_t_pedido_item_devolvido = 0
    n_t_pedido_item = 0
    n_t_pedido = 0
    n_contador = 0
    
  ' A ANÁLISE DEVE SER FEITA POR FAMÍLIA DE PEDIDOS, PORTANTO, SELECIONA OS PEDIDOS-BASE
    s_sql = "SELECT DISTINCT pedido FROM t_PEDIDO WHERE" & _
            " (" & bd_monta_len("pedido") & " = " & CStr(TAM_MIN_ID_PEDIDO) & ")" & _
            " AND (data < " & bd_monta_data(dt_corte, False) & ")"
    
    sx.Open s_sql, dbc, , , adCmdText
    Do While Not sx.EOF
        
        dt_ult_entrega = 0
        dt_ult_cancelamento = 0
        
        pagamento_quitado = False
        todos_cancelados = True
        
    '   ANALISA CADA UM DOS PEDIDOS QUE COMPÕEM A FAMÍLIA DE PEDIDOS
        s_sql = "SELECT * FROM t_PEDIDO WHERE" & _
                " (pedido LIKE '" & Trim$("" & sx("pedido")) & BD_CURINGA_TODOS & "')"
        If t.State <> adStateClosed Then t.Close
        t.Open s_sql, dbc, , , adCmdText
        Do While Not t.EOF
            n_contador = n_contador + 1
            DoEvents
            If (n_contador Mod 10) = 0 Then aguarde INFO_EXECUTANDO, "pedidos: " & Format$(n_contador, FORMATO_NUMERO) & " regs processados"
            
        '   HÁ PEDIDO NÃO ENCERRADO ?
            If (Trim$("" & t("st_entrega")) <> ST_ENTREGA_ENTREGUE) And _
               (Trim$("" & t("st_entrega")) <> ST_ENTREGA_CANCELADO) Then GoTo PROX_FAMILIA
                
        '   PAGAMENTO QUITADO ? (LEMBRANDO QUE O STATUS DE PAGAMENTO É ANOTADO SOMENTE NO PEDIDO-BASE)
            If Len(Trim$("" & t("pedido"))) = TAM_MIN_ID_PEDIDO Then
                If Trim$("" & t("st_pagto")) = ST_PAGTO_PAGO Then pagamento_quitado = True
                End If
                
        '   HÁ PEDIDO NÃO CANCELADO ?
            If Trim$("" & t("st_entrega")) <> ST_ENTREGA_CANCELADO Then todos_cancelados = False
            
        '   DATA DE CANCELAMENTO MAIS RECENTE
            If Trim$("" & t("st_entrega")) = ST_ENTREGA_CANCELADO Then
                If IsDate(t("cancelado_data")) Then If t("cancelado_data") > dt_ult_cancelamento Then dt_ult_cancelamento = t("cancelado_data")
                End If
                
        '   DATA DE ENTREGA MAIS RECENTE
            If Trim$("" & t("st_entrega")) = ST_ENTREGA_ENTREGUE Then
                If IsDate(t("entregue_data")) Then If t("entregue_data") > dt_ult_entrega Then dt_ult_entrega = t("entregue_data")
                End If
                            
            t.MoveNext
            Loop
        
    
    '   SOMENTE EXCLUI FAMÍLIA DE PEDIDOS SE ELA FOI TODA CANCELADA OU, SENÃO, SE ESTÁ C/ PAGAMENTO QUITADO
        pode_apagar = False
        If todos_cancelados Then
            pode_apagar = True
        Else
            If pagamento_quitado Then pode_apagar = True
            End If
            
        If Not pode_apagar Then GoTo PROX_FAMILIA

    '   ANALISA ÚLTIMAS DATAS DE ENTREGA E CANCELAMENTO DE PEDIDOS
        If dt_ult_entrega >= dt_corte Then GoTo PROX_FAMILIA
        If dt_ult_cancelamento >= dt_corte Then GoTo PROX_FAMILIA
        
        
    '   HOUVE ANOTAÇÃO DE ALGUMA PARCELA DE PAGAMENTO RECENTEMENTE ?
        s_sql = "SELECT COUNT(*) AS qtde FROM t_PEDIDO_PAGAMENTO WHERE" & _
                " (pedido LIKE '" & Trim$("" & sx("pedido")) & BD_CURINGA_TODOS & "')" & _
                " AND (data >= " & bd_monta_data(dt_corte, False) & ")"
        If t.State <> adStateClosed Then t.Close
        t.Open s_sql, dbc, , , adCmdText
        n_reg = 0
        If Not t.EOF Then
            If IsNumeric(t("qtde")) Then n_reg = CLng(t("qtde"))
            End If

        If n_reg > 0 Then GoTo PROX_FAMILIA
        
    '   HOUVE ALGUMA TRANSAÇÃO DE PAGAMENTO PELA VISANET RECENTEMENTE ?
        s_sql = "SELECT COUNT(*) AS qtde FROM t_PEDIDO_PAGTO_VISANET WHERE" & _
                " (pedido LIKE '" & Trim$("" & sx("pedido")) & BD_CURINGA_TODOS & "')" & _
                " AND (data >= " & bd_monta_data(dt_corte, False) & ")"
        If t.State <> adStateClosed Then t.Close
        t.Open s_sql, dbc, , , adCmdText
        n_reg = 0
        If Not t.EOF Then
            If IsNumeric(t("qtde")) Then n_reg = CLng(t("qtde"))
            End If

        If n_reg > 0 Then GoTo PROX_FAMILIA
        
    '   HOUVE DEVOLUÇÃO DE MERCADORIAS RECENTEMENTE ?
        s_sql = "SELECT COUNT(*) AS qtde FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE" & _
                " (pedido LIKE '" & Trim$("" & sx("pedido")) & BD_CURINGA_TODOS & "')" & _
                " AND (devolucao_data >= " & bd_monta_data(dt_corte, False) & ")"
        If t.State <> adStateClosed Then t.Close
        t.Open s_sql, dbc, , , adCmdText
        n_reg = 0
        If Not t.EOF Then
            If IsNumeric(t("qtde")) Then n_reg = CLng(t("qtde"))
            End If

        If n_reg > 0 Then GoTo PROX_FAMILIA
                

    '   APAGA A FAMÍLIA DE PEDIDOS !!
    '   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        s_sql = "DELETE FROM t_PEDIDO_PAGTO_VISANET WHERE" & _
                " (pedido LIKE '" & Trim$("" & sx("pedido")) & BD_CURINGA_TODOS & "')"
        dbc.Execute s_sql, n_reg
        If n_reg > 0 Then n_t_pedido_pagto_visanet = n_t_pedido_pagto_visanet + n_reg
        
        s_sql = "DELETE FROM t_PEDIDO_PAGAMENTO WHERE" & _
                " (pedido LIKE '" & Trim$("" & sx("pedido")) & BD_CURINGA_TODOS & "')"
        dbc.Execute s_sql, n_reg
        If n_reg > 0 Then n_t_pedido_pagamento = n_t_pedido_pagamento + n_reg
        
        s_sql = "DELETE FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE" & _
                " (pedido LIKE '" & Trim$("" & sx("pedido")) & BD_CURINGA_TODOS & "')"
        dbc.Execute s_sql, n_reg
        If n_reg > 0 Then n_t_pedido_item_devolvido = n_t_pedido_item_devolvido + n_reg
        
        s_sql = "DELETE FROM t_PEDIDO_ITEM WHERE" & _
                " (pedido LIKE '" & Trim$("" & sx("pedido")) & BD_CURINGA_TODOS & "')"
        dbc.Execute s_sql, n_reg
        If n_reg > 0 Then n_t_pedido_item = n_t_pedido_item + n_reg
                
        s_sql = "DELETE FROM t_PEDIDO WHERE" & _
                " (pedido LIKE '" & Trim$("" & sx("pedido")) & BD_CURINGA_TODOS & "')"
        dbc.Execute s_sql, n_reg
        If n_reg > 0 Then n_t_pedido = n_t_pedido + n_reg


PROX_FAMILIA:
'============

        sx.MoveNext
        Loop
        


'   LOG: QUANTIDADES EXCLUÍDAS
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "registros excluídos:" & _
            " T_PEDIDO_PAGTO_VISANET=" & Format$(n_t_pedido_pagto_visanet, FORMATO_NUMERO) & _
            " T_PEDIDO_PAGAMENTO=" & Format$(n_t_pedido_pagamento, FORMATO_NUMERO) & _
            " T_PEDIDO_ITEM_DEVOLVIDO=" & Format$(n_t_pedido_item_devolvido, FORMATO_NUMERO) & _
            " T_PEDIDO_ITEM=" & Format$(n_t_pedido_item, FORMATO_NUMERO) & _
            " T_PEDIDO=" & Format$(n_t_pedido, FORMATO_NUMERO)
            
'   DURAÇÃO
    hora_termino = Now
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "duração do processo: " & Format$(hora_termino - hora_inicio, "hh:mm:ss")
     
'   DATA DE CORTE
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "data de corte: " & Format$(dt_corte, FORMATO_DATA)
   
    If Not grava_log(usuario.id, "", "", "", OP_LOG_ELIMINA_PEDIDO_ANTIGO, s_log, msg_erro) Then
      ' CANCELA TRANSAÇÃO ANTES DA INTERAÇÃO C/ USUÁRIO !!
        GoSub EPA_EXECUTA_CANCELA_TRANSACAO
        GoTo EPA_FINALIZA_COM_FALHA
        End If
        
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dbc.CommitTrans
    On Error GoTo EPA_TRATA_ERRO
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
    elimina_pedido_antigo = True

    GoSub EPA_FECHA_TABELAS
    
    aguarde INFO_NORMAL

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPA_FINALIZA_COM_FALHA:
'======================
    GoSub EPA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPA_EXECUTA_CANCELA_TRANSACAO:
'=============================
   '~~~~~~~~~~~~~~~~~~~~
    On Error Resume Next
    dbc.RollbackTrans
   '~~~~~~~~~~~~~~~~~~~~
    On Error GoTo EPA_TRATA_ERRO
    Return



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPA_TRATA_ERRO_TRANSACAO:
'========================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
  ' CANCELA TRANSAÇÃO ANTES DA INTERAÇÃO C/ USUÁRIO !!
    GoSub EPA_EXECUTA_CANCELA_TRANSACAO
    GoSub EPA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPA_TRATA_ERRO:
'==============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    GoSub EPA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPA_FECHA_TABELAS:
'=================
    bd_desaloca_recordset t, True
    bd_desaloca_recordset sx, True
    Return


End Function


Private Function elimina_senha_desconto_antiga(ByRef msg_erro As String) As Boolean

Dim s As String
Dim s_log As String
Dim hora_inicio As Date
Dim hora_termino As Date
Dim dt_servidor As Date
Dim dt_corte As Date
Dim s_sql As String
Dim sx As ADODB.Recordset
Dim n_reg As Long
Dim n_contador As Long
Dim n_t_desconto As Long


    On Error GoTo ESDA_TRATA_ERRO


    elimina_senha_desconto_antiga = False
    msg_erro = ""
    
    hora_inicio = Now
    
    aguarde INFO_EXECUTANDO, "eliminando senhas de desconto antigas"
    
    If Not obtem_data_servidor(dt_servidor, msg_erro) Then
        s = "Erro ao consultar a data/hora do servidor!!"
        If msg_erro <> "" Then s = s & Chr(13) & Chr(13) & msg_erro
        aviso_erro s
        aguarde INFO_NORMAL
        Exit Function
        End If
        
    dt_corte = dt_servidor - CORTE_SENHA_DESCONTO_EM_DIAS
    
  ' RECORDSET
    Set sx = New ADODB.Recordset
    sx.CursorType = BD_CURSOR_SOMENTE_LEITURA
    sx.LockType = BD_POLITICA_LOCKING
    sx.CacheSize = BD_CACHE_CONSULTA
    
    
  ' IMPORTANTE: COMO A TRANSAÇÃO PODE BLOQUEAR O ACESSO DE OUTROS USUÁRIOS AO BD, NUNCA
  ' =========== DEVE HAVER INTERAÇÃO COM O USUÁRIO NAS ROTINAS CONTIDAS NA TRANSAÇÃO !!!
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dbc.BeginTrans
    On Error GoTo ESDA_TRATA_ERRO_TRANSACAO
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    n_t_desconto = 0
    n_contador = 0
    
'   APAGA REGISTROS DE AUTORIZAÇÃO DE DESCONTO SUPERIOR QUE NUNCA FORAM USADOS
    s_sql = "DELETE FROM t_DESCONTO WHERE" & _
            " (usado_status = 0)" & _
            " AND (data < " & bd_monta_data(dt_corte, False) & ")"
    dbc.Execute s_sql, n_reg
    If n_reg > 0 Then n_t_desconto = n_t_desconto + n_reg

'   APAGA REGISTROS DE AUTORIZAÇÃO DE DESCONTO SUPERIOR QUE JÁ FORAM USADOS,
'   MAS QUE O PEDIDO OU ORÇAMENTO CORRESPONDENTE JÁ FOI APAGADO
    s_sql = "SELECT t_DESCONTO.id AS desc_id," & _
            " t_PEDIDO_ITEM.abaixo_min_autorizacao AS ped_item_id," & _
            " t_ORCAMENTO_ITEM.abaixo_min_autorizacao AS orc_item_id" & _
            " FROM t_DESCONTO LEFT JOIN t_PEDIDO_ITEM ON (t_DESCONTO.id=t_PEDIDO_ITEM.abaixo_min_autorizacao)" & _
            " LEFT JOIN t_ORCAMENTO_ITEM ON (t_DESCONTO.id=t_ORCAMENTO_ITEM.abaixo_min_autorizacao)" & _
            " WHERE (t_DESCONTO.usado_status <> 0)" & _
            " AND (t_DESCONTO.data < " & bd_monta_data(dt_corte, False) & ")"
    If sx.State <> adStateClosed Then sx.Close
    sx.Open s_sql, dbc, , , adCmdText
    Do While Not sx.EOF
        n_contador = n_contador + 1
        DoEvents
        If (n_contador Mod 10) = 0 Then aguarde INFO_EXECUTANDO, "senhas para desconto: " & Format$(n_contador, FORMATO_NUMERO) & " regs processados"
    
        If (Trim$("" & sx("desc_id")) <> "") Then
            If (Trim$("" & sx("ped_item_id")) = "") And (Trim$("" & sx("orc_item_id")) = "") Then
                s_sql = "DELETE FROM t_DESCONTO WHERE" & _
                        " (id = '" & Trim$("" & sx("desc_id")) & "')"
                dbc.Execute s_sql, n_reg
                If n_reg > 0 Then n_t_desconto = n_t_desconto + n_reg
                End If
            End If
        
        sx.MoveNext
        Loop


'   LOG: QUANTIDADES EXCLUÍDAS
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "registros excluídos:" & _
            " T_DESCONTO=" & Format$(n_t_desconto, FORMATO_NUMERO)
            
'   DURAÇÃO
    hora_termino = Now
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "duração do processo: " & Format$(hora_termino - hora_inicio, "hh:mm:ss")
     
'   DATA DE CORTE
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "data de corte: " & Format$(dt_corte, FORMATO_DATA)
   
    If Not grava_log(usuario.id, "", "", "", OP_LOG_ELIMINA_SENHA_DESCONTO_ANTIGA, s_log, msg_erro) Then
      ' CANCELA TRANSAÇÃO ANTES DA INTERAÇÃO C/ USUÁRIO !!
        GoSub ESDA_EXECUTA_CANCELA_TRANSACAO
        GoTo ESDA_FINALIZA_COM_FALHA
        End If
        
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dbc.CommitTrans
    On Error GoTo ESDA_TRATA_ERRO
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
    elimina_senha_desconto_antiga = True

    GoSub ESDA_FECHA_TABELAS
    
    aguarde INFO_NORMAL

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ESDA_FINALIZA_COM_FALHA:
'=======================
    GoSub ESDA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ESDA_EXECUTA_CANCELA_TRANSACAO:
'==============================
   '~~~~~~~~~~~~~~~~~~~~
    On Error Resume Next
    dbc.RollbackTrans
   '~~~~~~~~~~~~~~~~~~~~
    On Error GoTo ESDA_TRATA_ERRO
    Return



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ESDA_TRATA_ERRO_TRANSACAO:
'=========================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
  ' CANCELA TRANSAÇÃO ANTES DA INTERAÇÃO C/ USUÁRIO !!
    GoSub ESDA_EXECUTA_CANCELA_TRANSACAO
    GoSub ESDA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ESDA_TRATA_ERRO:
'===============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    GoSub ESDA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ESDA_FECHA_TABELAS:
'==================
    bd_desaloca_recordset sx, True
    Return


End Function



Private Function elimina_orcamento_antigo(ByRef msg_erro As String) As Boolean

Dim s As String
Dim s_log As String
Dim hora_inicio As Date
Dim hora_termino As Date
Dim dt_servidor As Date
Dim dt_corte As Date
Dim s_sql As String
Dim sx As ADODB.Recordset
Dim n_reg As Long
Dim n_contador As Long
Dim dt_ult_st_fechamento As Date
Dim dt_ult_st_orcamento As Date
Dim n_t_orcamento_item As Long
Dim n_t_orcamento As Long


    On Error GoTo EOA_TRATA_ERRO


    elimina_orcamento_antigo = False
    msg_erro = ""
    
    hora_inicio = Now
    
    aguarde INFO_EXECUTANDO, "eliminando registros de orçamentos antigos"
    
    If Not obtem_data_servidor(dt_servidor, msg_erro) Then
        s = "Erro ao consultar a data/hora do servidor!!"
        If msg_erro <> "" Then s = s & Chr(13) & Chr(13) & msg_erro
        aviso_erro s
        aguarde INFO_NORMAL
        Exit Function
        End If
        
    dt_corte = dt_servidor - CORTE_ORCAMENTO_EM_DIAS
    
  ' RECORDSET
    Set sx = New ADODB.Recordset
    sx.CursorType = BD_CURSOR_SOMENTE_LEITURA
    sx.LockType = BD_POLITICA_LOCKING
    sx.CacheSize = BD_CACHE_CONSULTA
    
    
  ' EXECUTA CORTE NO BD POR DATA SOMENTE SE HOUVER MAIS REGISTROS QUE O LIMITE MÍNIMO
    s_sql = "SELECT COUNT(*) AS total FROM t_ORCAMENTO WHERE (data >= " & bd_monta_data(dt_corte, False) & ")"
    sx.Open s_sql, dbc, , , adCmdText
    n_reg = 0
    If Not sx.EOF Then If IsNumeric(sx("total")) Then n_reg = CLng(sx("total"))
    If n_reg <= CORTE_ORCAMENTO_EM_REGISTROS Then
        elimina_orcamento_antigo = True
        s_log = "Eliminação de dados antigos não foi feita porque restariam apenas " & _
                Format$(n_reg, FORMATO_NUMERO) & " registros em T_ORCAMENTO posteriores à data de corte " & _
                Format$(dt_corte, FORMATO_DATA) & " (limite mínimo: " & _
                Format$(CORTE_ORCAMENTO_EM_REGISTROS, FORMATO_NUMERO) & ")"
        Call grava_log(usuario.id, "", "", "", OP_LOG_ELIMINA_ORCAMENTO_ANTIGO, s_log, msg_erro)
        GoSub EOA_FECHA_TABELAS
        aguarde INFO_NORMAL
        Exit Function
        End If
            
            
  ' IMPORTANTE: COMO A TRANSAÇÃO PODE BLOQUEAR O ACESSO DE OUTROS USUÁRIOS AO BD, NUNCA
  ' =========== DEVE HAVER INTERAÇÃO COM O USUÁRIO NAS ROTINAS CONTIDAS NA TRANSAÇÃO !!!
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dbc.BeginTrans
    On Error GoTo EOA_TRATA_ERRO_TRANSACAO
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    n_t_orcamento_item = 0
    n_t_orcamento = 0
    n_contador = 0
    
    s_sql = "SELECT * FROM t_ORCAMENTO WHERE" & _
            " (data < " & bd_monta_data(dt_corte, False) & ")"
    
    If sx.State <> adStateClosed Then sx.Close
    sx.Open s_sql, dbc, , , adCmdText
    Do While Not sx.EOF
        
        n_contador = n_contador + 1
        DoEvents
        If (n_contador Mod 10) = 0 Then aguarde INFO_EXECUTANDO, "orçamentos: " & Format$(n_contador, FORMATO_NUMERO) & " regs processados"
            
        dt_ult_st_fechamento = 0
        dt_ult_st_orcamento = 0
                    
    '   DATA DE ST_FECHAMENTO
        If Trim$("" & sx("st_fechamento")) <> "" Then
            If IsDate(sx("fechamento_data")) Then dt_ult_st_fechamento = sx("fechamento_data")
            End If
            
    '   DATA DE ST_ORCAMENTO
        If Trim$("" & sx("st_orcamento")) = ST_ORCAMENTO_CANCELADO Then
            If IsDate(sx("cancelado_data")) Then dt_ult_st_orcamento = sx("cancelado_data")
            End If
    
    '   ANALISA DATAS DE ST_FECHAMENTO E ST_ORCAMENTO
        If dt_ult_st_fechamento >= dt_corte Then GoTo PROX_ORCAMENTO
        If dt_ult_st_orcamento >= dt_corte Then GoTo PROX_ORCAMENTO
        
        
    '   APAGA O ORÇAMENTO !!
    '   ~~~~~~~~~~~~~~~~~~~~
        s_sql = "DELETE FROM t_ORCAMENTO_ITEM WHERE" & _
                " (orcamento = '" & Trim$("" & sx("orcamento")) & "')"
        dbc.Execute s_sql, n_reg
        If n_reg > 0 Then n_t_orcamento_item = n_t_orcamento_item + n_reg
                
        s_sql = "DELETE FROM t_ORCAMENTO WHERE" & _
                " (orcamento = '" & Trim$("" & sx("orcamento")) & "')"
        dbc.Execute s_sql, n_reg
        If n_reg > 0 Then n_t_orcamento = n_t_orcamento + n_reg


PROX_ORCAMENTO:
'==============

        sx.MoveNext
        Loop
        
    
        
'   LOG: QUANTIDADES EXCLUÍDAS
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "registros excluídos:" & _
            " T_ORCAMENTO_ITEM=" & Format$(n_t_orcamento_item, FORMATO_NUMERO) & _
            " T_ORCAMENTO=" & Format$(n_t_orcamento, FORMATO_NUMERO)
            
'   DURAÇÃO
    hora_termino = Now
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "duração do processo: " & Format$(hora_termino - hora_inicio, "hh:mm:ss")
     
'   DATA DE CORTE
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "data de corte: " & Format$(dt_corte, FORMATO_DATA)
   
    If Not grava_log(usuario.id, "", "", "", OP_LOG_ELIMINA_ORCAMENTO_ANTIGO, s_log, msg_erro) Then
      ' CANCELA TRANSAÇÃO ANTES DA INTERAÇÃO C/ USUÁRIO !!
        GoSub EOA_EXECUTA_CANCELA_TRANSACAO
        GoTo EOA_FINALIZA_COM_FALHA
        End If
        
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dbc.CommitTrans
    On Error GoTo EOA_TRATA_ERRO
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
    elimina_orcamento_antigo = True

    GoSub EOA_FECHA_TABELAS
    
    aguarde INFO_NORMAL

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EOA_FINALIZA_COM_FALHA:
'======================
    GoSub EOA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EOA_EXECUTA_CANCELA_TRANSACAO:
'=============================
   '~~~~~~~~~~~~~~~~~~~~
    On Error Resume Next
    dbc.RollbackTrans
   '~~~~~~~~~~~~~~~~~~~~
    On Error GoTo EOA_TRATA_ERRO
    Return



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EOA_TRATA_ERRO_TRANSACAO:
'========================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
  ' CANCELA TRANSAÇÃO ANTES DA INTERAÇÃO C/ USUÁRIO !!
    GoSub EOA_EXECUTA_CANCELA_TRANSACAO
    GoSub EOA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EOA_TRATA_ERRO:
'==============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    GoSub EOA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EOA_FECHA_TABELAS:
'=================
    bd_desaloca_recordset sx, True
    Return


End Function



Private Function elimina_estoque_antigo(ByRef msg_erro As String) As Boolean

Dim s As String
Dim s_log As String
Dim hora_inicio As Date
Dim hora_termino As Date
Dim dt_servidor As Date
Dim dt_corte As Date
Dim s_sql As String
Dim t As ADODB.Recordset
Dim sx As ADODB.Recordset
Dim n_reg As Long
Dim n_contador As Long
Dim i_execucao As Integer
Dim n_t_estoque_movimento As Long
Dim n_t_estoque_item As Long
Dim n_t_estoque As Long
Dim pode_apagar As Boolean

    On Error GoTo EEA_TRATA_ERRO

    elimina_estoque_antigo = False
    msg_erro = ""
    
    hora_inicio = Now
    
    aguarde INFO_EXECUTANDO, "eliminando registros de estoque antigos"
    
    If Not obtem_data_servidor(dt_servidor, msg_erro) Then
        s = "Erro ao consultar a data/hora do servidor!!"
        If msg_erro <> "" Then s = s & Chr(13) & Chr(13) & msg_erro
        aviso_erro s
        aguarde INFO_NORMAL
        Exit Function
        End If
        
    dt_corte = dt_servidor - CORTE_ESTOQUE_EM_DIAS
    
  ' RECORDSET
    Set t = New ADODB.Recordset
    t.CursorType = BD_CURSOR_SOMENTE_LEITURA
    t.LockType = BD_POLITICA_LOCKING
    t.CacheSize = BD_CACHE_CONSULTA
    
    Set sx = New ADODB.Recordset
    sx.CursorType = BD_CURSOR_SOMENTE_LEITURA
    sx.LockType = BD_POLITICA_LOCKING
    sx.CacheSize = BD_CACHE_CONSULTA
    
  ' EXECUTA CORTE NO BD POR DATA SOMENTE SE HOUVER MAIS REGISTROS QUE O LIMITE MÍNIMO
    s_sql = "SELECT COUNT(*) AS total FROM t_ESTOQUE_MOVIMENTO WHERE (data >= " & bd_monta_data(dt_corte, False) & ")"
    t.Open s_sql, dbc, , , adCmdText
    n_reg = 0
    If Not t.EOF Then If IsNumeric(t("total")) Then n_reg = CLng(t("total"))
    If n_reg <= CORTE_ESTOQUE_EM_REGISTROS Then
        elimina_estoque_antigo = True
        s_log = "Eliminação de dados antigos não foi feita porque restariam apenas " & _
                Format$(n_reg, FORMATO_NUMERO) & " registros em T_ESTOQUE_MOVIMENTO posteriores à data de corte " & _
                Format$(dt_corte, FORMATO_DATA) & " (limite mínimo: " & _
                Format$(CORTE_ESTOQUE_EM_REGISTROS, FORMATO_NUMERO) & ")"
        Call grava_log(usuario.id, "", "", "", OP_LOG_ELIMINA_ESTOQUE_ANTIGO, s_log, msg_erro)
        GoSub EEA_FECHA_TABELAS
        aguarde INFO_NORMAL
        Exit Function
        End If
        
        
    i_execucao = 0
    n_contador = 0
    n_t_estoque_movimento = 0
    n_t_estoque_item = 0
    n_t_estoque = 0
        
  ' IMPORTANTE: COMO A TRANSAÇÃO PODE BLOQUEAR O ACESSO DE OUTROS USUÁRIOS AO BD, NUNCA
  ' =========== DEVE HAVER INTERAÇÃO COM O USUÁRIO NAS ROTINAS CONTIDAS NA TRANSAÇÃO !!!
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dbc.BeginTrans
    On Error GoTo EEA_TRATA_ERRO_TRANSACAO
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
'   APAGA OS REGISTROS DE MOVIMENTO QUE FORAM ANULADOS (MANTIDOS SOMENTE POR QUESTÃO DE HISTÓRICO)
    s_sql = "DELETE FROM t_ESTOQUE_MOVIMENTO WHERE" & _
            " (anulado_status <> 0)" & _
            " AND (data < " & bd_monta_data(dt_corte, False) & ")"
    dbc.Execute s_sql, n_reg
    If n_reg > 0 Then n_t_estoque_movimento = n_t_estoque_movimento + n_reg
            
            
EEA_EXECUTA_DE_NOVO:
'===================
    i_execucao = i_execucao + 1

'   APAGA REGISTROS DE MOVIMENTO GERADO POR PEDIDOS (OS REGISTROS DE MOVIMENTO QUE NÃO FORAM GERADOS POR PEDIDOS, DEVEM PERMANECER ATÉ SEREM USADOS E ANULADOS POR OPERAÇÕES DO PRÓPRIO SISTEMA)
    s_sql = "SELECT id_movimento, id_estoque," & _
            " t_ESTOQUE_MOVIMENTO.pedido AS mov_pedido, t_PEDIDO.pedido AS ped_pedido," & _
            " anulado_status, estoque" & _
            " FROM t_ESTOQUE_MOVIMENTO LEFT JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
            " WHERE" & _
            " (t_ESTOQUE_MOVIMENTO.data < " & bd_monta_data(dt_corte, False) & ")" & _
            " AND (kit = 0)" & _
            " AND (t_ESTOQUE_MOVIMENTO.pedido <> '') AND (" & bd_monta_is_not_null("t_ESTOQUE_MOVIMENTO.pedido") & ")"
                
    If sx.State <> adStateClosed Then sx.Close
    sx.Open s_sql, dbc, , , adCmdText
    Do While Not sx.EOF
        n_contador = n_contador + 1
        DoEvents
        If (n_contador Mod 10) = 0 Then aguarde INFO_EXECUTANDO, "estoque: " & Format$(n_contador, FORMATO_NUMERO) & " regs processados"
        
    '   É UM MOVIMENTO GERADO POR PEDIDO ?
        If Trim$("" & sx("mov_pedido")) <> "" Then
        '   O PEDIDO JÁ NÃO EXISTE MAIS ?
            If Trim$("" & sx("ped_pedido")) = "" Then
                pode_apagar = False
            '   NÃO ELIMINA REGISTROS ATIVOS DE DEVOLUÇÃO DE MERCADORIAS, MESMO QUE O PEDIDO JÁ TENHA SIDO ELIMINADO (ESTES REGISTROS DEVEM PERMANECER ATÉ SER USADO E ANULADO PELA OPERAÇÃO DE TRANSFERÊNCIA ENTRE ESTOQUES)
                If (sx("anulado_status") = 0) And (Trim$("" & sx("estoque")) = ID_ESTOQUE_DEVOLUCAO) Then
                '   NOP
                Else
                    If Trim$("" & sx("id_estoque")) = "" Then
                    '   NÃO ESTÁ VINCULADO A NENHUM LOTE DO ESTOQUE
                        pode_apagar = True
                    Else
                    '   O LOTE DO ESTOQUE VINCULADO JÁ FOI TODO CONSUMIDO ?
                        n_reg = 0
                        s_sql = "SELECT SUM(qtde-qtde_utilizada) AS total FROM t_ESTOQUE_ITEM WHERE" & _
                                " (id_estoque = '" & Trim$("" & sx("id_estoque")) & "')" & _
                                " AND ((qtde-qtde_utilizada) > 0)"
                        If t.State <> adStateClosed Then t.Close
                        t.Open s_sql, dbc, , , adCmdText
                        If Not t.EOF Then If IsNumeric(t("total")) Then n_reg = CLng(t("total"))
                        If n_reg = 0 Then pode_apagar = True
                        End If
                    End If
                    
            '   APAGA O REGISTRO DE MOVIMENTO !!
                If pode_apagar Then
                    s_sql = "DELETE FROM t_ESTOQUE_MOVIMENTO WHERE" & _
                            " (id_movimento = '" & Trim$("" & sx("id_movimento")) & "')"
                    dbc.Execute s_sql, n_reg
                    If n_reg > 0 Then n_t_estoque_movimento = n_t_estoque_movimento + n_reg
                    End If
                End If
            End If
            
        sx.MoveNext
        Loop
            
            
'   APAGA REGISTROS DE MOVIMENTO GERADO POR CONVERSÃO DE KITS
    s_sql = "SELECT id_movimento, t_ESTOQUE_MOVIMENTO.id_estoque," & _
            " kit_id_estoque, t_ESTOQUE.id_estoque AS main_kit_id_estoque" & _
            " FROM t_ESTOQUE_MOVIMENTO LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_MOVIMENTO.kit_id_estoque=t_ESTOQUE.id_estoque)" & _
            " WHERE" & _
            " (t_ESTOQUE_MOVIMENTO.data < " & bd_monta_data(dt_corte, False) & ")" & _
            " AND (t_ESTOQUE_MOVIMENTO.kit <> 0)" & _
            " AND (t_ESTOQUE_MOVIMENTO.kit_id_estoque <> '') AND (" & bd_monta_is_not_null("t_ESTOQUE_MOVIMENTO.kit_id_estoque") & ")"

    If sx.State <> adStateClosed Then sx.Close
    sx.Open s_sql, dbc, , , adCmdText
    Do While Not sx.EOF
        n_contador = n_contador + 1
        DoEvents
        If (n_contador Mod 10) = 0 Then aguarde INFO_EXECUTANDO, "estoque: " & Format$(n_contador, FORMATO_NUMERO) & " regs processados"
    
    '   O LOTE DO ESTOQUE DO KIT JÁ FOI ELIMINADO ?
        If Trim$("" & sx("main_kit_id_estoque")) = "" Then
            pode_apagar = False
        '   O LOTE DO ESTOQUE VINCULADO JÁ FOI TODO CONSUMIDO ?
            n_reg = 0
            s_sql = "SELECT SUM(qtde-qtde_utilizada) AS total FROM t_ESTOQUE_ITEM WHERE" & _
                    " (id_estoque = '" & Trim$("" & sx("id_estoque")) & "')" & _
                    " AND ((qtde-qtde_utilizada) > 0)"
            If t.State <> adStateClosed Then t.Close
            t.Open s_sql, dbc, , , adCmdText
            If Not t.EOF Then If IsNumeric(t("total")) Then n_reg = CLng(t("total"))
            If n_reg = 0 Then pode_apagar = True
                    
            If pode_apagar Then
                s_sql = "DELETE FROM t_ESTOQUE_MOVIMENTO WHERE" & _
                        " (id_movimento = '" & Trim$("" & sx("id_movimento")) & "')"
                dbc.Execute s_sql, n_reg
                If n_reg > 0 Then n_t_estoque_movimento = n_t_estoque_movimento + n_reg
                End If
            End If
            
        sx.MoveNext
        Loop
    
    
'   APAGA REGISTROS DE LOTES DO ESTOQUE
    s_sql = "SELECT id_estoque FROM t_ESTOQUE WHERE" & _
            " (data_entrada < " & bd_monta_data(dt_corte, False) & ")" & _
            " AND (data_ult_movimento < " & bd_monta_data(dt_corte, False) & ")"
    If sx.State <> adStateClosed Then sx.Close
    sx.Open s_sql, dbc, , , adCmdText
    Do While Not sx.EOF
        n_contador = n_contador + 1
        DoEvents
        If (n_contador Mod 10) = 0 Then aguarde INFO_EXECUTANDO, "estoque: " & Format$(n_contador, FORMATO_NUMERO) & " regs processados"
    
    '   O LOTE DO ESTOQUE JÁ FOI TODO CONSUMIDO ?
        s_sql = "SELECT SUM(qtde-qtde_utilizada) AS total FROM t_ESTOQUE_ITEM WHERE" & _
                " (id_estoque = '" & Trim$("" & sx("id_estoque")) & "')" & _
                " AND ((qtde-qtde_utilizada) > 0)"
        n_reg = 0
        If t.State <> adStateClosed Then t.Close
        t.Open s_sql, dbc, , , adCmdText
        If Not t.EOF Then If IsNumeric(t("total")) Then n_reg = CLng(t("total"))
        If n_reg = 0 Then
        '   HÁ REGISTROS DE MOVIMENTO RELACIONADOS ?
            s_sql = "SELECT id_movimento FROM t_ESTOQUE_MOVIMENTO WHERE" & _
                    " (id_estoque = '" & Trim$("" & sx("id_estoque")) & "')"
            If t.State <> adStateClosed Then t.Close
            t.Open s_sql, dbc, , , adCmdText
            If t.EOF Then
            '   APAGA O LOTE DO ESTOQUE !!
                s_sql = "DELETE FROM t_ESTOQUE_ITEM WHERE" & _
                        " (id_estoque = '" & Trim$("" & sx("id_estoque")) & "')"
                dbc.Execute s_sql, n_reg
                If n_reg > 0 Then n_t_estoque_item = n_t_estoque_item + n_reg
                
                s_sql = "DELETE FROM t_ESTOQUE WHERE" & _
                        " (id_estoque = '" & Trim$("" & sx("id_estoque")) & "')"
                dbc.Execute s_sql, n_reg
                If n_reg > 0 Then n_t_estoque = n_t_estoque + n_reg
                End If
            End If
                
        sx.MoveNext
        Loop
    
    
  ' EXECUTA NOVAMENTE ?
  ' LEMBRE-SE: AO CONVERTER PRODUTOS EM KITS, OS KITS SÃO GRAVADOS
  ' NO ESTOQUE COMO SE FOSSEM PRODUTOS NORMAIS.
  ' 1) OS REGISTROS DO ESTOQUE REFERENTES AOS KITS SOMENTE SERÃO
  ' ELIMINADOS SE OS PEDIDOS RELACIONADOS JÁ TIVEREM SIDO ELIMINADOS.
  ' 2) OS REGISTROS DO ESTOQUE DOS PRODUTOS QUE COMPÕEM OS KITS SOMENTE
  ' SERÃO ELIMINADOS SE OS REGISTROS DE ESTOQUE DOS KITS JÁ TIVEREM
  ' SIDO ELIMINADOS.
  ' PORTANTO, SÃO NECESSÁRIAS DUAS EXECUÇÕES P/ FAZER A LIMPEZA COMPLETA.
    If i_execucao = 1 Then GoTo EEA_EXECUTA_DE_NOVO
    
    
'   LOG: QUANTIDADES EXCLUÍDAS
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "registros excluídos:" & _
            " T_ESTOQUE_MOVIMENTO=" & Format$(n_t_estoque_movimento, FORMATO_NUMERO) & _
            " T_ESTOQUE_ITEM=" & Format$(n_t_estoque_item, FORMATO_NUMERO) & _
            " T_ESTOQUE=" & Format$(n_t_estoque, FORMATO_NUMERO)

'   DURAÇÃO
    hora_termino = Now
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "duração do processo: " & Format$(hora_termino - hora_inicio, "hh:mm:ss")
    
'   DATA DE CORTE
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "data de corte: " & Format$(dt_corte, FORMATO_DATA)
    
    If Not grava_log(usuario.id, "", "", "", OP_LOG_ELIMINA_ESTOQUE_ANTIGO, s_log, msg_erro) Then
      ' CANCELA TRANSAÇÃO ANTES DA INTERAÇÃO C/ USUÁRIO !!
        GoSub EEA_EXECUTA_CANCELA_TRANSACAO
        GoTo EEA_FINALIZA_COM_FALHA
        End If
    
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    dbc.CommitTrans
    On Error GoTo EEA_TRATA_ERRO
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    elimina_estoque_antigo = True

    GoSub EEA_FECHA_TABELAS
    
    aguarde INFO_NORMAL

Exit Function






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EEA_FINALIZA_COM_FALHA:
'======================
    GoSub EEA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EEA_EXECUTA_CANCELA_TRANSACAO:
'=============================
   '~~~~~~~~~~~~~~~~~~~~
    On Error Resume Next
    dbc.RollbackTrans
   '~~~~~~~~~~~~~~~~~~~~
    On Error GoTo EEA_TRATA_ERRO
    Return



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EEA_TRATA_ERRO_TRANSACAO:
'========================
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
  ' CANCELA TRANSAÇÃO ANTES DA INTERAÇÃO C/ USUÁRIO !!
    GoSub EEA_EXECUTA_CANCELA_TRANSACAO
    GoSub EEA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EEA_TRATA_ERRO:
'==============
    msg_erro = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    GoSub EEA_FECHA_TABELAS
    aguarde INFO_NORMAL
    Exit Function



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EEA_FECHA_TABELAS:
'=================
    bd_desaloca_recordset t, True
    bd_desaloca_recordset sx, True
    Return


End Function


Function le_registro(ByRef Fnum As Integer, ByVal codigo_delimitador As String) As String

Dim c As String
Dim s_registro As String

    s_registro = ""
    
    Do While Not EOF(Fnum)
        c = Input(1, Fnum)
        If c = codigo_delimitador Then
            Exit Do
        Else
            s_registro = s_registro & c
            End If
        Loop

    le_registro = s_registro
    
End Function



Private Sub tabela_produtos_carrega()
' ______________________________________________________________________________
'|
'|  CARREGA A PARTIR DE PLANILHA EXCEL OS DADOS REFERENTES À TABELA DE
'|  PRODUTOS (T_PRODUTO) E PRODUTOS POR LOJA (T_PRODUTO_LOJA).
'|
'|  PARA MINIMIZAR O TEMPO EM QUE OS REGISTROS DAS TABELAS DE PRODUÇÃO FICAM
'|  BLOQUEADAS DURANTE O PROCESSAMENTO, A ROTINA CRIA TABELAS TEMPORÁRIAS
'|  COM OS DADOS DE PRODUÇÃO. O PROCESSAMENTO É REALIZADO SOBRE AS TABELAS
'|  TEMPORÁRIAS E, AO FINAL, SUBSTITUI OS DADOS DAS TABELAS DE PRODUÇÃO DE
'|  UMA ÚNICA VEZ.
'|
'|  ********* IMPORTANTE!!! *********
'|  AUTOMAÇÃO EXCEL: LEIA OBSERVAÇÕES IMPORTANTES NA SEÇÃO DE DECLARAÇÕES DO MÓDULO.
'|

Const PREFIXO_TABELA_TEMPORARIA = "tmpAdm__"
Const MAX_LINHAS_PLANILHA = 65536
Const MAX_XLS_LINHAS_EM_BRANCO_CONSECUTIVAS = 20
Const SEPARADOR_CODIGO = "-"
Const xlUnderlineStyleSingle = 2
Const VL_COD_PRECO_A_IGNORAR = -1

' STRINGS
Dim s As String
Dim s_aux As String
Dim s_ret As String
Dim s_nome_planilha As String
Dim s_destino_com_prefixo As String
Dim s_destino_sem_prefixo As String
Dim s_loja_base_replicacao As String
Dim msg_erro As String
Dim s_log As String
Dim s_descricao As String
Dim strPrimeiraColunaPlanilha As String
Dim strUltimaColunaPlanilha As String
Dim strColunaDescricao As String
Dim strCaracter As String
Dim strNomeTabela As String
Dim strCST As String
Dim intPosIdentificador As Integer
Dim intPosAposIdentificador As Integer

' CONTADORES
Dim i As Long
Dim intCounterAux As Long
Dim n_reg As Long
Dim ic As Integer
Dim il As Integer
Dim idx_basica As Integer
Dim xls_linhas_em_branco_consecutivas As Long
Dim n_t_produto As Long
Dim n_t_produto_loja As Long
Dim icol As Integer
Dim ilinha As Long
Dim icor As Long
Dim lngRecordsAffected As Long

' FLAGS
Dim foi_selecionada As Boolean
Dim existe_loja As Boolean
Dim cadastrado As Boolean
Dim processou_tabela_basica As Boolean
Dim blnNegritoAtivado As Boolean
Dim blnFecharTagNegrito As Boolean
Dim blnItalicoAtivado As Boolean
Dim blnFecharTagItalico As Boolean
Dim blnSublinhadoAtivado As Boolean
Dim blnFecharTagSublinhado As Boolean

' ETC.
Dim q As Variant
Dim rp As TIPO_T_PRODUTO
Dim rpl As TIPO_T_PRODUTO_LOJA
Dim v_fabricante() As TIPO_CODIGO_X_DESCRICAO
Dim v_produto() As TIPO_CODIGO_X_DESCRICAO
Dim v_lj_processada() As String
Dim v_grupo_processado() As String
Dim v_loja() As String
Dim v_alerta
Dim matriz()
Dim hora_inicio As Date
Dim tempo_planilha_open As Date
Dim tempo_transferencia As Date
Dim v_xls_a_selecionar() As String
Dim v_xls_selecionada() As String

' BANCO DE DADOS
Dim rs As ADODB.Recordset
Dim tPL As ADODB.Recordset

' AUTOMAÇÃO EXCEL
Const xlEXTENSAO_ARQUIVO_SAIDA = "XLSX"
Dim xl_nome_arquivo As String
Dim oXL As Object
Dim oWS As Object
Dim oRANGE As Object
Dim oDescricao As Object

    
    On Error GoTo TPC_TRATA_ERRO
   
    aguarde INFO_EXECUTANDO, "preparando painel para selecionar planilha"
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' OBTÉM NOME DA PLANILHA
  
  ' RECUPERA DO REGISTRY O NOME DO ARQUIVO USADO DA ÚLTIMA VEZ
    s_ret = ""
    If registry_le_string(REG_CHAVE_ADM_TABELA_PRODUTOS, REG_CAMPO_ADM_TABELA_PRODUTOS_PLANILHA, s) Then
        If Trim$(s) <> "" Then If DirectoryExists(ExtractFilePath(s), msg_erro) Then s_ret = s
        End If
    
'   CONFIGURA DIALOG BOX
    fcmd.FileName = s_ret
    s = ExtractFilePath(s_ret)
    If s = "" Then s = "c:\"
    fcmd.InitDir = s
    fcmd.DialogTitle = "Abrir planilha com tabela de produtos ..."
    fcmd.DefaultExt = xlEXTENSAO_ARQUIVO_SAIDA
    fcmd.Filter = "Planilha MS-Excel|*." & LCase$(xlEXTENSAO_ARQUIVO_SAIDA)
    fcmd.FilterIndex = 0
    fcmd.Flags = fcmd.Flags Or cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNExplorer
    fcmd.CancelError = True
'   MOSTRA DIALOG BOX
    aguarde INFO_NORMAL, m_id
    On Error GoTo TPC_TRATA_CANCELA_DIALOG_BOX
    fcmd.ShowOpen
    On Error GoTo TPC_TRATA_ERRO
    xl_nome_arquivo = fcmd.FileName
'   CANCELOU ?
    If xl_nome_arquivo = "" Then Exit Sub
'   ARQUIVO EXISTE ?
    If Not FileExists(xl_nome_arquivo, msg_erro) Then
        aviso_erro "Arquivo " & xl_nome_arquivo & " não foi encontrado!!"
        Exit Sub
        End If
        
'   EXTENSÃO VÁLIDA ?
    If UCase$(ExtractFileExt(xl_nome_arquivo)) <> UCase$(xlEXTENSAO_ARQUIVO_SAIDA) Then
        aviso_erro "Arquivo de saída selecionado possui extensão inválida!!" & _
                    vbCrLf & "O arquivo deve possuir a extensão " & UCase$(xlEXTENSAO_ARQUIVO_SAIDA) & "!!"
        Exit Sub
        End If
       
'   NOME DO ARQUIVO POSSUI REFERÊNCIA AO IDENTIFICADOR DO AMBIENTE ?
    '(após o identificador, deve existir um caractere diferente de letra ou número)
    intPosIdentificador = InStr(UCase$(ExtractFileName(xl_nome_arquivo)), UCase$(identificador_ambiente_padrao))
    intPosAposIdentificador = intPosIdentificador + Len(identificador_ambiente_padrao)
    If intPosIdentificador <= 0 Or _
        IsLetra(Mid(ExtractFileName(xl_nome_arquivo), intPosAposIdentificador, 1)) Or _
        IsAlgarismo(Mid(ExtractFileName(xl_nome_arquivo), intPosAposIdentificador, 1)) Then
        aviso_erro "Nome do arquivo não possui referência ao ambiente atual!!" & _
                    vbCrLf & "O arquivo deve referenciar " & identificador_ambiente_padrao & "!!"
        Exit Sub
        End If
   
   
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' CRIA OBJETO DE AUTOMAÇÃO EXCEL
      
    aguarde INFO_EXECUTANDO, "abrindo planilha"
    
    hora_inicio = Now
    
  ' CRIA INSTÂNCIA DO EXCEL
    Set oXL = CreateObject("Excel.Application")
    With oXL
      ' MANTÉM A INSTÂNCIA DO EXCEL INVISÍVEL
        .Visible = False
      ' NÃO MOSTRA MENSAGENS DO EXCEL NA TELA (NÃO DEIXA O EXCEL INTERAGIR C/ O USUÁRIO)
        .DisplayAlerts = False
        End With
   
  ' PREVINE A MENSAGEM DE ERRO "This action cannot be completed because the other application is busy. Choose 'Switch To' to activate the busy application and correct the problem."
  ' O DEFAULT É 5000 (VALOR EM MILISEGUNDOS).
    App.OleRequestPendingTimeout = 2147483647
   
    oXL.Workbooks.Open xl_nome_arquivo
    tempo_planilha_open = Now - hora_inicio
        
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' PAINEL PARA SELEÇÃO DAS PLANILHAS A PROCESSAR

'   RECORDSET
    Set rs = New ADODB.Recordset
    With rs
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
    Set tPL = New ADODB.Recordset
    With tPL
        .CursorType = BD_CURSOR_SOMENTE_LEITURA
        .LockType = BD_POLITICA_LOCKING
        .CacheSize = BD_CACHE_CONSULTA
        End With
    
'   APRESENTA PAINEL P/ O USUÁRIO SELECIONAR AS PLANILHAS QUE DESEJA CARREGAR
    ReDim v_xls_a_selecionar(0)
    v_xls_a_selecionar(UBound(v_xls_a_selecionar)) = ""
    
    ReDim v_xls_selecionada(0)
    v_xls_selecionada(UBound(v_xls_selecionada)) = ""
    
'   COLOCA A TABELA BÁSICA, SE HOUVER, NA PRIMEIRA POSIÇÃO
    For ic = 1 To oXL.Workbooks(1).Worksheets.Count
        s = UCase$(Trim$(oXL.Workbooks(1).Worksheets(ic).Name))
        If IsTabelaBasicaProdutos(s) Then
            If v_xls_a_selecionar(UBound(v_xls_a_selecionar)) <> "" Then ReDim Preserve v_xls_a_selecionar(UBound(v_xls_a_selecionar) + 1)
            v_xls_a_selecionar(UBound(v_xls_a_selecionar)) = s
            Exit For
            End If
        Next

'   CARREGA A LISTA DE GRUPOS DE LOJAS
    For ic = 1 To oXL.Workbooks(1).Worksheets.Count
        s = UCase$(Trim$(oXL.Workbooks(1).Worksheets(ic).Name))
        If Not IsTabelaBasicaProdutos(s) And IsNumeroGrupodeLojas(s) Then
            If v_xls_a_selecionar(UBound(v_xls_a_selecionar)) <> "" Then ReDim Preserve v_xls_a_selecionar(UBound(v_xls_a_selecionar) + 1)
            v_xls_a_selecionar(UBound(v_xls_a_selecionar)) = s
            End If
        Next
        
'   CARREGA A LISTA DE LOJAS
    For ic = 1 To oXL.Workbooks(1).Worksheets.Count
        s = UCase$(Trim$(oXL.Workbooks(1).Worksheets(ic).Name))
        If Not IsTabelaBasicaProdutos(s) And IsNumeroLoja(s) Then
            If v_xls_a_selecionar(UBound(v_xls_a_selecionar)) <> "" Then ReDim Preserve v_xls_a_selecionar(UBound(v_xls_a_selecionar) + 1)
            v_xls_a_selecionar(UBound(v_xls_a_selecionar)) = s
            End If
        Next
        
'   NORMALIZA O TAMANHO DOS CÓDIGOS
    i = 0
    For ic = LBound(v_xls_a_selecionar) To UBound(v_xls_a_selecionar)
        If Len(v_xls_a_selecionar(ic)) > i Then i = Len(v_xls_a_selecionar(ic))
        Next
        
    For ic = LBound(v_xls_a_selecionar) To UBound(v_xls_a_selecionar)
        s = v_xls_a_selecionar(ic)
        s = String$(i - Len(s), " ") & s
        v_xls_a_selecionar(ic) = s
        Next
        
'   CARREGA DESCRIÇÃO PARA AS OPÇÕES
    For ic = LBound(v_xls_a_selecionar) To UBound(v_xls_a_selecionar)
        s_descricao = ""
        If IsTabelaBasicaProdutos(Trim$(v_xls_a_selecionar(ic))) Then
            s_descricao = "Tabela Básica"
        ElseIf IsNumeroGrupodeLojas(Trim$(v_xls_a_selecionar(ic))) Then
            s = Trim$(v_xls_a_selecionar(ic))
            s = remove_prefixo_do_numero(s, PREFIXO_NUMERO_GRUPO_LOJAS)
            s = normaliza_codigo(s, TAM_MIN_GRUPO_LOJAS)
            s = "SELECT loja FROM t_LOJA_GRUPO_ITEM WHERE (grupo='" & s & "') ORDER BY CONVERT(smallint,loja)"
            If rs.State <> adStateClosed Then rs.Close
            rs.Open s, dbc, , , adCmdText
            Do While Not rs.EOF
                If s_descricao <> "" Then s_descricao = s_descricao & ", "
                s_descricao = s_descricao & Trim$("" & rs("loja"))
                rs.MoveNext
                Loop
                
            If s_descricao = "" Then s_descricao = "nenhuma loja"
            s_descricao = "Grupo de Lojas: " & s_descricao
        Else
            s = Trim$(v_xls_a_selecionar(ic))
            s = remove_prefixo_do_numero(s, PREFIXO_NUMERO_LOJA)
            s = normaliza_codigo(s, TAM_MIN_LOJA)
            s = "SELECT nome, razao_social FROM t_LOJA WHERE (loja='" & s & "')"
            If rs.State <> adStateClosed Then rs.Close
            rs.Open s, dbc, , , adCmdText
            If rs.EOF Then
                s_descricao = "não cadastrada"
            Else
                s_descricao = Trim$("" & rs("nome"))
                If s_descricao = "" Then s_descricao = Trim$("" & rs("razao_social"))
                End If
            s_descricao = iniciais_em_maiusculas(s_descricao)
            End If
        
        v_xls_a_selecionar(ic) = v_xls_a_selecionar(ic) & " " & SEPARADOR_CODIGO & " " & s_descricao
        Next
        
        
'   EXIBE PAINEL PARA SELEÇÃO
    aguarde INFO_NORMAL, m_id
    f_CHECKBOX.Caption = "Seleção de planilhas para transferência de dados"
    If Not f_CHECKBOX.executa_selecao(v_xls_a_selecionar(), v_xls_selecionada()) Then
        GoSub TPC_FECHA_TABELAS_E_OBJETOS
        Exit Sub
        End If
        
'   SELECIONOU PELO MENOS UMA ?
    existe_loja = False
    For ic = LBound(v_xls_selecionada) To UBound(v_xls_selecionada)
        If Trim$(v_xls_selecionada(ic)) <> "" Then
            existe_loja = True
            Exit For
            End If
        Next
        
    If Not existe_loja Then
        aviso "Nenhuma planilha foi selecionada!!"
        GoSub TPC_FECHA_TABELAS_E_OBJETOS
        Exit Sub
        End If
        
'   CONFIRMA ?
    s = "Executa a transferência dos dados da planilha ?" & _
        vbCrLf & vbCrLf & "Planilha: " & xl_nome_arquivo
    If Not confirma(s) Then
        GoSub TPC_FECHA_TABELAS_E_OBJETOS
        Exit Sub
        End If
  
  ' GRAVA NOME DO ARQUIVO DA PLANILHA NO REGISTRY
    Call registry_grava_string(REG_CHAVE_ADM_TABELA_PRODUTOS, REG_CAMPO_ADM_TABELA_PRODUTOS_PLANILHA, xl_nome_arquivo)
   
   
   
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' INFORMAÇÕES PARA O LOG
    
    s_log = "Planilha: " & xl_nome_arquivo
    
    s_descricao = ""
    For ic = LBound(v_xls_selecionada) To UBound(v_xls_selecionada)
        If Trim$(v_xls_selecionada(ic)) <> "" Then
            If s_descricao <> "" Then s_descricao = s_descricao & ", "
            s_descricao = s_descricao & Trim$(separa_campo(Trim$(v_xls_selecionada(ic)), SEPARADOR_CODIGO))
            End If
        Next
        
    If s_descricao <> "" Then
        s_descricao = "planilhas selecionadas: " & s_descricao
        If s_log <> "" Then s_log = s_log & "; "
        s_log = s_log & s_descricao
        End If
    


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' PROCESSA A TABELA BÁSICA DE PRODUTOS (T_PRODUTO)
        
  ' MOSTRA TELA INFORMATIVA COM OPÇÃO DE CANCELAR
    f_PROCESSANDO.exibe_painel

    hora_inicio = Now
        
    idx_basica = -1
    For ic = 1 To oXL.Workbooks(1).Worksheets.Count
        s = Trim$(oXL.Workbooks(1).Worksheets(ic).Name)
        If IsTabelaBasicaProdutos(s) Then
            idx_basica = ic
            Exit For
            End If
        Next
        
    aguarde INFO_EXECUTANDO, "consultando lista de fabricantes"
    
'   CARREGA LISTA DE FABRICANTES
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ReDim v_fabricante(0)
    With v_fabricante(UBound(v_fabricante))
        .codigo = ""
        .descricao = ""
        End With
        
    s = "SELECT fabricante, nome, razao_social FROM t_FABRICANTE ORDER BY fabricante"
    If rs.State <> adStateClosed Then rs.Close
    rs.Open s, dbc, , , adCmdText
    i = 0
    Do While Not rs.EOF
        i = i + 1
        ReDim Preserve v_fabricante(i)
        With v_fabricante(i)
            .codigo = Trim$("" & rs("fabricante"))
            .descricao = Trim$("" & rs("nome"))
            If .descricao = "" Then .descricao = Trim$("" & rs("razao_social"))
            End With
        rs.MoveNext
        Loop
    
    ORDENA_codigo_X_descricao v_fabricante(), 1, UBound(v_fabricante)
    
    processou_tabela_basica = False
    
    ReDim v_lj_processada(0)
    v_lj_processada(UBound(v_lj_processada)) = ""
    
    ReDim v_grupo_processado(0)
    v_grupo_processado(UBound(v_grupo_processado)) = ""


'   CERTIFICA-SE DE QUE O PREFIXO DE TABELAS TEMPORÁRIAS FOI DEFINIDO
    If Trim$(PREFIXO_TABELA_TEMPORARIA) = "" Then
        aviso "O prefixo para nomear as tabelas temporárias não está definido!!"
        GoSub TPC_FECHA_TABELAS_E_OBJETOS
        Exit Sub
        End If
        
'   REMOVE TABELAS TEMPORÁRIAS QUE POSSA TER RESTADO DE PROCESSAMENTO ANTERIOR NÃO FINALIZADO CORRETAMENTE
    aguarde INFO_EXECUTANDO, "eliminando tabelas temporárias remanescentes"
    If Not remove_tabelas_temporarias(PREFIXO_TABELA_TEMPORARIA, msg_erro) Then
        If msg_erro <> "" Then msg_erro = vbCrLf & vbCrLf & msg_erro
        msg_erro = "Erro ao se certificar de que não restaram tabelas temporárias remanescentes de processamento anterior!!" & _
                    msg_erro
        aviso_erro msg_erro
        GoSub TPC_FECHA_TABELAS_E_OBJETOS
        Exit Sub
        End If
    
'   CRIA AS TABELAS TEMPORÁRIAS COM O CONTEÚDO DAS TABELAS DE PRODUÇÃO PARA SEREM USADAS NO PROCESSAMENTO DA PLANILHA
    On Error GoTo TPC_TRATA_ERRO_TAB_TEMP
    aguarde INFO_EXECUTANDO, "criando tabelas temporárias para processamento"
    
'   T_PRODUTO
'   ~~~~~~~~~
'   TABELA
    strNomeTabela = "t_PRODUTO"
    s = "SELECT * INTO " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & " FROM " & strNomeTabela
    dbc.Execute s, lngRecordsAffected
    
'   CONSTRAINT DEFAULT
    s = "ALTER TABLE " & _
            PREFIXO_TABELA_TEMPORARIA & strNomeTabela & _
        " ADD CONSTRAINT DF_" & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & "_deposito_zona_id DEFAULT(0)" & _
        " FOR deposito_zona_id"
    dbc.Execute s
    
'   CONSTRAINT DEFAULT
    s = "ALTER TABLE " & _
            PREFIXO_TABELA_TEMPORARIA & strNomeTabela & _
        " ADD CONSTRAINT DF_" & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & "_farol_qtde_comprada DEFAULT(0)" & _
        " FOR farol_qtde_comprada"
    dbc.Execute s
    
'   ÍNDICES
    s = "CREATE INDEX " & _
            PREFIXO_TABELA_TEMPORARIA & strNomeTabela & "_produto" & _
        " ON " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & _
            " (produto)"
    dbc.Execute s
    
    s = "CREATE INDEX " & _
            PREFIXO_TABELA_TEMPORARIA & strNomeTabela & "_ean" & _
        " ON " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & _
            " (ean)"
    dbc.Execute s
    
    s = "CREATE INDEX " & _
            PREFIXO_TABELA_TEMPORARIA & strNomeTabela & "_excl_stat" & _
        " ON " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & _
            " (excluido_status)"
    dbc.Execute s
    
'   T_PRODUTO_LOJA
'   ~~~~~~~~~~~~~~
'   TABELA
    strNomeTabela = "t_PRODUTO_LOJA"
    s = "SELECT * INTO " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & " FROM " & strNomeTabela
    dbc.Execute s, lngRecordsAffected
    
'   ÍNDICES
    s = "CREATE INDEX " & _
            PREFIXO_TABELA_TEMPORARIA & strNomeTabela & "_produto" & _
        " ON " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & _
            " (produto)"
    dbc.Execute s

    s = "CREATE INDEX " & _
            PREFIXO_TABELA_TEMPORARIA & strNomeTabela & "_loja" & _
        " ON " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & _
            " (loja)"
    dbc.Execute s

    s = "CREATE INDEX " & _
            PREFIXO_TABELA_TEMPORARIA & strNomeTabela & "_excl_stat" & _
        " ON " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & _
            " (excluido_status)"
    dbc.Execute s

'   T_PRODUTO_X_ALERTA
'   ~~~~~~~~~~~~~~~~~~
'   TABELA
    strNomeTabela = "t_PRODUTO_X_ALERTA"
    s = "SELECT * INTO " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & " FROM " & strNomeTabela
    dbc.Execute s, lngRecordsAffected

'   ÍNDICES
    s = "CREATE INDEX " & _
            PREFIXO_TABELA_TEMPORARIA & strNomeTabela & "_produto" & _
        " ON " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela & _
            " (produto)"
    dbc.Execute s


'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   INICIA PROCESSAMENTO DOS DADOS DA PLANILHA
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "processamento dos dados da planilha"
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    rs.CursorType = BD_CURSOR_EDICAO
    rs.LockType = BD_POLITICA_LOCKING
    
'   TABELA BÁSICA FOI SELECIONADA ?
    foi_selecionada = False
    For i = LBound(v_xls_selecionada) To UBound(v_xls_selecionada)
        s = Trim$(v_xls_selecionada(i))
        If s <> "" Then
            s = Trim$(separa_campo(s, SEPARADOR_CODIGO))
            If IsTabelaBasicaProdutos(s) Then
                foi_selecionada = True
                Exit For
                End If
            End If
        Next
        
    If (idx_basica > -1) And foi_selecionada Then
    '   CONTROLE DE PRODUTOS EXCLUÍDOS
        s = "UPDATE " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO SET excluido_status=1"
        dbc.Execute s, lngRecordsAffected
        
        s = "UPDATE " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA SET excluido_status=1"
        dbc.Execute s, lngRecordsAffected
        
        s = "UPDATE " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_X_ALERTA SET excluido_status=1"
        dbc.Execute s, lngRecordsAffected
        
        Set oWS = oXL.Workbooks(1).Worksheets(idx_basica)
        
        n_t_produto = 0
        xls_linhas_em_branco_consecutivas = 0
        n_reg = 0
        
        For ilinha = 2 To MAX_LINHAS_PLANILHA
            DoEvents
            If f_PROCESSANDO.processo_cancelado Then
                msg_erro = "Operação cancelada!!"
                GoTo TPC_ABORTA_PROCESSAMENTO
                End If
            
            If (n_reg Mod 10) = 0 Then aguarde INFO_EXECUTANDO, "transferindo dados da tabela básica: " & Format$(n_reg, "###,###,###,##0") & " regs"
            n_reg = n_reg + 1
            
        '   COLUNAS DA PLANILHA
        '   A = FABRICANTE
        '   B = PRODUTO
        '   C = DESCRIÇÃO
        '   D = EAN
        '   E = GRUPO
        '   F = POTÊNCIA (Btu/h)
        '   G = CICLO (F/QF)
        '   H = PREÇO FABRICANTE
        '   I = PESO (KG)
        '   J = QTDE VOLUMES
        '   K = CUBAGEM
        '   L = NCM
        '   M = CST
        '   N = POSIÇÃO MERCADO (BÁSICO/PREMIUM)
        '   O = PERCENTUAL MVA ST
        '   P = CUSTO2
        '   Q = ALERTAS
        '   R = ESTOQUE CRÍTICO
        '   S = DESCONTINUADO
            strPrimeiraColunaPlanilha = "A"
            strUltimaColunaPlanilha = "S"
            strColunaDescricao = "C"
            Set oRANGE = oWS.Range(strPrimeiraColunaPlanilha & CStr(ilinha) & ":" & strUltimaColunaPlanilha & CStr(ilinha))
            matriz = oRANGE.Value
            
            With rp
                icol = LBound(matriz, 2) - 1
                
            '   LÊ A LINHA
            '   COL "A" = FABRICANTE
                icol = icol + 1
                .fabricante = Trim$("" & matriz(LBound(matriz, 1), icol))
                If .fabricante <> "" Then .fabricante = normaliza_codigo(.fabricante, TAM_MIN_FABRICANTE)
                
            '   COL "B" = PRODUTO
                icol = icol + 1
                .produto = UCase$(Trim$("" & matriz(LBound(matriz, 1), icol)))
                If .produto <> "" Then .produto = normaliza_codigo(.produto, TAM_MIN_PRODUTO)
                
                If .produto = "" Then
                    xls_linhas_em_branco_consecutivas = xls_linhas_em_branco_consecutivas + 1
                    If xls_linhas_em_branco_consecutivas > MAX_XLS_LINHAS_EM_BRANCO_CONSECUTIVAS Then Exit For
                Else
                    xls_linhas_em_branco_consecutivas = 0
                
                '   COL "C" = DESCRIÇÃO
                    icol = icol + 1
                    .descricao = Trim$("" & matriz(LBound(matriz, 1), icol))
                    
                '   Armazena a descrição com formatação
                    .descricao_html = ""
                    blnNegritoAtivado = False
                    blnFecharTagNegrito = False
                    blnItalicoAtivado = False
                    blnFecharTagItalico = False
                    blnSublinhadoAtivado = False
                    blnFecharTagSublinhado = False
                    Set oDescricao = oWS.Range(strColunaDescricao & CStr(ilinha) & ":" & strColunaDescricao & CStr(ilinha))
                    For intCounterAux = 1 To Len(oDescricao.Value)
                        strCaracter = Mid$(oDescricao.Value, intCounterAux, 1)
                        
                    '   NEGRITO
                        If oDescricao.Characters(Start:=intCounterAux, Length:=1).Font.Bold Then
                            If Not blnNegritoAtivado Then
                                strCaracter = "<b>" & strCaracter
                                blnNegritoAtivado = True
                                End If
                        Else
                            If blnNegritoAtivado Then
                                blnFecharTagNegrito = True
                                blnNegritoAtivado = False
                                End If
                            End If
                        
                    '   ITÁLICO
                        If oDescricao.Characters(Start:=intCounterAux, Length:=1).Font.Italic Then
                            If Not blnItalicoAtivado Then
                                strCaracter = "<i>" & strCaracter
                                blnItalicoAtivado = True
                                End If
                        Else
                            If blnItalicoAtivado Then
                                blnFecharTagItalico = True
                                blnItalicoAtivado = False
                                End If
                            End If
                        
                    '   SUBLINHADO
                        If oDescricao.Characters(Start:=intCounterAux, Length:=1).Font.Underline = xlUnderlineStyleSingle Then
                            If Not blnSublinhadoAtivado Then
                                strCaracter = "<u>" & strCaracter
                                blnSublinhadoAtivado = True
                                End If
                        Else
                            If blnSublinhadoAtivado Then
                                blnFecharTagSublinhado = True
                                blnSublinhadoAtivado = False
                                End If
                            End If
                        
                    '   TAG DE FECHAMENTO (ORDEM INVERSA DA ABERTURA)
                        If blnFecharTagSublinhado Then
                            strCaracter = "</u>" & strCaracter
                            blnFecharTagSublinhado = False
                            End If
                            
                        If blnFecharTagItalico Then
                            strCaracter = "</i>" & strCaracter
                            blnFecharTagItalico = False
                            End If
                        
                        If blnFecharTagNegrito Then
                            strCaracter = "</b>" & strCaracter
                            blnFecharTagNegrito = False
                            End If
                            
                        .descricao_html = .descricao_html & strCaracter
                        Next
                
                '   TAG DE FECHAMENTO (ORDEM INVERSA DA ABERTURA)
                    If blnNegritoAtivado Then
                        .descricao_html = .descricao_html & "</b>"
                        End If
                
                    If blnItalicoAtivado Then
                        .descricao_html = .descricao_html & "</i>"
                        End If
                
                    If blnSublinhadoAtivado Then
                        .descricao_html = .descricao_html & "</u>"
                        End If
                    
                '   COL "D" = EAN
                    icol = icol + 1
                    .ean = Trim$("" & matriz(LBound(matriz, 1), icol))
                    
                '   COL "E" = GRUPO
                    icol = icol + 1
                    .grupo = UCase$(Trim$("" & matriz(LBound(matriz, 1), icol)))
                    
                '   COL "F" = POTÊNCIA (BTU/H)
                    icol = icol + 1
                    q = matriz(LBound(matriz, 1), icol)
                    If IsNumeric(q) Then .potencia_BTU = CLng(q) Else .potencia_BTU = 0
                    
                '   COL "G" = CICLO (F/QF)
                    icol = icol + 1
                    .ciclo = UCase$(Trim$("" & matriz(LBound(matriz, 1), icol)))
                    
                '   COL "H" = PREÇO FABRICANTE
                    icol = icol + 1
                    q = matriz(LBound(matriz, 1), icol)
                    If IsNumeric(q) Then .preco_fabricante = CCur(q) Else .preco_fabricante = 0
                                
                '   COL "I" = PESO (KG)
                    icol = icol + 1
                    q = matriz(LBound(matriz, 1), icol)
                    If IsNumeric(q) Then .peso = CDbl(q) Else .peso = 0
                    
                '   COL "J" = QTDE VOLUMES
                    icol = icol + 1
                    q = matriz(LBound(matriz, 1), icol)
                    If IsNumeric(q) Then .qtde_volumes = CLng(q) Else .qtde_volumes = 0
                    
                '   COL "K" = CUBAGEM
                    icol = icol + 1
                    q = matriz(LBound(matriz, 1), icol)
                    If IsNumeric(q) Then .cubagem = CDbl(q) Else .cubagem = 0
                    
                '   COL "L" = NCM
                    icol = icol + 1
                    .ncm = Trim$("" & matriz(LBound(matriz, 1), icol))
                    
                '   COL "M" = CST
                '   LEMBRANDO QUE NESTE CAMPO ESTÃO CONCATENADOS OS CÓDIGOS 'ORIG' E 'CST'
                    icol = icol + 1
                    .cst = Trim$("" & matriz(LBound(matriz, 1), icol))
                    
                '   COL "N" = POSIÇÃO MERCADO (BÁSICO/PREMIUM)
                    icol = icol + 1
                    .posicao_mercado = UCase$(Trim$("" & matriz(LBound(matriz, 1), icol)))
                    
                '   COL "O" = PERCENTUAL MVA ST
                    icol = icol + 1
                    q = matriz(LBound(matriz, 1), icol)
                    If IsNumeric(q) Then .perc_MVA_ST = CDbl(q) Else .perc_MVA_ST = 0
                    .perc_MVA_ST = .perc_MVA_ST * 100
                    
                '   COL "P" = CUSTO2
                    icol = icol + 1
                    q = matriz(LBound(matriz, 1), icol)
                    If IsNumeric(q) Then .vl_custo2 = CCur(q) Else .vl_custo2 = 0
                    
                '   COL "Q" = ALERTAS
                    icol = icol + 1
                    .alertas = Trim$("" & matriz(LBound(matriz, 1), icol))
                                
                '   COL "R" = ESTOQUE CRÍTICO
                    icol = icol + 1
                    q = matriz(LBound(matriz, 1), icol)
                    If IsNumeric(q) Then .estoque_critico = CLng(q) Else .estoque_critico = 0
                    
                '   COL "S" = DESCONTINUADO
                    icol = icol + 1
                    .descontinuado = Trim$("" & matriz(LBound(matriz, 1), icol))
                                        
                    .dt_cadastro = Null
                    .dt_ult_atualizacao = Null
                
                '   CONSISTÊNCIA
                    If Len(.produto) > TAM_MAX_PRODUTO Then
                        msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                  vbCrLf & "Produto " & .produto & " excede o tamanho máximo para o código!!"
                        GoTo TPC_ABORTA_PROCESSAMENTO
                        End If
                        
                    If .descricao = "" Then
                        msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                  vbCrLf & "Produto " & .produto & " não possui descrição!!"
                        GoTo TPC_ABORTA_PROCESSAMENTO
                        End If
                    
                    If (Len(.descricao) > 120) Then
                        msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                  vbCrLf & "Produto " & .produto & " possui descrição que excede o tamanho máximo!!"
                        GoTo TPC_ABORTA_PROCESSAMENTO
                        End If
                                            
                    If (Len(.ean) <> 0) And (Len(.ean) <> 8) And (Len(.ean) <> 12) And (Len(.ean) <> 13) And (Len(.ean) <> 14) Then
                        msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                  vbCrLf & "Produto " & .produto & " especifica código EAN (" & .ean & ") com tamanho inválido!!"
                        GoTo TPC_ABORTA_PROCESSAMENTO
                        End If
                        
                '   CÓDIGO EAN JÁ ESTÁ EM USO POR OUTRO PRODUTO ?
                    If Len(.ean) > 0 Then
                        s = "SELECT" & _
                                " fabricante," & _
                                " produto," & _
                                " ean" & _
                            " FROM " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO" & _
                            " WHERE" & _
                                " (ean='" & .ean & "')" & _
                                " AND (fabricante<>'" & .fabricante & "')"
                        If rs.State <> adStateClosed Then rs.Close
                        rs.Open s, dbc, , , adCmdText
                        If Not rs.EOF Then
                            msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                        vbCrLf & "Produto (" & .fabricante & ")" & .produto & " da planilha especifica código EAN (" & .ean & ") que já está sendo usado no banco de dados por produto de outro fabricante (" & Trim$("" & rs("fabricante")) & ")" & Trim$("" & rs("produto")) & "!!"
                            GoTo TPC_ABORTA_PROCESSAMENTO
                            End If
                        End If
                        
                    If (Len(.ncm) > 8) Then
                        msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                  vbCrLf & "Produto " & .produto & " especifica NCM (" & .ncm & ") que excede o tamanho máximo!!"
                        GoTo TPC_ABORTA_PROCESSAMENTO
                        End If
                        
                    If (Len(.grupo) > 2) Then
                        msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                  vbCrLf & "Produto " & .produto & " especifica grupo (" & .grupo & ") que excede o tamanho máximo!!"
                        GoTo TPC_ABORTA_PROCESSAMENTO
                        End If
                        
                    If .fabricante = "" Then
                        msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                  vbCrLf & "Produto " & .produto & " não especifica o código do fabricante!!"
                        GoTo TPC_ABORTA_PROCESSAMENTO
                        End If
                    
                    s = DESCRICAO_retorna(v_fabricante(), .fabricante, cadastrado)
                    If Not cadastrado Then
                        msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                  vbCrLf & "Produto " & .produto & " especifica fabricante não cadastrado: " & .fabricante & "!!"
                        GoTo TPC_ABORTA_PROCESSAMENTO
                        End If
                    
                    If .perc_MVA_ST <> 0 Then
                    '   LEMBRANDO QUE O CAMPO NA PLANILHA CONCATENA OS CAMPOS 'ORIG' E 'CST'
                        strCST = Right$(Trim$(.cst), 2)
                        If (strCST <> "10") And _
                           (strCST <> "30") And _
                           (strCST <> "70") And _
                           (strCST <> "90") Then
                            msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                      vbCrLf & "Produto " & .produto & " possui percentual de margem de valor adicionado do ICMS ST (" & formata_perc(.perc_MVA_ST) & "%) mas o código CST (" & strCST & ") é incompatível com cobrança do ICMS por substituição tributária!!"
                            GoTo TPC_ABORTA_PROCESSAMENTO
                            End If
                        End If
                        
                    If .descontinuado <> "" Then
                        If UCase$(.descontinuado) <> "S" And UCase$(.descontinuado) <> "N" Then
                            msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                      vbCrLf & "Produto " & .produto & " possui informação inválida na coluna 'descontinuado': " & .descontinuado
                            GoTo TPC_ABORTA_PROCESSAMENTO
                            End If
                        End If
                    
                    
                '   GRAVA NO BD!!
                    n_t_produto = n_t_produto + 1
                    s = "SELECT " & _
                            "*" & _
                        " FROM " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO" & _
                        " WHERE" & _
                            " (fabricante='" & .fabricante & "')" & _
                            " AND (produto='" & .produto & "')"
                    If rs.State <> adStateClosed Then rs.Close
                    rs.Open s, dbc, , , adCmdText
                    If rs.EOF Then
                        rs.AddNew
                        rs("fabricante") = .fabricante
                        rs("produto") = .produto
                        rs("dt_cadastro") = Date
                        End If
                    
                    rs("descricao") = .descricao
                    rs("descricao_html") = .descricao_html
                    rs("ean") = .ean
                    rs("grupo") = .grupo
                    rs("preco_fabricante") = .preco_fabricante
                    rs("vl_custo2") = .vl_custo2
                    rs("estoque_critico") = .estoque_critico
                    rs("peso") = .peso
                    rs("qtde_volumes") = .qtde_volumes
                    rs("cubagem") = .cubagem
                    rs("ncm") = .ncm
                    rs("cst") = .cst
                    rs("perc_MVA_ST") = .perc_MVA_ST
                    rs("descontinuado") = UCase$(.descontinuado)
                    rs("potencia_BTU") = .potencia_BTU
                    rs("ciclo") = .ciclo
                    rs("posicao_mercado") = .posicao_mercado
                    rs("dt_ult_atualizacao") = Date
                    rs("excluido_status") = 0
                    rs.Update
                    
                '   O SEPARADOR PADRÃO É VÍRGULA, SE DIGITARAM PONTO E VÍRGULA, SUBSTITUI POR VÍRGULA
                    .alertas = Replace$(.alertas, ";", ",")
                    v_alerta = Split(.alertas, ",")
                    For intCounterAux = LBound(v_alerta) To UBound(v_alerta)
                        If Trim("" & v_alerta(intCounterAux)) <> "" Then
                            s = "SELECT * FROM t_ALERTA_PRODUTO WHERE apelido = '" & Trim(v_alerta(intCounterAux)) & "'"
                            If rs.State <> adStateClosed Then rs.Close
                            rs.Open s, dbc, , , adCmdText
                            If rs.EOF Then
                                msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                          vbCrLf & "Produto " & .produto & " especifica uma mensagem de alerta que NÃO está cadastrada ('" & Trim(v_alerta(intCounterAux)) & "')"
                                GoTo TPC_ABORTA_PROCESSAMENTO
                                End If
                            
                            s = "SELECT " & _
                                    "*" & _
                                " FROM " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_X_ALERTA" & _
                                " WHERE" & _
                                    " (fabricante = '" & .fabricante & "')" & _
                                    " AND (produto = '" & .produto & "')" & _
                                    " AND (id_alerta = '" & Trim(v_alerta(intCounterAux)) & "')"
                            If rs.State <> adStateClosed Then rs.Close
                            rs.Open s, dbc, , , adCmdText
                            If rs.EOF Then
                                rs.AddNew
                                rs("fabricante") = .fabricante
                                rs("produto") = .produto
                                rs("id_alerta") = Trim(v_alerta(intCounterAux))
                                rs("dt_cadastro") = Now
                                rs("usuario_cadastro") = usuario.id
                                End If
                            
                            rs("excluido_status") = 0
                            rs.Update
                            End If
                        Next
                    End If
                End With
            Next
        
    '   REMOVE DA TABELA DE LOJAS OS PRODUTOS EXCLUÍDOS DA TABELA BÁSICA
    '   SUBSÍDIOS: A TABELA BÁSICA (T_PRODUTO) SEMPRE MANTÉM OS REGISTROS, MESMO QUE
    '              NA PLANILHA TENHA SIDO ELIMINADA.  ISSO PARA MANTER O VÍNCULO DAS
    '              OUTRAS TABELAS QUANDO FOREM NECESSÁRIOS DADOS COMO DESCRIÇÃO, ETC.
    '              SE UM PRODUTO FOI REMOVIDO DA TABELA BÁSICA (T_PRODUTO), ENTÃO ELE
    '              SERÁ REMOVIDO EM TODAS AS LOJAS (T_PRODUTO_LOJA).
        aguarde INFO_EXECUTANDO, "finalizando processamento da tabela básica"
        s = "UPDATE " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA" & _
            " SET" & _
                " excluido_status=0" & _
            " FROM " & _
                PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA tPL" & _
                " INNER JOIN " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO tP ON ((tPL.fabricante=tP.fabricante) AND (tPL.produto=tP.produto))" & _
            " WHERE" & _
                " (tP.excluido_status=0)"
        dbc.Execute s, lngRecordsAffected
        
        s = "DELETE FROM " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA WHERE (excluido_status <> 0)"
        dbc.Execute s, lngRecordsAffected
            
        s = "DELETE FROM " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_X_ALERTA WHERE (excluido_status <> 0)"
        dbc.Execute s, lngRecordsAffected
            
    '   INFORMAÇÕES PARA O LOG
        If s_log <> "" Then s_log = s_log & "; "
        s_log = s_log & "tabela básica: " & CStr(n_t_produto) & " produtos"
        
    '   CONTROLA TABELAS PROCESSADAS
        processou_tabela_basica = True
        End If
        
       
        

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' PROCESSA A TABELA DE PRODUTOS DE CADA LOJA
    
'   CARREGA LISTA DE PRODUTOS (T_PRODUTO)
    aguarde INFO_EXECUTANDO, "consultando lista de produtos"
    ReDim v_produto(0)
    With v_produto(UBound(v_produto))
        .codigo = ""
        .descricao = ""
        End With
    
    s = "SELECT" & _
            " fabricante," & _
            " produto," & _
            " descricao" & _
        " FROM " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO" & _
        " WHERE" & _
            " (excluido_status=0)" & _
        " ORDER BY" & _
            " fabricante," & _
            " produto"
    If rs.State <> adStateClosed Then rs.Close
    rs.Open s, dbc, , , adCmdText
    
    i = 0
    Do While Not rs.EOF
        i = i + 1
        ReDim Preserve v_produto(i)
        With v_produto(i)
            s = normaliza_codigo(Trim$("" & rs("fabricante")), TAM_MAX_FABRICANTE) & _
                "|" & _
                normaliza_codigo(Trim$("" & rs("produto")), TAM_MAX_PRODUTO)
            .codigo = s
            .descricao = Trim$("" & rs("descricao"))
            End With
        rs.MoveNext
        Loop
        
    ORDENA_codigo_X_descricao v_produto(), 1, UBound(v_produto)
    
    
'   LAÇO PARA PROCESSAR CADA UMA DAS PLANILHAS (CADA PLANILHA SE REFERE A UMA LOJA OU A UM GRUPO DE LOJAS)
    For ic = 1 To oXL.Workbooks(1).Worksheets.Count
    '   PULA A TABELA BÁSICA
        s_nome_planilha = UCase$(Trim$(oXL.Workbooks(1).Worksheets(ic).Name))
        
    '   ESTA PLANILHA FOI SELECIONADA ?
        foi_selecionada = False
        For i = LBound(v_xls_selecionada) To UBound(v_xls_selecionada)
            s = Trim$(v_xls_selecionada(i))
            If s <> "" Then
                s = Trim$(separa_campo(s, SEPARADOR_CODIGO))
                If s = s_nome_planilha Then
                    foi_selecionada = True
                    Exit For
                    End If
                End If
            Next
        
        If foi_selecionada And (ic <> idx_basica) And (Not IsTabelaBasicaProdutos(s_nome_planilha)) Then
            ReDim v_loja(0)
            v_loja(UBound(v_loja)) = ""
            existe_loja = False
            
            Set oWS = oXL.Workbooks(1).Worksheets(ic)
                                    
            If IsNumeroLoja(s_nome_planilha) Then
                s_destino_sem_prefixo = normaliza_codigo(remove_prefixo_do_numero(s_nome_planilha, PREFIXO_NUMERO_LOJA), TAM_MIN_LOJA)
                s_destino_com_prefixo = PREFIXO_NUMERO_LOJA & s_destino_sem_prefixo
                s = "iniciando transferência de dados da loja " & s_destino_sem_prefixo
                aguarde INFO_EXECUTANDO, s
                If s_destino_sem_prefixo <> "" Then existe_loja = True
                v_loja(UBound(v_loja)) = s_destino_sem_prefixo
            ElseIf IsNumeroGrupodeLojas(s_nome_planilha) Then
                s_destino_sem_prefixo = normaliza_codigo(remove_prefixo_do_numero(s_nome_planilha, PREFIXO_NUMERO_GRUPO_LOJAS), TAM_MIN_GRUPO_LOJAS)
                s_destino_com_prefixo = PREFIXO_NUMERO_GRUPO_LOJAS & s_destino_sem_prefixo
                s = "iniciando transferência de dados do grupo de lojas " & s_destino_com_prefixo
                aguarde INFO_EXECUTANDO, s
                s = "SELECT" & _
                        " loja" & _
                    " FROM t_LOJA_GRUPO_ITEM" & _
                    " WHERE" & _
                        " (grupo = '" & s_destino_sem_prefixo & "')" & _
                    " ORDER BY" & _
                        " CONVERT(smallint, loja)"
                If rs.State <> adStateClosed Then rs.Close
                rs.Open s, dbc, , , adCmdText
                Do While Not rs.EOF
                    If Trim$(v_loja(UBound(v_loja))) <> "" Then
                        ReDim Preserve v_loja(UBound(v_loja) + 1)
                        v_loja(UBound(v_loja)) = ""
                        End If
                    s = Trim$("" & rs("loja"))
                    s = normaliza_codigo(s, TAM_MIN_LOJA)
                    If s <> "" Then existe_loja = True
                    v_loja(UBound(v_loja)) = s
                    rs.MoveNext
                    Loop
                End If
            
            
            If existe_loja Then
                n_t_produto_loja = 0
                xls_linhas_em_branco_consecutivas = 0
                
            '   LOJA ESTÁ CADASTRADA ?
                For il = LBound(v_loja) To UBound(v_loja)
                    If Trim$(v_loja(il)) <> "" Then
                        s = "SELECT loja FROM t_LOJA WHERE (loja = '" & Trim$(v_loja(il)) & "')"
                        If rs.State <> adStateClosed Then rs.Close
                        rs.Open s, dbc, , , adCmdText
                        If rs.EOF Then
                            msg_erro = "Planilha " & oWS.Name & _
                                       vbCrLf & "Loja " & Trim$(v_loja(il)) & " não está cadastrada!!"
                            GoTo TPC_ABORTA_PROCESSAMENTO
                            End If
                        End If
                    Next
                                
            
            '   GRUPO DE LOJAS: CARREGA OS DADOS DA PLANILHA P/ APENAS UMA LOJA
            '   =============== E DEPOIS EXECUTA COMANDOS SQL P/ FAZER A REPLICAÇÃO P/ AS DEMAIS.
                s_loja_base_replicacao = ""
                For il = UBound(v_loja) To LBound(v_loja) Step -1
                    If Trim$(v_loja(il)) <> "" Then
                        s_loja_base_replicacao = Trim$(v_loja(il))
                        Exit For
                        End If
                    Next
            
            '   LIMPA RELAÇÃO DE PRODUTOS ANTES DE GRAVAR OS NOVOS DADOS
                s = sql_monta_criterio_texto_or(v_loja(), "loja", True)
                If s <> "" Then
                    s = "DELETE FROM " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA WHERE (" & s & ")"
                    dbc.Execute s, lngRecordsAffected
                    End If
            
                                
                n_reg = 0
                For ilinha = 2 To MAX_LINHAS_PLANILHA
                    
                '   CANCELOU ?
                    DoEvents
                    If f_PROCESSANDO.processo_cancelado Then
                        msg_erro = "Operação cancelada!!"
                        GoTo TPC_ABORTA_PROCESSAMENTO
                        End If
                        
                '   PROGRESSO
                    If (n_reg Mod 10) = 0 Then
                        If IsNumeroGrupodeLojas(s_nome_planilha) Then
                            s = "transferindo dados do grupo de lojas " & s_destino_com_prefixo & _
                                ": " & Format$(n_reg, FORMATO_INTEIRO) & " regs"
                        Else
                            s = "transferindo dados da loja " & s_destino_sem_prefixo & _
                                ": " & Format$(n_reg, FORMATO_INTEIRO) & " regs"
                            End If
                        aguarde INFO_EXECUTANDO, s
                        End If
                                                
                                                
                '   PROCESSA REGISTRO
                    n_reg = n_reg + 1
                    
                    strPrimeiraColunaPlanilha = "A"
                    strUltimaColunaPlanilha = "G"
                    Set oRANGE = oWS.Range(strPrimeiraColunaPlanilha & CStr(ilinha) & ":" & strUltimaColunaPlanilha & CStr(ilinha))
                    matriz = oRANGE.Value
                    
                    With rpl
                        icol = LBound(matriz, 2) - 1
                        
                    '   LÊ A LINHA
                        icol = icol + 1
                        .fabricante = Trim$("" & matriz(LBound(matriz, 1), icol))
                        If .fabricante <> "" Then .fabricante = normaliza_codigo(.fabricante, TAM_MIN_FABRICANTE)
                        
                        icol = icol + 1
                        .produto = UCase$(Trim$("" & matriz(LBound(matriz, 1), icol)))
                        If .produto <> "" Then .produto = normaliza_codigo(.produto, TAM_MIN_PRODUTO)
                        
                        If .produto = "" Then
                            xls_linhas_em_branco_consecutivas = xls_linhas_em_branco_consecutivas + 1
                            If xls_linhas_em_branco_consecutivas > MAX_XLS_LINHAS_EM_BRANCO_CONSECUTIVAS Then Exit For
                        Else
                            xls_linhas_em_branco_consecutivas = 0
                            
                            icol = icol + 1
                            q = matriz(LBound(matriz, 1), icol)
                            If IsNumeric(q) Then .preco_lista = arredonda_para_monetario(CCur(q)) Else .preco_lista = 0
                            
                            .margem = 0
                            
                            icol = icol + 1
                            q = matriz(LBound(matriz, 1), icol)
                            If IsNumeric(q) Then .desc_max = CDbl(q) Else .desc_max = 0
                            .desc_max = .desc_max * 100
                            
                            icol = icol + 1
                            q = matriz(LBound(matriz, 1), icol)
                            If IsNumeric(q) Then .comissao = CDbl(q) Else .comissao = 0
                            .comissao = .comissao * 100
                            
                            icol = icol + 1
                            .vendavel = UCase$(Trim$("" & matriz(LBound(matriz, 1), icol)))
                            
                            icol = icol + 1
                            q = matriz(LBound(matriz, 1), icol)
                            If IsNumeric(q) Then .qtde_max_venda = CLng(q) Else .qtde_max_venda = 0
                            
                            icol = icol + 1
                        '   SE A COR FOR BRANCA (DEFAULT NO EXCEL), TROCA PARA PRETO (COR DEFAULT DE FONTE)
                            icor = oWS.Cells(ilinha, icol).Interior.Color
                            If icor = COR_BRANCO Then icor = COR_PRETO
                            .cor = "#" & cor_em_decimal_para_rgb(icor)
                            
                            .dt_cadastro = Null
                            .dt_ult_atualizacao = Null
                        
                        '   CONSISTÊNCIA
                            If .fabricante = "" Then
                                msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                          vbCrLf & "Produto " & .produto & " não especifica o código do fabricante!!"
                                GoTo TPC_ABORTA_PROCESSAMENTO
                                End If
                            
                            s = DESCRICAO_retorna(v_fabricante(), .fabricante, cadastrado)
                            If Not cadastrado Then
                                msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                          vbCrLf & "Produto " & .produto & " especifica fabricante não cadastrado: " & .fabricante & "!!"
                                GoTo TPC_ABORTA_PROCESSAMENTO
                                End If
                            
                            If (.preco_lista < 0) And (.preco_lista <> VL_COD_PRECO_A_IGNORAR) Then
                                msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                          vbCrLf & "Produto " & .produto & " especifica preço de lista inválido: " & _
                                          Format$(.preco_lista, FORMATO_MOEDA) & "!!"
                                GoTo TPC_ABORTA_PROCESSAMENTO
                                End If
                            
                            If .preco_lista = VL_COD_PRECO_A_IGNORAR Then
                            '   SE FOR UM PRODUTO NOVO, A PRIMEIRA CARGA DEVE INFORMAR UM PREÇO, OU SEJA, A 1ª CARGA NÃO PODE USAR O CÓDIGO ESPECIAL -1,00
                                s = "SELECT" & _
                                        " preco_lista" & _
                                    " FROM t_PRODUTO_LOJA" & _
                                    " WHERE" & _
                                        " (fabricante = '" & .fabricante & "')" & _
                                        " AND (produto = '" & .produto & "')" & _
                                        " AND (loja = '" & s_loja_base_replicacao & "')"
                                If tPL.State <> adStateClosed Then tPL.Close
                                tPL.Open s, dbc, , , adCmdText
                                If tPL.EOF Then
                                    msg_erro = "Planilha " & oWS.Name & ", linha " & CStr(ilinha) & _
                                              vbCrLf & "Produto " & .produto & " especifica o preço de lista " & _
                                              Format$(.preco_lista, FORMATO_MOEDA) & " usado p/ indicar que o preço não deve ser atualizado, entretanto, o produto ainda NÃO está cadastrado na tabela de preços no sistema!!" & _
                                              vbCrLf & "O cadastramento de um novo produto sempre deve informar um preço válido na primeira carga!!"
                                    GoTo TPC_ABORTA_PROCESSAMENTO
                                    End If
                                End If
                                
                            s = normaliza_codigo(.fabricante, TAM_MAX_FABRICANTE) & _
                                "|" & _
                                normaliza_codigo(.produto, TAM_MAX_PRODUTO)
                            s = DESCRICAO_retorna(v_produto(), s, cadastrado)
                            If cadastrado Then
                            '   GRAVA NO BD!!
                                n_t_produto_loja = n_t_produto_loja + 1
                                
                            '   NO CASO DE GRUPO DE LOJAS, CARREGA A TABELA P/ A ÚLTIMA LOJA E DEPOIS REPLICA P/ AS DEMAIS
                                s = "INSERT INTO " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA ("
                                s = s & "fabricante, " & _
                                        "produto, " & _
                                        "loja, " & _
                                        "dt_cadastro, " & _
                                        "preco_lista, " & _
                                        "margem, " & _
                                        "desc_max, " & _
                                        "comissao, " & _
                                        "vendavel, " & _
                                        "qtde_max_venda, " & _
                                        "cor, " & _
                                        "dt_ult_atualizacao, " & _
                                        "excluido_status"
                                s = s & ") VALUES ("
                                s = s & "'" & .fabricante & "', " & _
                                        "'" & .produto & "', " & _
                                        "'" & s_loja_base_replicacao & "', " & _
                                        bd_monta_data(Now, False) & ", " & _
                                        bd_monta_numero(.preco_lista) & ", " & _
                                        bd_monta_numero(.margem) & ", " & _
                                        bd_monta_numero(.desc_max) & ", " & _
                                        bd_monta_numero(.comissao) & ", " & _
                                        "'" & .vendavel & "', " & _
                                        bd_monta_numero(.qtde_max_venda) & ", " & _
                                        "'" & .cor & "', " & _
                                        bd_monta_data(Now, False) & ", " & _
                                        "0" & _
                                        ")"
                                
                                dbc.Execute s, lngRecordsAffected
                                End If
                            End If
                        End With
                    Next 'For ilinha = 2 To MAX_LINHAS_PLANILHA
                
            '   GRUPO DE LOJAS: REPLICA A TABELA DA LOJA QUE FOI CARREGADA
                For il = LBound(v_loja) To UBound(v_loja)
                    If (Trim$(v_loja(il)) <> "") And (Trim$(v_loja(il)) <> s_loja_base_replicacao) Then
                    '   PROGRESSO
                        s = "replicando dados do grupo de lojas " & s_destino_com_prefixo & _
                            ": loja " & Trim$(v_loja(il))
                        aguarde INFO_EXECUTANDO, s
                        
                        s = "INSERT INTO " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA ("
                        s = s & "fabricante, " & _
                                "produto, " & _
                                "loja, " & _
                                "dt_cadastro, " & _
                                "preco_lista, " & _
                                "margem, " & _
                                "desc_max, " & _
                                "comissao, " & _
                                "vendavel, " & _
                                "qtde_max_venda, " & _
                                "cor, " & _
                                "dt_ult_atualizacao, " & _
                                "excluido_status" & _
                            ")"
                        s = s & " SELECT " & _
                                "fabricante, " & _
                                "produto, " & _
                                "'" & Trim$(v_loja(il)) & "', " & _
                                "dt_cadastro, " & _
                                "preco_lista, " & _
                                "margem, " & _
                                "desc_max, " & _
                                "comissao, " & _
                                "vendavel, " & _
                                "qtde_max_venda, " & _
                                "cor, " & _
                                "dt_ult_atualizacao, " & _
                                "excluido_status" & _
                            " FROM " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA" & _
                            " WHERE" & _
                                " (loja = '" & s_loja_base_replicacao & "')"

                        dbc.Execute s, lngRecordsAffected
                        End If
                    Next 'For il = LBound(v_loja) To UBound(v_loja)
                    
                
            '   ATUALIZA O PREÇO DE TABELA C/ O PREÇO ORIGINAL P/ OS PRODUTOS QUE ESTÃO C/ O CÓDIGO ESPECIAL VL_COD_PRECO_A_IGNORAR
            '   IMPORTANTE: ESTA ETAPA DEVE SER REALIZADA APÓS A REPLICAÇÃO DA TABELA DE PREÇOS P/ AS OUTRAS LOJAS, POIS O PROCESSAMENTO
            '   DO MÓDULO ConsolidadorXlsEC PODE TER ALTERADO O PREÇO DA TABELA DE PREÇOS DA LOJA DO E-COMMERCE, PORTANTO, AS TABELAS
            '   DE PREÇOS PODEM ESTAR DIFERENTES ENTRE AS LOJAS DO MESMO GRUPO.
            '   ASSIM, O PROCESSAMENTO DO CÓDIGO ESPECIAL VL_COD_PRECO_A_IGNORAR DEVE SER FEITO USANDO O PREÇO ORIGINAL DE CADA LOJA.
                s = "UPDATE tmp__PL" & _
                        " SET tmp__PL.preco_lista = tPL.preco_lista" & _
                    " FROM " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA tmp__PL" & _
                        " INNER JOIN t_PRODUTO_LOJA tPL ON (tmp__PL.fabricante = tPL.fabricante) AND (tmp__PL.produto = tPL.produto) AND (tmp__PL.loja = tPL.loja)" & _
                    " WHERE" & _
                        " (tmp__PL.preco_lista = " & bd_monta_moeda(VL_COD_PRECO_A_IGNORAR) & ")"
                dbc.Execute s, lngRecordsAffected
                
            '   SE RESTOU ALGUM PRODUTO PREÇO C/ O VALOR ESPECIAL VL_COD_PRECO_A_IGNORAR, COLOCA O PRODUTO COMO NÃO-VENDÁVEL, POIS ESSA SITUAÇÃO NÃO DEVERIA TER OCORRIDO.
            '   UMA DAS HIPÓTESES QUE PODERIAM CAUSAR ISSO, SERIA O CADASTRAMENTO DE UM NOVO PRODUTO INFORMANDO O VALOR DO CÓDIGO ESPECIAL.
            '   PARA A LOJA DO E-COMMERCE, A PRIMEIRA CARGA DA TABELA DE PREÇOS DE UM PRODUTO NOVO DEVE OBRIGATORIAMENTE INFORMAR UM PREÇO VÁLIDO, POIS O MÓDULO ConsolidadorXlsEC
            '   AO ATUALIZAR A TABELA DE PREÇOS NO BD A PARTIR DOS DADOS CONSOLIDADOS DA PLANILHA DE GERENCIAMENTO DE PREÇOS IRÁ SE BASEAR NA PROPORÇÃO ENTRE OS PREÇOS DAS PARTES
            '   AO ATUALIZAR O PREÇO DE UM CÓDIGO UNIFICADO. PORTANTO, É IMPRESCINDÍVEL QUE A PRIMEIRA CARGA DE UM PRODUTO NOVO INFORME OS PREÇOS QUE TENHAM UMA PROPORÇÃO CORRETA,
            '   MESMO QUE O PREÇO TOTAL NÃO SEJA O REAL.
                s = "UPDATE " & PREFIXO_TABELA_TEMPORARIA & "t_PRODUTO_LOJA" & _
                    " SET" & _
                        " vendavel = 'N'" & _
                    " WHERE" & _
                        " (preco_lista = " & bd_monta_moeda(VL_COD_PRECO_A_IGNORAR) & ")"
                dbc.Execute s, lngRecordsAffected
                               
                               
            '   LOG E RESUMO DO PROCESSAMENTO
                If IsNumeroGrupodeLojas(s_nome_planilha) Then
                    s = "finalizando transferência de dados do grupo de lojas " & s_destino_com_prefixo
                    aguarde INFO_EXECUTANDO, s
                '   INFORMAÇÕES PARA O LOG
                    If s_log <> "" Then s_log = s_log & "; "
                    s_log = s_log & "grupo de lojas " & s_destino_com_prefixo & _
                            " (lojas: " & Join(v_loja, ", ") & "): " & _
                            CStr(n_t_produto_loja) & " produtos"
                '   CONTROLA LISTA DE LOJAS PROCESSADAS
                    s = s_destino_com_prefixo & " = lojas " & Join(v_loja, ", ")
                    If v_grupo_processado(UBound(v_grupo_processado)) <> "" Then ReDim Preserve v_grupo_processado(UBound(v_grupo_processado) + 1)
                    v_grupo_processado(UBound(v_grupo_processado)) = s
                Else
                    s = "finalizando transferência de dados da loja " & s_destino_sem_prefixo
                    aguarde INFO_EXECUTANDO, s
                '   INFORMAÇÕES PARA O LOG
                    If s_log <> "" Then s_log = s_log & "; "
                    s_log = s_log & "loja " & s_destino_sem_prefixo & _
                            ": " & CStr(n_t_produto_loja) & " produtos"
                '   CONTROLA LISTA DE LOJAS PROCESSADAS
                    If v_lj_processada(UBound(v_lj_processada)) <> "" Then ReDim Preserve v_lj_processada(UBound(v_lj_processada) + 1)
                    v_lj_processada(UBound(v_lj_processada)) = s_destino_sem_prefixo
                    End If
                End If
            End If
        Next
        
    

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' INÍCIO - TRANSAÇÃO
    dbc.BeginTrans
    On Error GoTo TPC_TRATA_ERRO_TRANSACAO  ' TRATAMENTO COM ROLLBACK
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    aguarde INFO_EXECUTANDO, "transferindo tabelas temporárias para produção"
    
'   LIMPA AS TABELAS DE PRODUÇÃO E TRANSFERE OS DADOS PROCESSADOS NAS TABELAS TEMPORÁRIAS
    strNomeTabela = "t_PRODUTO"
    s = "DELETE FROM " & strNomeTabela
    dbc.Execute s, lngRecordsAffected
    s = "INSERT INTO " & strNomeTabela & " SELECT * FROM " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela
    dbc.Execute s, lngRecordsAffected
    
    strNomeTabela = "t_PRODUTO_LOJA"
    s = "DELETE FROM " & strNomeTabela
    dbc.Execute s, lngRecordsAffected
    s = "INSERT INTO " & strNomeTabela & " SELECT * FROM " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela
    dbc.Execute s, lngRecordsAffected
    
    strNomeTabela = "t_PRODUTO_X_ALERTA"
    s = "DELETE FROM " & strNomeTabela
    dbc.Execute s, lngRecordsAffected
    s = "INSERT INTO " & strNomeTabela & " SELECT * FROM " & PREFIXO_TABELA_TEMPORARIA & strNomeTabela
    dbc.Execute s, lngRecordsAffected

'   EXCLUI O VÍNCULO DO PRODUTO COM AS REGRAS DE MULTIPLOS CD'S P/ OS PRODUTOS EXCLUÍDOS
    s = "DELETE FROM t_PRODUTO_X_WMS_REGRA_CD WHERE (fabricante + '|' + produto NOT IN (SELECT fabricante + '|' + produto FROM t_PRODUTO WHERE (excluido_status = 0)))"
    dbc.Execute s, lngRecordsAffected
    
'   DURAÇÃO
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "tempo de abertura da planilha: " & Format$(tempo_planilha_open, "hh:mm:ss")
        
    tempo_transferencia = Now - hora_inicio
    If s_log <> "" Then s_log = s_log & "; "
    s_log = s_log & "tempo de transferência dos dados: " & Format$(tempo_transferencia, "hh:mm:ss")
        
'   LOG!!
    If Not grava_log(usuario.id, "", "", "", OP_LOG_CARREGA_TABELA_PRODUTOS, s_log, msg_erro) Then
        GoTo TPC_ABORTA_PROCESSAMENTO
        End If
    
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' FIM - TRANSAÇÃO
    dbc.CommitTrans
    On Error GoTo TPC_TRATA_ERRO     ' CANCELA TRATAMENTO COM ROLLBACK
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    f_PROCESSANDO.fecha_painel


  ' REMOVE TABELAS TEMPORÁRIAS
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call remove_tabelas_temporarias(PREFIXO_TABELA_TEMPORARIA, s_aux)

  ' FECHA TABELAS E OBJETOS
  ' ~~~~~~~~~~~~~~~~~~~~~~~
    GoSub TPC_FECHA_TABELAS_E_OBJETOS

    aguarde INFO_NORMAL, m_id

'   MENSAGEM DE STATUS DO PROCESSAMENTO
    s = "Operação concluída com sucesso!!" & _
        vbCrLf & vbCrLf & _
        "Tabelas processadas: "
        
    i = 0
    If processou_tabela_basica Then
        i = i + 1
        s = s & vbCrLf & vbCrLf & CStr(i) & ") Tabela Básica"
        End If
    
    s_aux = Join(v_lj_processada, ", ")
    If s_aux <> "" Then
        i = i + 1
        s_aux = CStr(i) & ") Lojas: " & s_aux
        s = s & vbCrLf & s_aux
        End If
    
    s_aux = Join(v_grupo_processado, vbCrLf & Space$(8))
    If s_aux <> "" Then
        i = i + 1
        s_aux = CStr(i) & ") Grupos de Lojas: " & vbCrLf & Space$(8) & s_aux
        s = s & vbCrLf & s_aux
        End If
    
    aviso s
    
Exit Sub







'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   ABORTA O PROCESSAMENTO DE CARGA DE DADOS DA PLANILHA
' 
TPC_ABORTA_PROCESSAMENTO:
'========================
    GoSub TPC_EXECUTA_CANCELA_TRANSACAO
    GoSub TPC_FECHA_TABELAS_E_OBJETOS
    Call remove_tabelas_temporarias(PREFIXO_TABELA_TEMPORARIA, s_aux)
    f_PROCESSANDO.fecha_painel
    aguarde INFO_NORMAL, m_id
    aviso_erro msg_erro
    Exit Sub
    
    

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   TRATAMENTO DE ERRO EM ROTINAS DENTRO DE TRANSAÇÃO
' 
TPC_TRATA_ERRO_TRANSACAO:
'========================
    s = CStr(Err) & ": " & Error$(Err)
  ' CANCELA TRANSAÇÃO ANTES DA INTERAÇÃO C/ USUÁRIO!!
    GoSub TPC_EXECUTA_CANCELA_TRANSACAO
    GoSub TPC_FECHA_TABELAS_E_OBJETOS
    Call remove_tabelas_temporarias(PREFIXO_TABELA_TEMPORARIA, s_aux)
    f_PROCESSANDO.fecha_painel
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   TRATAMENTO PARA CANCELAR TRANSAÇÃO
' 
TPC_EXECUTA_CANCELA_TRANSACAO:
'=============================
   '~~~~~~~~~~~~~~~~~~~~
    On Error Resume Next
    dbc.RollbackTrans
   '~~~~~~~~~~~~~~~~~~~~
    Return



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   TRATAMENTO PARA CANCELAMENTO DO DIALOG BOX
' 
TPC_TRATA_CANCELA_DIALOG_BOX:
'============================
    fcmd.FileName = ""
    Resume Next
    
    
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   TRATAMENTO DE ERRO
' 
TPC_TRATA_ERRO:
'==============
    s = CStr(Err) & ": " & Error$(Err)
    If Err = 429 Then s = s & vbCrLf & vbCrLf & "Certifique-se de que o MS-Excel está instalado corretamente!!"
    GoSub TPC_FECHA_TABELAS_E_OBJETOS
    f_PROCESSANDO.fecha_painel
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   TRATAMENTO DE ERRO (COM TABELAS TEMPORÁRIAS JÁ CRIADAS)
' 
TPC_TRATA_ERRO_TAB_TEMP:
'=======================
    s = CStr(Err) & ": " & Error$(Err)
    If Err = 429 Then s = s & vbCrLf & vbCrLf & "Certifique-se de que o MS-Excel está instalado corretamente!!"
    GoSub TPC_FECHA_TABELAS_E_OBJETOS
    Call remove_tabelas_temporarias(PREFIXO_TABELA_TEMPORARIA, s_aux)
    f_PROCESSANDO.fecha_painel
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
TPC_FECHA_TABELAS_E_OBJETOS:
'===========================
    bd_desaloca_recordset rs, True
    bd_desaloca_recordset tPL, True
    
    On Error Resume Next
    
    If Not (oDescricao Is Nothing) Then
        Set oDescricao = Nothing
        End If
    
    If Not (oRANGE Is Nothing) Then
        Set oRANGE = Nothing
        End If
    
    If Not (oWS Is Nothing) Then
        Set oWS = Nothing
        End If
        
    If Not (oXL Is Nothing) Then
        oXL.Quit
        Set oXL = Nothing
        End If
    
    Return


End Sub

Private Sub barra_Click(Index As Integer)

Dim s As String

    On Error GoTo BARRA_CLICK_TRATA_ERRO
    
'   EM EXECUÇÃO ?
    If em_execucao Then Exit Sub
    
    em_execucao = True
    
    Select Case Index
        Case o_CARREGA_PRODUTOS: tabela_produtos_carrega
        End Select
        
    em_execucao = False
    
Exit Sub





'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BARRA_CLICK_TRATA_ERRO:
'======================
    s = CStr(Err) & ": " & Error$(Err)
    em_execucao = False
    f_PROCESSANDO.fecha_painel
    aguarde INFO_NORMAL, m_id
    aviso_erro s
    Exit Sub
    
End Sub

Private Sub barra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

    For i = seta.LBound To seta.UBound
        If i <> Index Then
            seta(i).Visible = False
        Else
            seta(i).Visible = True
            End If
        Next
        
End Sub


Private Sub Form_Activate()

Dim s As String
Dim msg_erro As String
Dim strMensagem As String
Dim cor_inicial As String
Dim qtdEmits As Integer

    On Error GoTo FORMACTIVATE_TRATA_ERRO

'   ESTE FLAG IMPEDE QUE O EVENTO DO FORM ACTIVATE SEJA PROCESSADO NOVAMENTE.
'   O EVENTO PODE OCORRER NOVAMENTE SE, APÓS O FECHAMENTO DO PAINEL DE LOGIN,
'   FOR EXIBIDO UM AVISO ATRAVÉS DE MESSAGE BOX EM ALGUMA DAS VERIFICAÇÕES
'   FEITAS EM SEGUIDA.
'   NESTE CASO, O 2º EVENTO COMEÇA A SER PROCESSADO NO MOMENTO EM QUE A MESSAGE
'   BOX É EXIBIDA, OU SEJA, O PROCESSAMENTO DO 1º EVENTO É INTERROMPIDO NESSE PONTO.
'   APÓS CONCLUIR O PROCESSAMENTO DO 2º EVENTO, O 1º EVENTO VOLTA A SER PROCESSADO.
'   O EFEITO CAUSADO NESTE CASO É QUE A TELA DE LOGIN E AS MENSAGENS DE ALERTA
'   SÃO EXIBIDAS 2 VEZES.
'   ISTO OCORRE SOMENTE NO EXECUTÁVEL (LEMBRE-SE DE QUE EXISTEM DIFERENÇAS ENTRE
'   EXECUTAR DENTRO DO VISUAL BASIC E ATRAVÉS DO EXECUTÁVEL).
    If Not modulo_executou_inicializacao Then
        
        modulo_executou_inicializacao = True
                
        If DESENVOLVIMENTO Then
            LISTA_TABELAS_EXCLUIDAS_DO_BACKUP = LISTA_TABELAS_EXCLUIDAS_DO_BACKUP_DESENVOLVIMENTO
            Caption = Caption & " (Desenvolvimento)"
            s = "Versão exclusiva de desenvolvimento"
            s = s & vbCrLf & vbCrLf & "Continua?"
            If Not confirma(s) Then
                BD_Fecha
                '***
                 End
                '***
                End If
        Else
            'Execução sendo feita em ambiente de PRODUÇÃO !!
            LISTA_TABELAS_EXCLUIDAS_DO_BACKUP = LISTA_TABELAS_EXCLUIDAS_DO_BACKUP_PRODUCAO
        End If
    
      ' CONFIGURAÇÃO REGIONAL ESTÁ OK ?
        If Not verifica_configuracao_regional() Then
            s = "Há parâmetros da configuração regional que NÃO estão de acordo com as necessidades deste programa!!" & _
                vbCrLf & "Deseja que esses parâmetros sejam corrigidos agora ?"
            If Not confirma(s) Then
                aviso_erro "O programa não pode prosseguir enquanto a configuração regional não for corrigida!!"
               '~~~
                End
               '~~~
                End If
          
            If verifica_configuracao_regional(True) Then
                s = "A configuração regional foi alterada com sucesso!!" & _
                    vbCrLf & "O programa será encerrado agora e deve ser executado novamente para que possa operar corretamente!!"
                alerta s
            Else
                s = "Não foi possível alterar a configuração regional automaticamente!!" & _
                    vbCrLf & "Execute este programa novamente para tentar outra vez ou então faça a configuração manualmente!!"
                alerta s
                End If
                  
           '~~~
            End
           '~~~
            End If
        
      ' CONFIGURA PARÂMETROS DO CLIENT DO SQL SERVER NO REGISTRY DO WINDOWS
        If Not configura_registry_client_sql_server(msg_erro) Then
            s = "Falha ao configurar acesso do cliente do banco de dados!!" & _
                vbCrLf & "Não é possível continuar!!"
            If msg_erro <> "" Then s = s & vbCrLf & vbCrLf & msg_erro
            aviso_erro s
           '~~~
            End
           '~~~
            End If
                    
      ' LÊ PARÂMETROS P/ CONEXÃO AO BD
        If Not le_arquivo_ini(msg_erro) Then
            s = "Falha ao ler arquivo de configuração!!" & _
                vbCrLf & "Não é possível continuar!!"
            If msg_erro <> "" Then s = s & vbCrLf & vbCrLf & msg_erro
            aviso_erro s
           '~~~
            End
           '~~~
            End If
            
      ' INICIA BD
        aguarde INFO_EXECUTANDO, "conectando ao banco de dados"
        If Not BD_inicia() Then
            s = "Falha ao conectar com o Banco de Dados!!" & _
                vbCrLf & "Não é possível continuar!!"
            aviso_erro s
           '~~~
            End
           '~~~
            End If
                
    '   PARÂMETROS DA T_VERSAO
        obtem_parametros_t_versao cor_fundo_padrao, identificador_ambiente_padrao
        If cor_fundo_padrao <> "" Then
            cor_fundo_padrao = converte_cor_Web2VB(cor_fundo_padrao)
            Me.BackColor = cor_fundo_padrao
            ' SE A COR DE FUNDO DO BANCO DE DADOS É DIFERENTE, GRAVAR NO REGISTRY
            If cor_fundo_padrao <> cor_inicial Then
                If Not configura_registry_usuario_cor_fundo_padrao(converte_cor_VB2Web(cor_fundo_padrao)) Then
                    aviso "Não foi possível gravar as configurações de cor de fundo para futuros acessos!"
                    End If
                End If
        Else
            'SE EXISTIR UMA COR DE FUNDO GRAVADA NO REGISTRY, UTILIZAR
            If le_registry_usuario_cor_fundo_padrao(cor_inicial) Then
                cor_inicial = converte_cor_Web2VB(cor_inicial)
                Me.BackColor = cor_inicial
                End If
            End If
        If identificador_ambiente_padrao <> "" Then
            lblTitAmb = "Ambiente: " & identificador_ambiente_padrao
        Else
            lblTitAmb = ""
            End If
                
      ' INFORMA ONDE ESTÁ SE CONECTANDO !
        strMensagem = obtem_site_sistema
        lblTitSite = lblTitSite & strMensagem
        aviso "Você está se conectando em: " & strMensagem

    '   LOGIN
        aguarde INFO_NORMAL, m_id
        f_LOGIN.Show vbModal
        Set painel_ativo = Me
        
    '   NÍVEL DE ACESSO
        If Not usuario.perfil_acesso_ok Then
            aviso_erro "ACESSO NEGADO!!" & vbCrLf & "Você não possui o perfil de acesso necessário!!"
          ' ENCERRA O PROGRAMA
            BD_Fecha
           '~~~
            End
           '~~~
            End If
      
      
    '   SELEÇÃO DO CD A SER UTILIZADO
        If obtem_emitentes_usuario(usuario.id, vEmitsUsuario, qtdEmits) Then
            If qtdEmits = 1 Then
                usuario.emit = Mid$(vEmitsUsuario(UBound(vEmitsUsuario)).c1, 1, Len(vEmitsUsuario(UBound(vEmitsUsuario)).c1) - 5)
                usuario.emit_uf = Mid$(vEmitsUsuario(UBound(vEmitsUsuario)).c1, Len(vEmitsUsuario(UBound(vEmitsUsuario)).c1) - 2, 2)
                usuario.emit_id = vEmitsUsuario(UBound(vEmitsUsuario)).c2
                txtFixoEspecifico = vEmitsUsuario(UBound(vEmitsUsuario)).c3
            Else
                f_CD.Show vbModal, Me
                End If
        Else
            aviso_erro "Nenhum Centro de Distribuição habilitado para o usuário!!"
          ' ENCERRA O PROGRAMA
            BD_Fecha
           '~~~
            End
           '~~~
            End If

      ' OK !!
        aguarde INFO_NORMAL, m_id
        End If
    
Exit Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FORMACTIVATE_TRATA_ERRO:
'=======================
    s = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    aviso_erro s
    Exit Sub


End Sub

Private Sub BARRA_Inicia()

Dim i As Integer
  
  ' CARREGA SETAS
    For i = (barra.LBound + 1) To barra.UBound
        Load seta(i)
        seta(i).Picture = seta(0).Picture
        Next
    
    For i = barra.LBound To barra.UBound
        If barra(i).Visible Then
            seta(i).Visible = False
            barra(i).BackColor = COR_BARRA
            barra_desenha_borda_3dc barra(i), 1
            seta(i).Move barra(i).Left + 3, barra(i).Top + 1
            seta(i).ZOrder
            End If
        Next
    
End Sub


Sub barra_desenha_borda_3dc(c As Label, l)

Dim claro As Long
Dim escuro As Long
Dim f As Form

    Set f = c.Parent
    
' DESENHA AS SOMBRAS
' ~~~~~~~~~~~~~~~~~~
    If l >= 0 Then
        claro = cor_nova_tonalidade(c.BackColor, FATOR_COR_SOMBREAMENTO)
        escuro = cor_nova_tonalidade(c.BackColor, -FATOR_COR_SOMBREAMENTO)
    Else
        claro = cor_nova_tonalidade(c.BackColor, -FATOR_COR_SOMBREAMENTO)
        escuro = cor_nova_tonalidade(c.BackColor, FATOR_COR_SOMBREAMENTO)
        End If
    
    
    f.ScaleMode = vbPixels
    
    f.Line (c.Left, c.Top)-(c.Left + c.Width, c.Top + c.Height), c.BackColor, BF
    
    f.Line (c.Left - 1, c.Top - 1)-(c.Left + c.Width, c.Top - 1), claro
    f.Line (c.Left - 1, c.Top - 1)-(c.Left - 1, c.Top + c.Height), claro
    f.Line (c.Left + c.Width, c.Top)-(c.Left + c.Width, c.Top + c.Height + 1), escuro
    f.Line (c.Left - 1, c.Top + c.Height)-(c.Left + c.Width, c.Top + c.Height), escuro

    If l < 0 Then
        f.Line (c.Left + c.Width + 1, c.Top - 1)-(c.Left + c.Width + 1, c.Top + c.Height + 1), COR_CINZA_CLARO
        f.Line (c.Left - 1, c.Top + c.Height + 1)-(c.Left + c.Width + 1, c.Top + c.Height + 1), COR_CINZA_CLARO
    Else
        f.Line (c.Left + c.Width + 1, c.Top - 1)-(c.Left + c.Width + 1, c.Top + c.Height + 1), f.BackColor
        f.Line (c.Left - 1, c.Top + c.Height + 1)-(c.Left + c.Width + 1, c.Top + c.Height + 1), f.BackColor
        End If
        
        
End Sub


Private Sub Form_Load()

Dim s As String
    
    On Error GoTo FORMLOAD_TRATA_ERRO
    
    Set painel_ativo = Me
    Set painel_principal = Me
        
    b_dummy.Top = -500
    
    ScaleMode = vbPixels
    
    BARRA_Inicia
    
    INFO_Inicia Me

Exit Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FORMLOAD_TRATA_ERRO:
'===================
    s = CStr(Err) & ": " & Error$(Err)
    Err.Clear
    aviso_erro s
    Exit Sub

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    
    For i = seta.LBound To seta.UBound
        seta(i).Visible = False
        Next
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

'   EM EXECUÇÃO ?
    If em_execucao Then
        Cancel = True
        Exit Sub
        End If

'   FECHA BANCO DE DADOS
    BD_Fecha
    End
    
End Sub

Private Sub relogio_Timer()

Dim s As String
Dim t As String

' ATUALIZAÇÃO DO RELÓGIO
' ~~~~~~~~~~~~~~~~~~~~~~
    t = Time$

    s = Left$(t, 5)
    If Val(Right$(t, 1)) Mod 2 Then Mid$(s, 3, 1) = " "
    agora = s
    
    s = Format$(Date, "dd/mm/yyyy")
    s = Mid$(s, 1, 2) & "." & obtem_mes_em_portugues(Mid(s, 4, 2), True) & "." & Mid$(s, 7, 4)
    hoje = s


End Sub



