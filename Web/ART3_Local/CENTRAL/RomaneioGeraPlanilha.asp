<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R O M A N E I O G E R A P L A N I L H A . A S P
'     =================================================================
'
'
'	  S E R V E R   S I D E   S C R I P T I N G
'
'      SSSSSSS   EEEEEEEEE  RRRRRRRR   VVV   VVV  IIIII  DDDDDDDD    OOOOOOO   RRRRRRRR
'     SSS   SSS  EEE        RRR   RRR  VVV   VVV   III   DDD   DDD  OOO   OOO  RRR   RRR
'      SSS       EEE        RRR   RRR  VVV   VVV   III   DDD   DDD  OOO   OOO  RRR   RRR
'       SSSS     EEEEEE     RRRRRRRR   VVV   VVV   III   DDD   DDD  OOO   OOO  RRRRRRRR
'          SSS   EEE        RRR RRR     VVV VVV    III   DDD   DDD  OOO   OOO  RRR RRR
'     SSS   SSS  EEE        RRR  RRR     VVVVV     III   DDD   DDD  OOO   OOO  RRR  RRR
'      SSSSSSS   EEEEEEEEE  RRR   RRR     VVV     IIIII  DDDDDDDD    OOOOOOO   RRR   RRR
'
'
'	REVISADO P/ IE10


	On Error GoTo 0
	Err.Clear

	Const MAX_TAM_CAMPO_OBS = 250
	Const MAX_TAM_CAMPO_PAGTO = 250

	dim s, s2, strScript, usuario, msg_erro, s_transportadora, lista_pedidos, v_pedido, i, achou
	dim c_num_coleta, c_dt_entrega, c_transportadora_contato, c_conferente, c_motorista, c_placa_veiculo, c_nfe_emitente
	dim lngNsuWmsRomaneioN1, lngNsuWmsRomaneioN2, intSequenciaN2
	dim intQtdeNF
	
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	OBTÉM DADOS DO FORMULÁRIO
	s_transportadora = Trim(request("c_transportadora"))
	lista_pedidos = ucase(Trim(request("pedidos_selecionados")))
	if (lista_pedidos = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	c_num_coleta = Trim(Request.Form("c_num_coleta"))
	c_dt_entrega = Trim(Request.Form("c_dt_entrega"))
	c_transportadora_contato = Trim(Request.Form("c_transportadora_contato"))
	c_conferente = Trim(Request.Form("c_conferente"))
	c_motorista = Trim(Request.Form("c_motorista"))
	c_placa_veiculo = Trim(Request.Form("c_placa_veiculo"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	
	lista_pedidos=substitui_caracteres(lista_pedidos,chr(10),"")
	v_pedido = split(lista_pedidos,chr(13),-1)
	achou=False
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if Trim(v_pedido(i))<>"" then
			achou = True
			s = normaliza_num_pedido(v_pedido(i))
			if s <> "" then v_pedido(i) = s
			end if
		next

	if Not achou then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim alerta
	alerta=""

	dim vDadosPedido()
	redim vDadosPedido(0)
	set vDadosPedido(Ubound(vDadosPedido)) = new cl_TRES_COLUNAS
	
	dim n
	n = Request.Form("c_pedido").Count
	for i = 1 to n
		s = Trim(Request.Form("c_pedido")(i))
		if s <> "" then
			if Trim("" & vDadosPedido(UBound(vDadosPedido)).c1) <> "" then
				redim preserve vDadosPedido(UBound(vDadosPedido)+1)
				set vDadosPedido(UBound(vDadosPedido)) = new cl_TRES_COLUNAS
				end if
			with vDadosPedido(UBound(vDadosPedido))
				.c1 = s
				.c2 = Trim(Request.Form("c_obs")(i))
				.c3 = Trim(Request.Form("c_pagto")(i))
				end with
			end if
		next

	if alerta = "" then
		for i=LBound(vDadosPedido) to UBound(vDadosPedido)
			with vDadosPedido(i)
				if Trim("" & .c1) <> "" then
					if Len("" & .c2) > MAX_TAM_CAMPO_OBS then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & Trim("" & .c1) & ": texto digitado no campo 'OBS' possui " & CStr(Len("" & .c2)) & " caracteres e excedeu o tamanho máximo de " & MAX_TAM_CAMPO_OBS
						end if
					if Len("" & .c3) > MAX_TAM_CAMPO_PAGTO then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & Trim("" & .c1) & ": texto digitado no campo '(sem título)' possui " & Cstr(Len("" & .c3)) & " caracteres e excedeu o tamanho máximo de " & MAX_TAM_CAMPO_PAGTO
						end if
					end if
				end with
			next
		end if
		
	intQtdeNF = 0
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i) <> "" then intQtdeNF = intQtdeNF + 1
		next


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, tN1, tN2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim s_log, s_log_aux
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if alerta = "" then
			if Not cria_recordset_pessimista(tN1, msg_erro) then
				alerta = "FALHA AO TENTAR CRIAR CRIAR OBJETO PARA GRAVAR OS DADOS DO ROMANEIO"
				end if
			end if
		
		if alerta = "" then
			if Not fin_gera_nsu(T_WMS_ROMANEIO_N1, lngNsuWmsRomaneioN1, msg_erro) then
				alerta = "FALHA AO GERAR NSU PARA GRAVAR OS DADOS DO ROMANEIO (" & msg_erro & ")"
				end if
			end if
		
		if alerta = "" then
			s = "SELECT * FROM t_WMS_ROMANEIO_N1 WHERE (id = -1)"
			tN1.Open s, cn
			tN1.AddNew
			tN1("id") = lngNsuWmsRomaneioN1
			tN1("usuario") = usuario
			tN1("transportadora_id") = s_transportadora
			tN1("transportadora_num_coleta") = c_num_coleta
			tN1("transportadora_contato") = c_transportadora_contato
			tN1("transportadora_conferente") = c_conferente
			tN1("transportadora_motorista") = c_motorista
			tN1("transportadora_placa_veiculo") = c_placa_veiculo
			tN1("a_entregar_data_marcada") = StrToDate(c_dt_entrega)
			tN1.Update
			tN1.Close
			set tN1 = nothing
			end if
		
		if alerta = "" then
			if Not cria_recordset_pessimista(tN2, msg_erro) then
				alerta = "FALHA AO TENTAR CRIAR CRIAR OBJETO PARA GRAVAR OS DADOS COMPLEMENTARES DO ROMANEIO"
				end if
			end if
		
		if alerta = "" then
			intSequenciaN2 = 0
			
			for i=Lbound(v_pedido) to Ubound(v_pedido)
				if v_pedido(i) <> "" then
					if Not fin_gera_nsu(T_WMS_ROMANEIO_N2, lngNsuWmsRomaneioN2, msg_erro) then
						alerta = "FALHA AO GERAR NSU PARA GRAVAR OS DADOS COMPLEMENTARES DO ROMANEIO (" & msg_erro & ")"
						end if
					
					if alerta = "" then
						intSequenciaN2 = intSequenciaN2 + 1
						s = "SELECT * FROM t_WMS_ROMANEIO_N2 WHERE (id = -1)"
						if tN2.State <> 0 then tN2.Close
						tN2.Open s, cn
						tN2.AddNew
						tN2("id") = lngNsuWmsRomaneioN2
						tN2("id_wms_romaneio_n1") = lngNsuWmsRomaneioN1
						tN2("sequencia") = intSequenciaN2
						tN2("pedido") = v_pedido(i)
						tN2.Update
						end if
					
				'	SE HOUVE ERRO, CANCELA O LAÇO
					if alerta <> "" then exit for
					end if
				next
			
			if tN2.State <> 0 then tN2.Close
			set tN2 = nothing
			end if
		
		if alerta = "" then
			s_log_aux = ""
			if c_num_coleta <> "" then s_log_aux = s_log_aux & " Nº Coleta: " & c_num_coleta & ";"
			if c_dt_entrega <> "" then s_log_aux = s_log_aux & " Data de coleta: " & c_dt_entrega & ";"
			if c_transportadora_contato <> "" then s_log_aux = s_log_aux & " Contato: " & c_transportadora_contato & ";"
			if c_conferente <> "" then s_log_aux = s_log_aux & " Conferente: " & c_conferente & ";"
			if c_motorista <> "" then s_log_aux = s_log_aux & " Motorista: " & c_motorista & ";"
			if c_placa_veiculo <> "" then s_log_aux = s_log_aux & " Placa do veículo: " & c_placa_veiculo & ";"
			
			s_log = "Sucesso na exibição da página p/ geração da planilha do romaneio de entrega (NSU=" & normaliza_a_esq(Cstr(lngNsuWmsRomaneioN1), 3) & "; CD=" & obtem_apelido_empresa_NFe_emitente(c_nfe_emitente) & ") p/ a transportadora " & s_transportadora & ";" & s_log_aux & " Pedido(s) = " & substitui_caracteres(lista_pedidos,chr(13),", ") & "; Obs: gravação dos dados do romaneio no(s) pedido(s) somente após o usuário confirmar que a planilha foi gerada com sucesso."
			grava_log usuario, "", "", "", OP_LOG_ROMANEIO_ENTREGA, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err<>0 then 
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			end if
		end if
%>




<%
'	  C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE
%>


<%=DOCTYPE_LEGADO%>


<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__CONSTXL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fConfirma( f ) {
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}
</script>

<% 
'	GERA O SCRIPT QUE IRÁ ATIVAR O EXCEL E GERAR A PLANILHA NO LADO DO CLIENTE
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'	PREPARA CONSULTA AO BANCO DE DADOS
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	dim idxDadosPedido
	dim r, r_aux
	dim n_reg, n_qtde, n_qtde_volumes, n_qtde_total, n_qtde_volumes_total, vl_pedido
	dim s_sql, s_from, s_where, s_where_pedido
	dim s_pedido, s_pedido_a, s_obs, s_pagto, s_texto_contato, s_cidade, s_uf, s_endereco
	
'	CRIA O RECORDSET AUXILIAR
	if Not cria_recordset_otimista(r_aux, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
'	MONTA CLÁUSULA WHERE
	s_where = ""
	s_where_pedido = ""
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i) <> "" then
			if s_where_pedido <> "" then s_where_pedido = s_where_pedido & " OR "
			s_where_pedido = s_where_pedido & " (item.pedido = '" & v_pedido(i) & "')"
			end if
		next
	
'	CLÁUSULA WHERE
	if s_where_pedido <> "" then s_where = " WHERE " & s_where_pedido
			
'   MONTA CLÁUSULA FROM
	s_from = " FROM t_PEDIDO_ITEM AS item" & _
				" LEFT JOIN t_PRODUTO AS prod ON (item.fabricante=prod.fabricante) AND (item.produto=prod.produto)" & _
				" INNER JOIN t_PEDIDO AS ped ON (item.pedido=ped.pedido)" & _
				" INNER JOIN t_CLIENTE AS cli ON (ped.id_cliente=cli.id)"
	
'   Tipo de NFe: 0-Entrada  1-Saída
	s_sql = "SELECT" & _
				" item.pedido, ped.obs_2, ped.obs_1, item.fabricante, item.produto, prod.descricao, prod.descricao_html," & _
				" item.qtde, item.qtde_volumes," & _
				" ped.st_end_entrega," & _
				" ped.EndEtg_endereco," & _
				" ped.EndEtg_endereco_numero," & _
				" ped.EndEtg_endereco_complemento," & _
				" ped.EndEtg_bairro," & _
				" ped.EndEtg_cidade," & _
				" ped.EndEtg_uf," & _
				" ped.EndEtg_cep,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" ped.endereco_logradouro AS endereco," & _
				" ped.endereco_numero AS endereco_numero," & _
				" ped.endereco_complemento AS endereco_complemento," & _
				" ped.endereco_bairro AS bairro," & _
				" ped.endereco_cidade AS cidade," & _
				" ped.endereco_uf AS uf," & _
				" ped.endereco_cep AS cep,"
	else
		s_sql = s_sql & _
				" cli.endereco," & _
				" cli.endereco_numero," & _
				" cli.endereco_complemento," & _
				" cli.bairro," & _
				" cli.cidade," & _
				" cli.uf," & _
				" cli.cep,"
		end if

	s_sql = s_sql & _
				" (" & _
					"CASE" & _
						" WHEN ped.num_obs_3 > 0 THEN ped.num_obs_3" & _
						" WHEN ped.num_obs_2 > 0 THEN ped.num_obs_2" & _
						" ELSE NULL" & _
				" END) AS numeroNFe" & _
			s_from & _
			s_where & _
			" ORDER BY" & _
				" item.pedido, item.sequencia"
	
	s_texto_contato = ""
	if c_transportadora_contato <> "" then s_texto_contato = "   (contato: " & c_transportadora_contato & ")"
	
'	MONTA SCRIPT P/ EXECUTAR NO BROWSER
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'	DECLARAÇÕES
	strScript = _
	"<script language='JavaScript' type='text/javascript'>" & chr(13) & _
	"function GeraPlanilha() {" & chr(13) & _
	"	var xlFN_LISTAGEM = 'Arial';" & chr(13) & _
	"	var xlFS_LISTAGEM = 8;" & chr(13) & _
	"	var xlFS_CABECALHO = 10;" & chr(13) & _
	"	var xlMargemEsq=1;" & chr(13) & _
	"	var xlOffSetArray=2;" & chr(13) & _
	"	var xlNumLinha, xlNumLinhaAux, xlDadosMinIndex, xlDadosMaxIndex;" & chr(13) & _
	"	var xlNsu, xlEntrega, xlData, xlTransportadora, xlColeta, xlPedido, xlNF, xlQtde, xlProduto, xlValor, xlCidade, xlEndereco, xlObs, xlPagto;" & chr(13) & _
	"	var oXL, oWB, oWS, oRange, oBorders, oFont, oStyle;" & chr(13) & _
	"	var i, s;" & chr(13)

'	TENTA CRIAR UMA INSTÂNCIA DO EXCEL
	strScript = strScript & _
	"	try {" & chr(13) & _
	"		oXL = new ActiveXObject('Excel.Application');" & chr(13) & _ 
	"		}" & chr(13) & _
	"	catch (e) {" & chr(13) & _
	"		alert('Falha ao acionar o Excel!!\n   1) Verifique se o Excel está corretamente instalado\n   2) Verifique se está ativada a opção: Inicializar e executar scripts de controles ActiveX não marcados como seguros');" & chr(13) & _
	"		history.back();" & chr(13) & _
	"		return;" & chr(13) & _
	"		}" & chr(13) 

'	CONFIGURA O EXCEL
	strScript = strScript & _
	"	oXL.Visible=true;" & chr(13) & _
	"	oXL.DisplayAlerts=false;" & chr(13) & _
	"	oXL.SheetsInNewWorkbook=1;" & chr(13) & _
	"	oWB=oXL.Workbooks.Add;" & chr(13) & _
	"	oWB.Windows(1).WindowState=xlMaximized;" & chr(13) & _
	"	oWS=oWB.Worksheets(1);" & chr(13) & _
	"	oWS.PageSetup.PaperSize = xlPaperA4;" & chr(13) & _
	"	oWS.PageSetup.Orientation = xlLandscape;" & chr(13) & _
	"	oWS.PageSetup.LeftMargin = 2;" & chr(13) & _
	"	oWS.PageSetup.RightMargin = 2;" & chr(13) & _
	"	oWS.PageSetup.TopMargin = 15;" & chr(13) & _
	"	oWS.PageSetup.BottomMargin = 15;" & chr(13) & _
	"	oWS.PageSetup.HeaderMargin = 5;" & chr(13) & _
	"	oWS.PageSetup.FooterMargin = 5;" & chr(13) & _
	"	oWS.PageSetup.CenterHorizontally = true;" & chr(13) & _
	"	oXL.Windows(1).DisplayGridlines=false;" & chr(13) & _
	"	oXL.Windows(1).DisplayHeadings=true;" & chr(13) & _
	"	oStyles=oWB.Styles('Normal');" & chr(13) & _
	"	oStyles.IncludeNumber=true;" & chr(13) & _
	"	oStyles.IncludeFont=true;" & chr(13) & _
	"	oStyles.IncludeAlignment=true;" & chr(13) & _
	"	oStyles.IncludeBorder=true;" & chr(13) & _
	"	oStyles.IncludePatterns=true;" & chr(13) & _
	"	oStyles.IncludeProtection=true;" & chr(13) & _
	"	oFont=oStyles.Font" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
	"	oFont.Bold=false;" & chr(13) & _
	"	oFont.Italic=false;" & chr(13) & _
	"	oFont.Underline=xlUnderlineStyleNone;" & chr(13) & _
	"	oFont.Strikethrough=false;" & chr(13) & _
	"	oFont.ColorIndex=xlAutomatic;" & chr(13) & _
	"	oStyles.HorizontalAlignment=xlLeft;" & chr(13) & _
	"	oStyles.VerticalAlignment=xlTop;" & chr(13) & _
	"	oStyles.WrapText=false;" & chr(13) & _
	"	oStyles.Orientation=0;" & chr(13) & _
	"	oStyles.IndentLevel=0;" & chr(13) & _
	"	oStyles.ShrinkToFit=false;" & chr(13) & _
	"	oStyles.Borders(xlLeft).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlRight).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlTop).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlBottom).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlDiagonalDown).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlDiagonalUp).LineStyle=xlNone;" & chr(13) & _
	"	oWS.Cells.Style='Normal';" & chr(13) & _
	"	oWS.Cells.NumberFormat='@';" & chr(13) & _
	"	oWS.DisplayPageBreaks=false;" & chr(13) & _
	"	oWS.Name='Romaneio-Entrega';" & chr(13) & _
	"	oXL.DisplayAlerts=true;" & chr(13)

'	POSIÇÃO DAS COLUNAS
	strScript = strScript & _
	"	xlNF=xlMargemEsq+1;" & chr(13) & _
	"	xlPedido=xlNF+2;" & chr(13) & _
	"	xlQtde=xlPedido+2;" & chr(13) & _
	"	xlProduto=xlQtde+2;" & chr(13) & _
	"	xlValor=xlProduto+2;" & chr(13) & _
	"	xlCidade=xlValor+2;" & chr(13) & _
	"	xlEndereco=xlCidade+2;" & chr(13) & _
	"	xlObs=xlEndereco+2;" & chr(13) & _
	"	xlPagto=xlObs+2;" & chr(13) & _
	"	xlNsu=xlMargemEsq+1;" & chr(13) & _
	"	xlEntrega=xlMargemEsq+1;" & chr(13) & _
	"	xlTransportadora=xlMargemEsq+1;" & chr(13) & _
	"	xlData=xlProduto;" & chr(13) & _
	"	xlColeta=xlObs;" & chr(13)

'	ARRAY USADO P/ TRANSFERIR DADOS P/ O EXCEL
	strScript = strScript & _
	"	xlDadosMinIndex=(xlMargemEsq+1);" & chr(13) & _
	"	xlDadosMaxIndex=xlPagto;" & chr(13)

'	CONFIGURA COLUNAS
	strScript = strScript & _
	"	oWS.Columns(xlMargemEsq).ColumnWidth=1;" & chr(13) & _
	"	oWS.Columns(xlNF).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlNF).ColumnWidth=8;" & chr(13) & _
	"	oWS.Columns(xlNF+1).ColumnWidth=0;" & chr(13) & _
	"	oWS.Columns(xlPedido).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlPedido).ColumnWidth=11;" & chr(13) & _
	"	oWS.Columns(xlPedido+1).ColumnWidth=0;" & chr(13) & _
	"	oWS.Columns(xlQtde).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlQtde).HorizontalAlignment=xlRight;" & chr(13) & _
	"	oWS.Columns(xlQtde).Font.Bold=true;" & chr(13) & _
	"	oWS.Columns(xlQtde).ColumnWidth=4;" & chr(13) & _
	"	oWS.Columns(xlQtde+1).ColumnWidth=0;" & chr(13) & _
	"	oWS.Columns(xlProduto).WrapText=true;" & chr(13) & _
	"	oWS.Columns(xlProduto).ColumnWidth=30;" & chr(13) & _
	"	oWS.Columns(xlProduto+1).ColumnWidth=0;" & chr(13) & _
	"	oWS.Columns(xlValor).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlValor).HorizontalAlignment=xlRight;" & chr(13) & _
	"	oWS.Columns(xlValor).ColumnWidth=10;" & chr(13) & _
	"	oWS.Columns(xlValor+1).ColumnWidth=0;" & chr(13) & _
	"	oWS.Columns(xlCidade).WrapText=true;" & chr(13) & _
	"	oWS.Columns(xlCidade).ColumnWidth=20;" & chr(13) & _
	"	oWS.Columns(xlCidade+1).ColumnWidth=0;" & chr(13) & _
	"	oWS.Columns(xlEndereco).WrapText=true;" & chr(13) & _
	"	oWS.Columns(xlEndereco).ColumnWidth=30;" & chr(13) & _
	"	oWS.Columns(xlEndereco+1).ColumnWidth=0;" & chr(13) & _
	"	oWS.Columns(xlObs).WrapText=true;" & chr(13) & _
	"	oWS.Columns(xlObs).ColumnWidth=43;" & chr(13) & _
	"	oWS.Columns(xlObs+1).ColumnWidth=0;" & chr(13) & _
	"	oWS.Columns(xlPagto).WrapText=true;" & chr(13) & _
	"	oWS.Columns(xlPagto).ColumnWidth=43;" & chr(13)

'	CAMPOS DO CABEÇALHO
	strScript = strScript & _
	"	xlNumLinha=1;" & chr(13) & _
	"	oWS.Rows(xlNumLinha).RowHeight=1;" & chr(13) & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlNsu);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='NSU: " & normaliza_a_esq(Cstr(lngNsuWmsRomaneioN1), 3) & "';" & chr(13) & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlEntrega);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='ENTREGA';" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlData);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='DATA: " & c_dt_entrega & "  (" & formata_hora_hhmm(Now) & ")';" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlColeta);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Gerado em " & formata_data_hora_sem_seg(Now) & "';" & chr(13) & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlTransportadora);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='TRANSPORTADORA: " & s_transportadora & s_texto_contato & "';" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlColeta);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='COLETA: " & c_num_coleta & "';" & chr(13)

'	BORDAS P/ OS TÍTULOS DAS COLUNAS
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	oWS.Rows(xlNumLinha).RowHeight=4;" & chr(13) & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDadosMinIndex) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlDadosMaxIndex) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeTop);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlMedium;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeBottom);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlMedium;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13)

'	TÍTULOS DA COLUNAS
	strScript = strScript & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oFont.Italic=false;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlNF) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='N.F.';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlPedido) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='PEDIDO';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtde) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='QTD';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlProduto) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='PRODUTO';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlValor) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='VALOR';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlCidade) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='CIDADE';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlEndereco) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='ENDERECO';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlObs) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='OBS';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlPagto) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='';" & chr(13)
	
'	BORDAS VERTICAIS
	strScript = strScript & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlNF) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPedido) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtde) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlProduto) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlValor) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlCidade) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlEndereco) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlObs) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPagto) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPagto) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeRight);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13)

	Response.Write strScript
	strScript = ""


'	LAÇO P/ LEITURA DOS DADOS DO BD
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	s_pedido_a = "XXX"
	n_reg = 0
	n_qtde_total = 0
	n_qtde_volumes_total = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

	'	CONTAGEM
		n_reg = n_reg + 1

	'	Nº PEDIDO
		s_pedido = Trim("" & r("pedido"))
		
		idxDadosPedido = LBound(vDadosPedido)-1
		for i=LBound(vDadosPedido) to UBound(vDadosPedido)
			if s_pedido = Trim("" & vDadosPedido(i).c1) then
				idxDadosPedido = i
				exit for
				end if
			next

	'	LINHA DE DADOS
		strScript = strScript & _
			"	xlNumLinha++;" & chr(13) & _
			"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDadosMinIndex) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlDadosMaxIndex) + xlNumLinha.toString();" & chr(13) & _
			"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13)
		
	'	LINHA DE SEPARAÇÃO GROSSA SE MUDOU DE PEDIDO
		if (s_pedido <> s_pedido_a) And (n_reg > 1) then
			strScript = strScript & _
				"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDadosMinIndex) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlDadosMaxIndex) + xlNumLinha.toString();" & chr(13) & _
				"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
				"	oBorders=oRange.Borders(xlEdgeTop);" & chr(13) & _
				"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
				"	oBorders.Weight=xlMedium;" & chr(13) & _
				"	oBorders.ColorIndex=xlAutomatic;" & chr(13)
			end if
		
	'	NF
		s = ""
		if s_pedido <> s_pedido_a then s = Trim("" & r("numeroNFe"))
		strScript = strScript & _
			"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlNF) + xlNumLinha.toString());" & chr(13) & _
			"	oRange.Value='" & s & "';" & chr(13)
		
	'	Nº PEDIDO
		s = ""
		if s_pedido <> s_pedido_a then s = s_pedido
		strScript = strScript & _
			"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlPedido) + xlNumLinha.toString());" & chr(13) & _
			"	oRange.Value='" & s & "';" & chr(13)
			
	'	QTDE
		n_qtde = CLng(r("qtde"))
		n_qtde_volumes = n_qtde * CLng(r("qtde_volumes"))
		s = "SELECT" & _
				" ISNULL(SUM(qtde), 0) AS qtde," & _
				" ISNULL(SUM(qtde * qtde_volumes), 0) AS qtde_volumes_calculado" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE" & _
				" (pedido='" & Trim("" & r("pedido")) & "') AND" & _
				" (fabricante='" & Trim("" & r("fabricante")) & "') AND" & _
				" (produto='" & Trim("" & r("produto")) & "')" & _
			" GROUP BY pedido, fabricante, produto"
		if r_aux.State <> 0 then r_aux.Close
		r_aux.Open s, cn
		if Not r_aux.Eof then
			n_qtde = n_qtde - CLng(r_aux("qtde"))
			n_qtde_volumes = n_qtde_volumes - CLng(r_aux("qtde_volumes_calculado"))
			end if
			
		n_qtde_total = n_qtde_total + n_qtde
		n_qtde_volumes_total = n_qtde_volumes_total + n_qtde_volumes
		
		strScript = strScript & _
			"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtde) + xlNumLinha.toString());" & chr(13) & _
			"	oRange.Value='" & formata_inteiro(n_qtde) & "';" & chr(13)
		
	'	PRODUTO
		strScript = strScript & _
			"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlProduto) + xlNumLinha.toString());" & chr(13) & _
			"	oRange.Value='" & Trim("" & r("descricao")) & "';" & chr(13)
		
	'	VALOR
		s = ""
		if s_pedido <> s_pedido_a then
			vl_pedido = 0
			s = "SELECT ISNULL(SUM(qtde*preco_NF), 0) AS vl_pedido FROM t_PEDIDO_ITEM WHERE" & _
				" (pedido='" & Trim("" & r("pedido")) & "')" & _
				" GROUP BY pedido"
			if r_aux.State <> 0 then r_aux.Close
			r_aux.Open s, cn
			if Not r_aux.Eof then
				vl_pedido = Ccur(r_aux("vl_pedido"))
				end if

			s = "SELECT ISNULL(SUM(qtde*preco_NF), 0) AS vl_pedido FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE" & _
				" (pedido='" & Trim("" & r("pedido")) & "')" & _
				" GROUP BY pedido"
			if r_aux.State <> 0 then r_aux.Close
			r_aux.Open s, cn
			if Not r_aux.Eof then
				vl_pedido = vl_pedido - Ccur(r_aux("vl_pedido"))
				end if
			
			s = formata_moeda(vl_pedido)
			end if
		
		strScript = strScript & _
			"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlValor) + xlNumLinha.toString());" & chr(13) & _
			"	oRange.Value='" & s & "';" & chr(13)
		
	'	CIDADE
		s = ""
		s_endereco = ""
		if s_pedido <> s_pedido_a then
			if Trim("" & r("st_end_entrega")) <> "0" then
				s_cidade = Trim("" & r("EndEtg_cidade"))
				s_uf = Ucase(Trim("" & r("EndEtg_uf")))
				s_endereco = formata_endereco(Trim("" & r("EndEtg_endereco")), Trim("" & r("EndEtg_endereco_numero")), Trim("" & r("EndEtg_endereco_complemento")), Trim("" & r("EndEtg_bairro")), Trim("" & r("EndEtg_cidade")), Trim("" & r("EndEtg_uf")), Trim("" & r("EndEtg_cep")))
			else
				s_cidade = Trim("" & r("cidade"))
				s_uf = Ucase(Trim("" & r("uf")))
				s_endereco = formata_endereco(Trim("" & r("endereco")), Trim("" & r("endereco_numero")), Trim("" & r("endereco_complemento")), Trim("" & r("bairro")), Trim("" & r("cidade")), Trim("" & r("uf")), Trim("" & r("cep")))
				end if
			
			if s_cidade <> "" then s_cidade = iniciais_em_maiusculas(s_cidade)
			
			if (s_cidade <> "") And (s_uf <> "") then
				s = s_cidade & " / " & s_uf
			else
				s = s_cidade & s_uf
				end if
			end if
		
		strScript = strScript & _
			"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlCidade) + xlNumLinha.toString());" & chr(13) & _
			"	oRange.Value='" & JsQuotedStr(s) & "';" & chr(13) & _
			"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlEndereco) + xlNumLinha.toString());" & chr(13) & _
			"	oRange.Value='" & JsQuotedStr(s_endereco) & "';" & chr(13)
		
	'	OBS
		s_obs = ""
		s_pagto = ""
		if s_pedido <> s_pedido_a then
			if idxDadosPedido >= LBound(vDadosPedido) then
				s_obs = "" & vDadosPedido(idxDadosPedido).c2
				s_pagto = "" & vDadosPedido(idxDadosPedido).c3
				end if
			s_obs = Replace(s_obs, vbCrLf, "\n")
			s_pagto = Replace(s_pagto, vbCrLf, "\n")
		'	OBSERVA O TAMANHO MÁXIMO QUE CABE EM UMA CÉLULA
			if Len(s_obs) > 255 then
				s = Left(s_obs, 255)
				s_obs = mid(s_obs, 256)
			else
				s = s_obs
				s_obs = ""
				end if
			
			if Len(s_pagto) > 255 then
				s2 = Left(s_pagto, 255)
				s_pagto = mid(s_pagto, 256)
			else
				s2 = s_pagto
				s_pagto = ""
				end if
			
			strScript = strScript & _
				"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlObs) + xlNumLinha.toString());" & chr(13) & _
				"	oRange.Value='" & substitui_caracteres(s, "'", "´") & "';" & chr(13) & _
				"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlPagto) + xlNumLinha.toString());" & chr(13) & _
				"	oRange.Value='" & s2 & "';" & chr(13)
			end if
		
	'	SE A OBSERVAÇÃO NÃO COUBE, CONTINUA NA PRÓXIMA LINHA
		do while (Len(s_obs) > 0) Or (Len(s_pagto) > 0)
			s = ""
			s2 = ""
			
			if Len(s_obs) > 0 then
			'	OBSERVA O TAMANHO MÁXIMO QUE CABE EM UMA CÉLULA
				if Len(s_obs) > 255 then
					s = Left(s_obs, 255)
					s_obs = mid(s_obs, 256)
				else
					s = s_obs
					s_obs = ""
					end if
				end if
			
			if Len(s_pagto) > 0 then
			'	OBSERVA O TAMANHO MÁXIMO QUE CABE EM UMA CÉLULA
				if Len(s_pagto) > 255 then
					s2 = Left(s_pagto, 255)
					s_pagto = mid(s_pagto, 256)
				else
					s2 = s_pagto
					s_pagto = ""
					end if
				end if
				
			strScript = strScript & _
				"	xlNumLinha++;" & chr(13) & _
				"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlObs) + xlNumLinha.toString();" & chr(13) & _
				"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
				"	oRange.Value='" & s & "';" & chr(13) & _
				"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPagto) + xlNumLinha.toString();" & chr(13) & _
				"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
				"	oRange.Value='" & s2 & "';" & chr(13)
			loop
			
	'	LINHA DE SEPARAÇÃO
		strScript = strScript & _
			"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDadosMinIndex) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlDadosMaxIndex) + xlNumLinha.toString();" & chr(13) & _
			"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
			"	oBorders=oRange.Borders(xlEdgeBottom);" & chr(13) & _
			"	oBorders.LineStyle=xlDot;" & chr(13) & _
			"	oBorders.Weight=xlHairline;" & chr(13) & _
			"	oBorders.ColorIndex=xlAutomatic;" & chr(13)
		
	'	BORDAS VERTICAIS
		strScript = strScript & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlNF) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPedido) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtde) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlProduto) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlValor) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlCidade) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlEndereco) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlObs) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPagto) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPagto) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeRight);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlThin;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13)
		
		s_pedido_a = s_pedido
		
		Response.Write strScript
		strScript = ""
		
		r.MoveNext
		loop
	
	
'	LINHA SEPARADORA
	strScript = strScript & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDadosMinIndex) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlDadosMaxIndex) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeBottom);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlMedium;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13)
	
'	TOTAL DE VOLUMES
	strScript = strScript & _
		"	xlNumLinha++;" & chr(13) & _
		"	oWS.Rows(xlNumLinha).RowHeight=6;" & chr(13) & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPedido) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlNF) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.Merge();" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oRange.HorizontalAlignment=xlRight;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	oRange.Value='TOTAL VOLUMES:';" & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtde) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	oRange.Value='" & formata_inteiro(n_qtde_volumes_total) & "';" & chr(13)
		
'	QTDE NF
	strScript = strScript & _
		"	xlNumLinha++;" & chr(13) & _
		"	oWS.Rows(xlNumLinha).RowHeight=6;" & chr(13) & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPedido) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlNF) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.Merge();" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oRange.HorizontalAlignment=xlRight;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	oRange.Value='QTDE NF:';" & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtde) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	oRange.Value='" & formata_inteiro(intQtdeNF) & "';" & chr(13)

'	CONFERENTE
	strScript = strScript & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlNF);" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oRange.Value='CONFERENTE: " & Ucase(c_conferente) & "';" & chr(13)

'	MOTORISTA
	strScript = strScript & _
		"	xlNumLinha++;" & chr(13) & _
		"	oWS.Rows(xlNumLinha).RowHeight=4;" & chr(13) & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlNF);" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oRange.Value='MOTORISTA: " & Ucase(c_motorista) & "';" & chr(13)

'	PLACA DO VEÍCULO
	strScript = strScript & _
		"	xlNumLinha++;" & chr(13) & _
		"	oWS.Rows(xlNumLinha).RowHeight=4;" & chr(13) & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlNF);" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oRange.Value='PLACA DO VEÍCULO: " & Ucase(c_placa_veiculo) & "';" & chr(13)

'	LINHA P/ ASSINATURA
	strScript = strScript & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlNF);" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oRange.Value='Ass: _____________________________________________';" & chr(13)
	
'	AJUSTES FINAIS NA PLANILHA
	strScript = strScript & _
		"	oXL.Windows(oWB.Name).Activate;" & chr(13) & _
		"	oXL.Sheets(1).Select;" & chr(13) & _
		"	oXL.Sheets(1).Range('A1').Select;" & chr(13)

'	FECHAMENTO DO SCRIPT
	strScript = strScript & _
		"}" & chr(13) & _
		"</script>" & chr(13)

	Response.Write strScript

%>




<!-- C A S C A D I N G   S T Y L E   S H E E T

	 CCCCCCC    SSSSSSS    SSSSSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	CCC        SSS        SSS
	CCC         SSSS       SSSS
	CCC            SSSS       SSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	 CCCCCCC    SSSSSSS    SSSSSSS
-->

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% else %>

<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';bCONFIRMA.focus();GeraPlanilha();" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="f" name="f" method="post" action="RomaneioConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedidos_selecionados" id="pedidos_selecionados" value="<%=lista_pedidos%>">
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=s_transportadora%>">
<input type="hidden" name="c_num_coleta" id="c_num_coleta" value="<%=c_num_coleta%>">
<input type="hidden" name="c_dt_entrega" id="c_dt_entrega" value="<%=c_dt_entrega%>">
<input type="hidden" name="c_transportadora_contato" id="c_transportadora_contato" value="<%=c_transportadora_contato%>">
<input type="hidden" name="c_conferente" id="c_conferente" value="<%=c_conferente%>" />
<input type="hidden" name="c_motorista" id="c_motorista" value="<%=c_motorista%>" />
<input type="hidden" name="c_placa_veiculo" id="c_placa_veiculo" value="<%=c_placa_veiculo%>" />
<input type="hidden" name="c_nsu_romaneio" id="c_nsu_romaneio" value="<%=Cstr(lngNsuWmsRomaneioN1)%>" />
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Romaneio de Entrega<span class="C">&nbsp;</span></span></td>
</tr>
</table>
<br>
<br>


<!-- ************   MENSAGEM  ************ -->
<span class="Lbl">ATENÇÃO</span>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><span style='margin:5px 2px 5px 2px;'>Confirme se a planilha com NSU=<%=normaliza_a_esq(Cstr(lngNsuWmsRomaneioN1), 3)%> foi gerada corretamente</span></div>
<br><br>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fConfirma(f)" title="confirma o romaneio de entrega">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>