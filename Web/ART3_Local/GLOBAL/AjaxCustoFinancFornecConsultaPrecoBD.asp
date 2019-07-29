<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =========================================
'	  AjaxCustoFinancFornecConsultaPrecoBD.asp
'     =========================================
'
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear

	class cl_RespConsultaParcelamentoBD
		dim fabricante
		dim produto
		dim status
		dim precoLista
		dim descricao
		dim descricao_html
		dim codigo_erro
		dim msg_erro
		end class
	
'	OBTEM O ID
	dim strSql, strResp, msg_erro
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim strLoja, strTipoParcelamento, strQtdeParcelas, strListaProdutos, vProdutos, vAux, intCounter
	dim coeficiente
	strLoja = Trim(Request("loja"))
	strTipoParcelamento = Trim(Request("tipoParcelamento"))
	strQtdeParcelas = Trim(Request("qtdeParcelas"))
	strListaProdutos = Trim(Request("ListaProdutos"))
		
	if (strTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA) And _
	   (strTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) And _
	   (strTipoParcelamento <> COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) then
		Response.End
	elseif (converte_numero(strQtdeParcelas)=0) And _
		   ((strTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA)Or(strTipoParcelamento=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA)) then
		Response.End
	elseif strListaProdutos = "" then
		Response.End
	elseif converte_numero(strLoja) = 0 then
		Response.End
		end if

	dim vResp
	redim vResp(0)
	set vResp(0) = new cl_RespConsultaParcelamentoBD
	vResp(0).produto = ""
	
	vProdutos = Split(strListaProdutos, ";")
	for intCounter=Lbound(vProdutos) to UBound(vProdutos)
		if Trim(vProdutos(intCounter)) <> "" then
			if vResp(Ubound(vResp)).produto <> "" then
				redim preserve vResp(Ubound(vResp)+1)
				set vResp(Ubound(vResp)) = new cl_RespConsultaParcelamentoBD
				end if
			vAux = Split(vProdutos(intCounter), "|")
			vResp(Ubound(vResp)).fabricante = vAux(0)
			vResp(Ubound(vResp)).produto = vAux(1)
			end if
		next
	
	for intCounter=Lbound(vResp) to Ubound(vResp)
		if (Trim(vResp(intCounter).produto) <> "") And _
		   (Trim(vResp(intCounter).status) = "") then
			   
			if strTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA then
				coeficiente = 1
			else
				strSql = _
					"SELECT " & _
						"*" & _
					" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
					" WHERE" & _
						" (fabricante = '" & vResp(intCounter).fabricante & "')" & _
						" AND (tipo_parcelamento = '" & strTipoParcelamento & "')" & _
						" AND (qtde_parcelas = " & strQtdeParcelas & ")"
				if rs.State <> 0 then rs.Close
				rs.open strSql, cn
				if rs.Eof then
					vResp(intCounter).status = "ERR"
					vResp(intCounter).codigo_erro = "1"
					vResp(intCounter).msg_erro = "Opção de parcelamento não disponível para fornecedor " & vResp(intCounter).fabricante & ": " & decodificaCustoFinancFornecQtdeParcelas(strTipoParcelamento, strQtdeParcelas) & " parcela(s)"
				else
					coeficiente = converte_numero(rs("coeficiente"))
					end if
				end if
			
			strSql = _
				"SELECT " & _
					"*" & _
				" FROM t_PRODUTO" & _
					" INNER JOIN t_PRODUTO_LOJA" & _
						" ON (t_PRODUTO.fabricante=t_PRODUTO_LOJA.fabricante) AND (t_PRODUTO.produto=t_PRODUTO_LOJA.produto)" & _
				" WHERE" & _
					" (t_PRODUTO.fabricante = '" & vResp(intCounter).fabricante & "')" & _
					" AND (t_PRODUTO.produto = '" & vResp(intCounter).produto & "')" & _
					" AND (CONVERT(smallint,loja) = " & strLoja & ")"
			if rs.State <> 0 then rs.Close
			rs.open strSql, cn
			if rs.Eof then
				vResp(intCounter).status = "ERR"
				vResp(intCounter).codigo_erro = "2"
				vResp(intCounter).msg_erro = "Produto " & vResp(intCounter).produto & " não localizado para a loja " & strLoja & "."
			else
				vResp(intCounter).descricao = Trim("" & rs("descricao"))
				vResp(intCounter).descricao_html = produto_formata_descricao_em_html(Trim("" & rs("descricao_html")))
				if vResp(intCounter).status = "" then
					vResp(intCounter).status = "OK"
					vResp(intCounter).precoLista = formata_moeda(coeficiente * rs("preco_lista"))
					end if
				end if
			end if
		next
	
'	MONTA A RESPOSTA
	strResp = ""
	
	for intCounter=Lbound(vResp) to Ubound(vResp)
		if Trim(vResp(intCounter).produto) <> "" then
			strResp = strResp & _
					  "<ItemConsulta>" & _
						"<fabricante>" & _
							vResp(intCounter).fabricante & _
						"</fabricante>" & _
						"<produto>" & _
							vResp(intCounter).produto & _
						"</produto>" & _
						"<status>" & _
							vResp(intCounter).status & _
						"</status>" & _
						"<precoLista>" & _
							vResp(intCounter).precoLista & _
						"</precoLista>" & _
						"<descricao>" & _
							vResp(intCounter).descricao & _
						"</descricao>" & _
						"<descricao_html>" & _
							"<![CDATA[" & _
							vResp(intCounter).descricao_html & _
							"]]>" & _
						"</descricao_html>" & _
						"<codigo_erro>" & _
							vResp(intCounter).codigo_erro & _
						"</codigo_erro>" & _
						"<msg_erro>" & _
							vResp(intCounter).msg_erro & _
						"</msg_erro>" & _
					  "</ItemConsulta>"
			end if
		next


'	HÁ RESPOSTA?
	if strResp <> "" then 
		Response.ContentType="text/xml"
		strResp = "<?xml version='1.0' encoding='ISO-8859-1'?>" & _
				  "<TabelaResultado>" & _
				  strResp & _
				  "</TabelaResultado>"
		end if
		
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
