<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  C O M I S S A O D E S C C O N S I S T E . A S P
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

	dim s, usuario, rb_comissao_descontada, lista_pedidos, v_pedido, v_aux, i, j, achou, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	rb_comissao_descontada = Trim(Request.Form("rb_comissao_descontada"))
	
	lista_pedidos = ucase(Trim(request("c_pedidos")))
	if (lista_pedidos = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	lista_pedidos=substitui_caracteres(lista_pedidos,chr(10),"")
	v_aux = split(lista_pedidos,chr(13),-1)
	achou=False
	for i=Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			achou = True
			s = normaliza_num_pedido(v_aux(i))
			if s <> "" then v_aux(i) = s
			end if
		next

	if Not achou then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	redim v_pedido(0)
	v_pedido(Ubound(v_pedido))=""
	for i = Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			if Trim(v_pedido(Ubound(v_pedido)))<>"" then
				redim preserve v_pedido(Ubound(v_pedido)+1)
				end if
			v_pedido(Ubound(v_pedido)) = Trim(v_aux(i))
			end if
		next
	
	lista_pedidos = join(v_pedido,chr(13))
	
	dim alerta
	alerta = ""

	dim observacoes
	observacoes = ""
	
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i)<>"" then
			for j=Lbound(v_pedido) to (i-1)
				if v_pedido(i) = v_pedido(j) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & ": linha " & renumera_com_base1(Lbound(v_pedido),i) & " repete o mesmo pedido da linha " & renumera_com_base1(Lbound(v_pedido),j) & "."
					exit for
					end if
				next
			end if
		next

	if alerta = "" then
		if rb_comissao_descontada = "" then
			alerta = "Informe se o valor do item devolvido e/ou perda deve ser assinalado como 'descontado' ou 'não-descontado'."
		elseif (rb_comissao_descontada <> "S") And (rb_comissao_descontada <> "N") then
			alerta = "Opção desconhecida (" & rb_comissao_descontada & ")"
			end if
		end if

	if alerta = "" then
		for i = Lbound(v_pedido) to Ubound(v_pedido)
			if v_pedido(i) <> "" then
			'	PEDIDO EXISTE?
				s = "SELECT pedido FROM t_PEDIDO WHERE (pedido='" & Trim(v_pedido(i)) & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & Trim(v_pedido(i)) & " não está cadastrado."
				else
				'	PEDIDO TEM ITEM DEVOLVIDO OU VALOR DE PERDA?  EM QUE SITUAÇÃO?
					s = "SELECT pedido, comissao_descontada FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (pedido='" & Trim(v_pedido(i)) & "')" & _
						" UNION " & _
						"SELECT pedido, comissao_descontada FROM t_PEDIDO_PERDA WHERE (pedido='" & Trim(v_pedido(i)) & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & Trim(v_pedido(i)) & " não possui item devolvido ou valor de perda."
					else
						if rb_comissao_descontada = "S" then
							s = "SELECT pedido, comissao_descontada FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (pedido='" & Trim(v_pedido(i)) & "') AND (comissao_descontada=" & COD_COMISSAO_NAO_DESCONTADA & ")" & _
								" UNION " & _
								"SELECT pedido, comissao_descontada FROM t_PEDIDO_PERDA WHERE (pedido='" & Trim(v_pedido(i)) & "') AND (comissao_descontada=" & COD_COMISSAO_NAO_DESCONTADA & ")"
							if rs.State <> 0 then rs.Close
							rs.open s, cn
							if rs.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Pedido " & Trim(v_pedido(i)) & " não possui item devolvido ou valor de perda ainda não descontados das comissões."
								end if
						elseif rb_comissao_descontada = "N" then
							s = "SELECT pedido, comissao_descontada FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (pedido='" & Trim(v_pedido(i)) & "') AND (comissao_descontada=" & COD_COMISSAO_DESCONTADA & ")" & _
								" UNION " & _
								"SELECT pedido, comissao_descontada FROM t_PEDIDO_PERDA WHERE (pedido='" & Trim(v_pedido(i)) & "') AND (comissao_descontada=" & COD_COMISSAO_DESCONTADA & ")"
							if rs.State <> 0 then rs.Close
							rs.open s, cn
							if rs.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Pedido " & Trim(v_pedido(i)) & " não possui item devolvido ou valor de perda já descontados das comissões."
								end if
							end if
						end if
					end if
				end if
			next
		end if






' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const c_OP_DEVOLUCAO = "DEVOLUCAO"
const c_OP_PERDA = "PERDA"
dim r
dim s, s_sql, cab_table, cab
dim s_where, s_where_pedido
dim x
dim i, n_indice
dim w_ckb_alterar, w_pedido, w_data, w_valor, w_operacao

'	MONTA CLÁUSULA WHERE
'	SE O OBJETIVO É ANOTAR COMO "COMISSÃO DESCONTADA", ENTÃO EXIBE APENAS 
'	OS QUE ESTÃO COMO "COMISSÃO NÃO-DESCONTADA" E VICE-VERSA.
	if rb_comissao_descontada = "S" then
		s_where = " (comissao_descontada=" & COD_COMISSAO_NAO_DESCONTADA & ")"
	else
		s_where = " (comissao_descontada=" & COD_COMISSAO_DESCONTADA & ")"
		end if
	
'	RESTRINGE AOS PEDIDOS ESPECIFICADOS
	s_where_pedido = ""
	for i = Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i) <> "" then
			if s_where_pedido <> "" then s_where_pedido = s_where_pedido & " OR"
			s_where_pedido = s_where_pedido & " (pedido='" & Trim(v_pedido(i)) & "')"
			end if
		next

	if s_where_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_pedido & ")"
		end if

	if s_where <> "" then s_where = " WHERE" & s_where
	
'	MONTA CLÁUSULA FROM
	s_sql = "SELECT pedido, devolucao_data AS data, (qtde*preco_venda) AS valor," & _
			" id, '" & c_OP_DEVOLUCAO & "' AS operacao" & _
			" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
			s_where & _
			" UNION " & _
			"SELECT pedido, data, valor," & _
			" id, '" & c_OP_PERDA & "' AS operacao" & _
			" FROM t_PEDIDO_PERDA" & _
			s_where & _
			" ORDER BY pedido, operacao, data, valor"
	
  ' CABEÇALHO
	w_ckb_alterar = 50
	w_pedido = 70
	w_data = 65
	w_valor = 75
	w_operacao = 70
	
	cab_table = "<table style='border:2px solid black;' cellspacing=0 cellpadding=0>" & chr(13)
	cab = "	<tr style='background:mintcream;'>" & chr(13) & _
		  "		<td class='MDTE' style='width:" & cstr(w_ckb_alterar) & "px' align='center' valign='bottom' nowrap><span class='PLTc'>Alterar</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_pedido) & "px' align='left' valign='bottom'><span class='PLTe'>Pedido</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_data) & "px' align='center' valign='bottom'><span class='PLTc'>Data</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_valor) & "px' align='right' valign='bottom'><span class='PLTd'>Valor</span></td>" & chr(13) & _
		  "		<td class='MTD' style='width:" & cstr(w_operacao) & "px' align='left' valign='bottom'><span class='PLTe'>Operação</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = cab_table & cab
	n_indice = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		n_indice = n_indice + 1
			
		x = x & "	<tr>" & chr(13)

	'> ALTERAR?
		s = Trim("" & r("pedido"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td valign='top' align='center' class='MDTE'>" & chr(13) & _
				"			<input type='checkbox' tabindex='-1' id='ckb_alterar' name='ckb_alterar'" & _
				" value='" & Trim("" & r("pedido")) & ";" & Trim("" & r("id")) & ";" & Trim("" & r("operacao")) & "'>" & chr(13) & _
				"		</td>" & chr(13)

	'> PEDIDO
		s = Trim("" & r("pedido"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD' align='left'>" & chr(13) & _
				"			<span class='Cn'>" & s & "</span>" & chr(13) & _
				"		</td>" & chr(13)

	'> DATA
		s = formata_data(r("data"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD' align='center'>" & chr(13) & _
				"			<span class='Cnc'>" & s & "</span>" & chr(13) & _
				"		</td>" & chr(13)

	'> VALOR
		s = formata_moeda(r("valor"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD' align='right'>" & chr(13) & _
				"			<span class='Cnd'>" & s & "</span>" & chr(13) & _
				"		</td>" & chr(13)

	'> OPERAÇÃO
		s = Trim("" & r("operacao"))
		if s = c_OP_DEVOLUCAO then
			s = "Devolução"
		elseif s = c_OP_PERDA then
			s = "Perda"
		else
			s = "&nbsp;"
			end if
		
		x = x & "		<td class='MTD' align='left'>" & chr(13) & _
				"			<span class='Cn'>" & s & "</span>" & chr(13) & _
				"		</td>" & chr(13)

		x = x & "	</tr>" & chr(13)
			
		r.MoveNext
		loop


  ' FECHA TABELA
	x = x & "</table>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub

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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function marca_todos( ) {
var f,i;
	f=fCOMISSAODESC;
	for (i=0; i<f.ckb_alterar.length; i++) {
		if (trim(f.ckb_alterar.value)!="") {
			if (!f.ckb_alterar[i].checked) {
				f.ckb_alterar[i].checked=true;
				}
			}
		}
}

function desmarca_todos( ) {
var f,i;
	f=fCOMISSAODESC;
	for (i=0; i<f.ckb_alterar.length; i++) {
		if (trim(f.ckb_alterar.value)!="") {
			if (f.ckb_alterar[i].checked) {
				f.ckb_alterar[i].checked=false;
				}
			}
		}
}

function fCOMISSAODESCConfirma( f ) {
var i,blnFlagOk;
	blnFlagOk=false;
	for (i=0; i<f.ckb_alterar.length; i++) {
		if (f.ckb_alterar[i].checked) {
			blnFlagOk=true;
			break;
			}
		}

	if (!blnFlagOk) {
		alert("Nenhum item foi selecionado!!");
		return;
		}
	
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}
</script>




<!-- C A S C A D I N G   S T Y L E   S H E E T

	 CCCCCCC    SSSSSSS    SSSSSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	CCC        SSS        SSS
	CCC         SSSS       SSSS
	CCC            SSSS       SSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	 CCCCCCC    SSSSSSS    SSSSSSS
-->

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">


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
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ***************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS PARA CONFIRMAÇÃO  ********** -->
<!-- ***************************************************************** -->
<body onload="focus();">
<center>

<form id="fCOMISSAODESC" name="fCOMISSAODESC" method="post" action="ComissaoDescConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedidos_selecionados" id="pedidos_selecionados" value="<%=lista_pedidos%>">
<input type="hidden" name="rb_comissao_descontada" id="rb_comissao_descontada" value="<%=rb_comissao_descontada%>">
<!-- FORÇA A CRIAÇÃO DE UM ARRAY MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" class="CBOX" name="ckb_alterar" id="ckb_alterar" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Comissão (Descontos)</span></td>
</tr>
</table>
<br>


<!-- ************   HÁ OBSERVAÇÕES?  ************ -->
<% if observacoes <> "" then %>
		<span class="Lbl">OBSERVAÇÕES</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'><%=observacoes%></p></div>
		<br><br>
<% end if %>

<!--  PEDIDOS  -->
<table class="Qx" cellspacing="0">
	<tr bgcolor="#FFFFFF">
		<td class="MT" nowrap align='center' style='background:azure;'><span class="PLTc">Alterar para&nbsp;</span></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap>
		<% if rb_comissao_descontada = "S" then %>
			<span class="PLLe" style='margin-left:4pt;margin-right:4pt;'>Comissão Descontada</span>
		<% elseif rb_comissao_descontada = "N" then %>
			<span class="PLLe" style='margin-left:4pt;margin-right:4pt;'>Comissão Não-Descontada</span>
		<% else %>
			<span class="PLLe" style='color:red;margin-left:4pt;margin-right:4pt;'>Erro: código desconhecido</span>
		<% end if %>
		</td>
	</tr>
</table>

<br>

<!-- MONTA TABELA COM OS ITENS DEVOLVIDOS E VALORES DE PERDA P/ ASSINALAR QUAIS DEVEM SER PROCESSADOS -->
<% consulta_executa %>

<br>
<input name="bMarcaTodos" id="bMarcaTodos" type="button" class="Button" onclick="marca_todos();" value="marca todos" title="seleciona todos os itens"
>&nbsp;&nbsp;&nbsp;<input name="bDesmarcaTodos" id="bDesmarcaTodos" type="button" class="Button" onclick="desmarca_todos();" value="desmarca todos" title="desmarca todos os itens">


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fCOMISSAODESCConfirma(fCOMISSAODESC)" title="confirma a operação">
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
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>