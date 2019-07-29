<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<%
'     =====================================
'	  C L I E N T E P E S Q U I S A . A S P
'     =====================================
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
'	OBTEM USUÁRIO
	dim usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("Aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	Dim cn
	If Not bdd_conecta(cn) then Response.Redirect("Aviso.asp?id=" & ERR_CONEXAO)

'	PARÂMETROS DE PESQUISA
	dim cnpj_cpf_selecionado, nome_selecionado
	cnpj_cpf_selecionado = retorna_so_digitos(trim(request("cnpj_cpf_selecionado")))
	nome_selecionado = trim(request("nome_selecionado"))
	
	dim ordenacao_selecionada
	ordenacao_selecionada=Trim(request("ordenacao_selecionada"))

	dim s, s_tabela, cliente_unico_encontrado, qtde_clientes
	s_tabela=executa_consulta(cliente_unico_encontrado, qtde_clientes)

	if cliente_unico_encontrado<>"" then
		Response.Redirect("clienteedita.asp?cliente_selecionado=" & cliente_unico_encontrado & "&operacao_selecionada=" & OP_CONSULTA)
		end if
		
		


' ________________________________
' E X E C U T A _ C O N S U L T A
'
function executa_consulta(byref id_registro_unico_encontrado, byref qtde_registros)
dim consulta, s, i, x, cab, s_aux, s_ult_id
dim r

	id_registro_unico_encontrado = ""
	qtde_registros = 0
	
  ' CABEÇALHO
	cab="<TABLE class='Q' cellSpacing=0 width='640'>" & chr(13)
	cab=cab & "<TR style='background: #FFF0E0'>"
	cab=cab & "<TD width='35' class='MD MB'><P class='R' style='cursor:pointer; margin-right:2pt;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "fORDConcluir('1');" & chr(34) & ">CNPJ/CPF</P></TD>"
	cab=cab & "<TD width='200' class='MD MB'><P class='R' style='cursor:pointer; margin-left:2pt' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "fORDConcluir('2');" & chr(34) & ">NOME</P></TD>"
	cab=cab & "<TD width='250' class='MB'><P class='R' style='cursor:pointer; margin-left:2pt' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "fORDConcluir('3');" & chr(34) & ">ENDEREÇO</P></TD>"
	cab=cab & "</TR>" & chr(13)

	s=""
	if cnpj_cpf_selecionado<>"" then
		s=" (cnpj_cpf='" & cnpj_cpf_selecionado & "')"
	elseif nome_selecionado<>"" then
		s=" (nome LIKE '" & BD_CURINGA_TODOS & nome_selecionado & BD_CURINGA_TODOS & "'" & SQL_COLLATE_CASE_ACCENT & ")"
		end if
	
	if s <> "" then s = " WHERE " & s
	
	consulta= "SELECT * FROM t_CLIENTE" & s & " ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "cnpj_cpf"
		case "2": consulta = consulta & "nome, cnpj_cpf"
		case "3": consulta = consulta & "endereco, cnpj_cpf"
		case else: consulta = consulta & "cnpj_cpf"
		end select

  ' EXECUTA CONSULTA
	x=cab
	i=0
	
	set r = cn.Execute( consulta )

	while not r.eof 
	  ' CONTAGEM
		i = i + 1
		qtde_registros = qtde_registros + 1

	  ' ALTERNÂNCIA NAS CORES DAS LINHAS
		if (i AND 1)=0 then
			x=x & "<TR NOWRAP style='background: #FFF0E0'>"
		else
			x=x & "<TR NOWRAP >"
			end if

	 '> CNPJ/CPF
		x=x & " <TD class='MDB' NOWRAP valign='top'><P class='C' style='margin-left:2pt'>"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste cliente'>"
		x=x & cnpj_cpf_formata(r("cnpj_cpf")) & "</a></P></TD>"

	 '> NOME
		x=x & " <TD class='MDB' valign='top'><P class='C' style='margin-left:2pt'>" 
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste cliente'>"
		x=x & Trim("" & r("nome")) & "</a></P></TD>"

	 '> ENDEREÇO
		s=Trim("" & r("endereco"))
		if s<>"" then
			s_aux=Trim("" & r("endereco_numero"))
			if s_aux<>"" then s=s & ", " & s_aux
			s_aux=Trim("" & r("endereco_complemento"))
			if s_aux<>"" then s=s & " " & s_aux
			s_aux=Trim("" & r("bairro"))
			if s_aux<>"" then s=s & " - " & s_aux
			s_aux=Trim("" & r("cidade"))
			if s_aux<>"" then s=s & " - " & s_aux
			s_aux=Trim("" & r("uf"))
			if s_aux<>"" then s=s & " - " & s_aux
			end if

		if s="" then s="&nbsp;"
		x=x & " <TD class='MB' valign='top'><P class='C' style='margin-left:2pt'>"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste cliente'>"
		x=x & s & "</a></P></TD>"

		x=x & "</TR>"

		s_ult_id = Trim("" & r("id"))
		
		r.MoveNext
		wend
	

  ' SE FOI ENCONTRADO APENAS UM ÚNICO REGISTRO, RETORNA SEU ID
	if i = 1 then id_registro_unico_encontrado = s_ult_id

  ' MOSTRA TOTAL DE CLIENTES
	x=x & "<TR NOWRAP style='background: #FFFFDD'><TD COLSPAN='3' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "&nbsp;&nbsp;clientes" & "</P></TD></TR>"

  ' FECHA TABELA
	x=x & "</TABLE>"
	

	executa_consulta = x

	r.close
	set r=nothing

End function

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

<html>


<head>
	<title><%=TITULO_JANELA_MODULO_ORCAMENTO%></title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando consulta ...';

function fORDConcluir(ordenacao){
	window.status = "Aguarde ...";
	fORD.ordenacao_selecionada.value=ordenacao;
	fORD.submit(); 
}

function fCADConcluir( f ){
var s_cnpj_cpf;
	s_cnpj_cpf=trim(f.c_novo.value);
	if ((s_cnpj_cpf=="")||(!cnpj_cpf_ok(s_cnpj_cpf))) {
		alert("CNPJ/CPF inválido!!");
		f.c_novo.focus();
		return false;
		}
	f.cnpj_cpf_selecionado.value=s_cnpj_cpf;
	window.status = "Aguarde ...";
	f.submit(); 
}

function fOPConcluir(s_id){
	window.status = "Aguarde ...";
	fOP.cliente_selecionado.value=s_id;
	fOP.submit(); 
}

function fNEWConcluir( f ){
	window.status = "Aguarde ...";
	f.submit(); 
}

</script>

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">



<% if (qtde_clientes = 0) And (cnpj_cpf_selecionado="") then s="fCAD.c_novo.focus();" else s="" %>
<body onload="window.status='Concluído';<%=s%>" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Resultado da Pesquisa de Clientes</span>
	<br><span class="Rc">
		<a href="resumo.asp" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DE CLIENTES  -->
<br>
<center>
<form method="post" action="ClientePesquisa.asp" id="fORD" name="fORD">
<input type="hidden" name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" value='<%=cnpj_cpf_selecionado%>'>
<input type="hidden" name="nome_selecionado" id="nome_selecionado" value='<%=nome_selecionado%>'>
<input type="hidden" name="ordenacao_selecionada" id="ordenacao_selecionada" value=''>
</form>

<form method="post" action="ClienteEdita.asp" id="fOP" name="fOP">
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=cliente_unico_encontrado%>'>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<% =s_tabela %>
</form>

<br>
<br>

<% if (qtde_clientes = 0) And (cnpj_cpf_selecionado<>"") And cnpj_cpf_ok(cnpj_cpf_selecionado) then %>
	<!--  PESQUISOU POR CNPJ/CPF E NÃO ENCONTROU CLIENTE: APRESENTA LINK PARA CADASTRAR O CNPJ/CPF JÁ INFORMADO  -->
	<form action="ClienteEdita.asp" method="post" id="fNEW" name="fNEW" onsubmit="if (!fNEWConcluir(fNEW)) return false">
	<input type="hidden" name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" value='<%=cnpj_cpf_selecionado%>'>
	<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_INCLUI%>'>
		
		<a href="javascript:fNEWConcluir(fNEW);">
		<div class='MtLink' style="width:400px;font-weight:bold;" align="center">
		<p style='margin:5px 2px 5px 2px;'>Cliente&nbsp;<%=cnpj_cpf_formata(cnpj_cpf_selecionado)%>&nbsp;ainda não está cadastrado.<br>Clique aqui para cadastrá-lo agora.</p>
		</div></a>
		
		<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

	</div>
	</form>
	
<% elseif (qtde_clientes = 0) And (cnpj_cpf_selecionado="") then %>
	<!--  PESQUISOU POR NOME E NÃO ENCONTROU CLIENTE: APRESENTA CAMPO PARA CADASTRAR NOVO CLIENTE  -->
	<form action="ClienteEdita.asp" method="post" id="fCAD" name="fCAD" onsubmit="if (!fCADConcluir(fCAD)) return false">
	<input type="hidden" name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" value=''>
	<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_INCLUI%>'>

	<span class="Lbl">CADASTRAR NOVO CLIENTE</span>
	<div class="QFn" style="width:300px;background:floralwhite;" align="center">
	<table class="TFn">
		<tr>
			<td nowrap>
				<span class="Lbl" style="cursor:default" onclick="fCAD.c_novo.focus();">CNPJ/CPF</span>&nbsp;
					<input name="c_novo" id="c_novo" type="text" maxlength="18" size="20" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true) && tem_info(this.value) && cnpj_cpf_ok(this.value)) if (fCADConcluir(fCAD)) submit(); filtra_cnpj_cpf();"><br>
				</td>
			</tr>
		</table>
		<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
		<input name="EXECUTAR" id="EXECUTAR" type="submit" class="Botao" value="EXECUTAR" title="executa">
		<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
	</div>
	</form>
<% end if %>

<p class="TracoBottom"></p>

<table class="notPrint" cellspacing="0">
<tr>
	<td align="center"><a href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>


</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
