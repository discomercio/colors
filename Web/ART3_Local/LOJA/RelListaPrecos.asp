<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L L I S T A P R E C O S . A S P
'     ======================================================
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim alerta
	alerta = ""

	dim s, c_fabricante
	c_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	if c_fabricante <> "" then
		s = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
		if s <> "" then c_fabricante = s
		end if

	if alerta = "" then
		if c_fabricante <> "" then
			s = "SELECT fabricante, nome, razao_social FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta = "Fabricante " & c_fabricante & " não está cadastrado."
				end if
			end if
		end if
	




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA MONTA
'
function consulta_monta
dim s
	s = "SELECT t_PRODUTO_LOJA.fabricante, t_PRODUTO_LOJA.produto, descricao, descricao_html, preco_lista, cor, vendavel" & _
		" FROM t_PRODUTO_LOJA INNER JOIN t_PRODUTO ON (t_PRODUTO_LOJA.fabricante=t_PRODUTO.fabricante) AND (t_PRODUTO_LOJA.produto=t_PRODUTO.produto)" & _
		" WHERE" & _
		" ((vendavel='S') OR (vendavel='X'))" & _
		" AND (loja='" & loja & "')"

	if c_fabricante <> "" then
		s = s & " AND (t_PRODUTO_LOJA.fabricante='" & c_fabricante & "')"
		end if
	
	s = s & " ORDER BY t_PRODUTO_LOJA.fabricante, Upper(descricao)"
	
	consulta_monta = s
	
end function


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const W_INDISPONIVEL = 5
const W_CODIGO = 70
const W_DESCRICAO = 280
const W_PRECO = 100
dim r
dim s, s_aux, s_sql, s_cor, fabricante_a, x, cab_table, cab, msg_erro
dim n_reg, n_reg_total
dim inclui_largura, blnIndisponivel

	s_sql = consulta_monta()

  ' CABEÇALHO
	cab_table = "<TABLE class='MD' cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "<TR style='background: #FFF0E0' nowrap>" & _
		  "<TD width='" & Cstr(W_INDISPONIVEL) & "' valign='bottom' nowrap class='MD' style='background: #FFFFFF'>&nbsp;</TD>" & _
		  "<TD width='" & Cstr(W_CODIGO) & "' valign='bottom' nowrap class='MD MB'><P class='R'>PRODUTO</P></TD>" & _
		  "<TD width='" & Cstr(W_DESCRICAO) & "' valign='bottom' nowrap class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & _
		  "<TD width='" & Cstr(W_PRECO) & "' valign='bottom' nowrap class='MB'><P class='Rd' style='font-weight:bold;'>PREÇO</P></TD>" & _
		  "</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	fabricante_a = "XXXXXX"
	inclui_largura = false
	
	if Not cria_recordset_otimista(r, msg_erro) then
		Response.Write msg_erro
		exit sub
		end if
	
	r.Open s_sql, cn
	do while Not r.Eof
	'	MUDOU FABRICANTE?
		if (Trim("" & r("fabricante"))<>fabricante_a) then
			if n_reg_total > 0 then
			  ' FECHA TABELA DO FABRICANTE ANTERIOR
				x = x & "</TABLE>" & chr(13)
				Response.Write x
				x="<BR>"
				end if

		  ' INICIA NOVA TABELA P/ O NOVO FABRICANTE
			if n_reg_total > 0 then x = x & "<BR>"
			s = Trim("" & r("fabricante"))
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table & _
				"<TR><TD class='MD'>&nbsp;</TD><TD COLSPAN='3' align='center' class='MTB' style='background:azure;'>" & _
				"<P class='F'>" & s & "</P></TD></TR>"
			
			x = x & chr(13) & cab 
			n_reg = 0
			fabricante_a = Trim("" & r("fabricante"))
			end if
		
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "<TR NOWRAP>" & chr(13)

		s_cor = Trim("" & r("cor"))
		if s_cor = "" then s_cor = "black"
		
	'	PRODUTO INDISPONÍVEL?
		if r("vendavel") = "X" then blnIndisponivel = True else blnIndisponivel = False
		if blnIndisponivel then s_cor = "gray"
		
	 '> PRODUTO INDISPONÍVEL
		x = x & "	<TD valign='bottom'	class='MD'"
		if inclui_largura then x = x & " width='" & Cstr(W_INDISPONIVEL) & "'"
		x = x & "><P class='C' style='color:red'>"
		if blnIndisponivel then x = x & "X" else x = x & "&nbsp;"
		x = x & "</P></TD>" & chr(13)
	 
	 '> CÓDIGO DO PRODUTO
		x = x & "	<TD valign='bottom'	class='MDB'"
		if inclui_largura then x = x & " width='" & Cstr(W_CODIGO) & "'"
		x = x & "><P class='C' style='color:" & s_cor & "'>&nbsp;" & Trim("" & r("produto")) & "</P></TD>" & chr(13)

	 '> DESCRIÇÃO
		x = x & "	<TD valign='bottom'	class='MDB'"
		if inclui_largura then x = x & " width='" & Cstr(W_DESCRICAO) & "'"
		x = x & "><P class='C' style='color:" & s_cor & "'>&nbsp;" & produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13)

 	 '> PREÇO
		x = x & "	<TD valign='bottom' class='MB'"
		if inclui_largura then x = x & " width='" & Cstr(W_PRECO) & "'"
		x = x & " NOWRAP><P class='Cd' style='color:" & s_cor & "'>&nbsp;" & formata_moeda(r("preco_lista")) & "</P></TD>" & chr(13)

		x = x & "</TR>" & chr(13)
		
		inclui_largura = False
		
		if (n_reg mod 200) = 0 then
			Response.Write x & "</table>" & chr(13)
			x = "" & "<TABLE class='MD' cellSpacing=0 cellPadding=0>" & chr(13)
			inclui_largura = True
			end if
				
		r.movenext
		loop
	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table
		if c_fabricante <> "" then
			s = c_fabricante
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "<TR><TD class='MD'>&nbsp;</TD><TD COLSPAN='3' align='center' class='MTB' style='background:azure;'>" & _
				"<P class='F'>" & s & "</P></TD></TR>" & chr(13)
			end if
			
		x = x & cab & _
			"<TR NOWRAP>" & _
			"	<TD class='MD'>&nbsp;</TD><TD colspan='3' class='MB'><P class='ALERTA'>&nbsp;NÃO HÁ PRODUTOS&nbsp;</P></TD>" & _
			"</TR>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO FABRICANTE
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing

end sub


' _____________________________________
' LISTA MONTA
'
sub lista_monta
dim r, x, s_sql, fabricante_a, nome_fabricante, n_reg, msg_erro

	x = "<script language='JavaScript'>" & chr(13) & _
		"var data_emissao = '" & formata_data_hora(Now) & "';" & chr(13) & _
		"var Pd = new Array();" & chr(13) & _
		"Pd[0] = new oPd('','','','','','');" & chr(13)

	s_sql = consulta_monta()
	
	fabricante_a = "XXXXXX"
	nome_fabricante = ""
	
	n_reg = 0

	if Not cria_recordset_otimista(r, msg_erro) then
		Response.Write msg_erro
		exit sub
		end if
	
	r.Open s_sql, cn
	do while Not r.Eof
		n_reg = n_reg + 1
		
	'	MUDOU FABRICANTE?
		if (Trim("" & r("fabricante"))<>fabricante_a) then
			nome_fabricante = x_fabricante(Trim("" & r("fabricante")))
			fabricante_a = Trim("" & r("fabricante"))
			end if
		
	 '> MONTA LINHA
		x = x & "Pd[Pd.length]=new oPd('" & Trim("" & r("fabricante")) & "'" & _
				",'" & nome_fabricante & "'" & _
				",'" & Trim("" & r("produto")) & "'" & _
				",'" & filtra_nome_identificador(Trim("" & r("descricao"))) & "'" & _
				",'" & formata_moeda(r("preco_lista")) & "'" & _
				",'" & UCase(Trim("" & r("vendavel"))) & "'" & _
				");" & chr(13)
		
		nome_fabricante = ""

		if (n_reg mod 200) = 0 then
			Response.Write x
			x = ""
			end if

		r.movenext
		loop

	x = x & "</script>" & chr(13)
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




<html>


<head>
<!-- #include file = "../global/printerx.txt"    -->
	<title>LOJA</title>
	</head>


<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

var ja_carregou=false;

function impressora_carregada() {
var s;
	if (!( "object" == typeof(printer) && "string" == typeof(printer.module_id))) {
		s = "Componente necessário para impressão do pedido não foi carregado corretamente!!";
		alert(s);
		return false;
		}
	return true;
}

function fLPRECOSImpressora ( f ) {
	if (!ja_carregou) return;
	if (!impressora_carregada()) return;
	printer.Initialize;
	if (printer.printing) printer.EndDoc();
	printer.seleciona_impressora();
}

function fLPRECOSMargens( f ) {
	if (!ja_carregou) return;
	if (!impressora_carregada()) return;
	printer.configura_margens();
}

function oPd( fabricante, nome_fabricante, produto, descricao, preco, vendavel ) {
	this.fabricante = fabricante;
	this.nome_fabricante = nome_fabricante;
	this.produto = produto;
	this.descricao = descricao;
	this.preco = preco;
	this.vendavel = vendavel;
}

function fLPRECOSImprime( f ) {
var s, cx, cy, h, margemx, margemy, altura, iv, fabricante_a;
var ix_vendavel, wx_vendavel, ix_produto, wx_produto, ix_descricao, wx_descricao, ix_preco, wx_preco;
var imprime_cabecalho, titulo, titulo_base, pagina, b,tam_listagem;
	tam_listagem=10;
	b=false;
	for (iv=1; iv < Pd.length; iv++) {
		if (Pd[iv].produto!='') {
			b=true;
			break;
			}
		}
	if (!b) {
		alert("Não há produtos!!");
		return;
		}
	if (!ja_carregou) return;
	if (!impressora_carregada(printer)) return;
	printer.Initialize;
	if (printer.printing) printer.EndDoc();
	printer.landscape=false;
	printer.setpapersizeletter();
	printer.job_title='LISTA DE PREÇOS';
	printer.brushstyle='bsClear';
	printer.fontcolor=0;
	printer.fontname='Arial';
	printer.fontsize=tam_listagem;
	printer.fontnormal=true;
	printer.penmode='pmBlack';
	printer.penstyle='psSolid';
	printer.pencolor=0;
	printer.penwidth=1;
	printer.BeginDoc();
	margemx=25;
	margemy=2;
	altura=printer.pageheight/printer.pixelspermmY - 30;
	
	ix_vendavel = margemx;
	wx_vendavel = 4;
	ix_produto = ix_vendavel + wx_vendavel;
	wx_produto = 22;
	ix_descricao = ix_produto + wx_produto + 2;
	wx_descricao = 110;
	ix_preco = ix_descricao + wx_descricao + 2;
	wx_preco = 25;
	
	fabricante_a = 'XXXXXX';
	cy=margemy;
	pagina=0;
	for (iv=1; iv < Pd.length; iv++) {
		if (fabricante_a!=Pd[iv].fabricante) {
			if (iv > 1) {
				printer.newpage();
				cy=margemy;
				}
			fabricante_a=Pd[iv].fabricante;
			s=Pd[iv].fabricante;
			if ((s!='')&&(Pd[iv].nome_fabricante!='')) s=s + ' - ';
			s=s + Pd[iv].nome_fabricante;
			titulo_base = s;
			titulo=titulo_base;
			imprime_cabecalho=true;
			}
			
		if (cy > altura) {
			printer.newpage();
			cy=margemy;
			titulo=titulo_base + '  (continuação)';
			imprime_cabecalho=true;
			}
			
		if (imprime_cabecalho) {
			imprime_cabecalho=false;
			cx=ix_produto;
			printer.fontsize=12;
			printer.fontbold=true;
			printer.imprime(cx, cy, titulo);
			cy=cy+printer.texto_altura('X');
			printer.fontsize=tam_listagem;
			printer.fontbold=true;
			pagina=pagina+1;
			s=formata_inteiro(pagina);
			printer.imprime(ix_produto, altura+2.5*printer.texto_altura('X'), data_emissao);
			printer.imprime(ix_preco+wx_preco-printer.texto_largura(s), altura+2.5*printer.texto_altura('X'), s);
			h=printer.texto_altura('X');
			printer.linha(ix_produto-1, cy, ix_preco+wx_preco+1, cy);
			printer.linha(ix_produto-1,cy, ix_produto-1,cy+h);
			printer.linha(ix_descricao-1, cy, ix_descricao-1, cy+h);
			printer.linha(ix_preco-1, cy, ix_preco-1, cy+h);
			printer.linha(ix_preco+wx_preco+1, cy, ix_preco+wx_preco+1, cy+h);
			printer.linha(ix_produto-1, cy+h, ix_preco+wx_preco+1, cy+h);
			printer.imprime(ix_produto,cy,'CÓDIGO');
			printer.imprime(ix_descricao,cy,'DESCRIÇÃO');
			s='PREÇO';
			cx=ix_preco+wx_preco-printer.texto_largura(s);
			printer.imprime(cx,cy,s);
			cy=cy+h;
			printer.fontbold=false;
			}

		h=printer.texto_altura('X');
		printer.linha(ix_produto-1,cy,ix_produto-1,cy+h);
		printer.linha(ix_descricao-1, cy, ix_descricao-1, cy+h);
		printer.linha(ix_preco-1, cy, ix_preco-1, cy+h);
		printer.linha(ix_preco+wx_preco+1, cy, ix_preco+wx_preco+1, cy+h);
		printer.linha(ix_produto-1, cy+h, ix_preco+wx_preco+1, cy+h);
		
		if (Pd[iv].vendavel == 'X') {
			cx=ix_vendavel;
			printer.imprime_campo(cx,cy,wx_vendavel,'X');
		}
		
		cx=ix_produto;
		printer.imprime_campo(cx,cy,wx_produto,Pd[iv].produto);
		
		cx=ix_descricao;
		printer.imprime_campo(cx,cy,wx_descricao,Pd[iv].descricao);
		
		cx=ix_preco+wx_preco-printer.texto_largura(Pd[iv].preco);
		printer.imprime_campo(cx,cy,wx_preco,Pd[iv].preco);
		
		cy=cy+h;
		}
		
	printer.EndDoc();
	alert('Lista de preços foi impressa!!');
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">


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
<body onload="window.status='Concluído';ja_carregou=true;">

<center>

<form id="fLPRECOS" name="fLPRECOS" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Lista de Preços</span>
	<br>
	<%	s = "<span class='N'>Emissão:&nbsp;" & formata_data_hora(Now) & "</span>"
		Response.Write s
	%>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>


<!--  LISTA DE PREÇOS  -->
<% lista_monta %>

<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="center"><div name="dIMPRESSORA" id="dIMPRESSORA">
		<a name="bIMPRESSORA" id="bIMPRESSORA" href="javascript:fLPRECOSImpressora(fLPRECOS)" title="seleciona a impressora">
		<img src="../botao/impressora.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="center"><div name="dMARGENS" id="dMARGENS">
		<a name="bMARGENS" id="bMARGENS" href="javascript:fLPRECOSMargens(fLPRECOS)" title="configura as margens de impressão">
		<img src="../botao/margens.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dIMPRIME" id="dIMPRIME">
		<a name="bIMPRIME" id="bIMPRIME" href="javascript:fLPRECOSImprime(fLPRECOS)" title="imprime a listagem em formulário contínuo">
		<img src="../botao/imprimir.gif" width="176" height="55" border="0"></a></div>
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
