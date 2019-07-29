<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelTabelaCustoFinanceiro.asp
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
' CONSULTA EXECUTA
'
sub consulta_executa
const W_PARCELAS = 60
const W_COEFICIENTE = 80
dim rsFabr, rsCE, rsSE
dim s, s_aux, s_sql, s_where, fabricante_a, x, cab_table, cab, msg_erro
dim n_reg_sub_total, n_reg_total
dim intIndiceNumParcela
dim strQtdeParcelas, strCoeficiente

  ' CABEÇALHO
	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)

	cab = _
		  "	<TR style='height:4px;'>" & chr(13) & _
		  "		<TD colspan='5'></TD>" & chr(13) & _
		  "	</TR>" & chr(13) & _
		  "	<TR style='background:#FFF0E0' NOWRAP>" & chr(13) & _
		  "		<TD colspan='2' align='center' NOWRAP class='ME MD MC MB'><P class='Rc' style='color:black;'>Com Entrada</p></TD>" & chr(13) & _
		  "		<TD style='background:#FFFFFF'>&nbsp;</TD>" & chr(13) & _
		  "		<TD colspan='2' align='center' NOWRAP class='ME MD MC MB'><P class='Rc' style='color:black;'>Sem Entrada</p></TD>" & chr(13) & _
		  "	</TR>" & chr(13) & _
		  "	<TR style='background:#FFF0E0' NOWRAP>" & chr(13) & _
		  "		<TD width='" & Cstr(W_PARCELAS) & "' align='center' valign='bottom' NOWRAP class='ME MD MB'><P class='Rc' style='color:black;'>Parcelas</p></TD>" & chr(13) & _
		  "		<TD width='" & Cstr(W_COEFICIENTE) & "' align='right' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='color:black;font-weight:bold;'>Coeficiente</P></TD>" & chr(13) & _
		  "		<TD style='background:#FFFFFF'><span style='width:30px;'>&nbsp;</span></TD>" & chr(13) & _
		  "		<TD width='" & Cstr(W_PARCELAS) & "' align='center' valign='bottom' NOWRAP class='ME MD MB'><P class='Rc' style='color:black;'>Parcelas</p></TD>" & chr(13) & _
		  "		<TD width='" & Cstr(W_COEFICIENTE) & "' align='right' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='color:black;font-weight:bold;'>Coeficiente</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)

	x = ""
	n_reg_sub_total = 0
	n_reg_total = 0
	fabricante_a = "XXXXXX"

	s_where = ""
	if c_fabricante <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (fabricante = '" & c_fabricante & "')"
		end if

	if s_where <> "" then s_where = " WHERE" & s_where
	
	s_sql = "SELECT" & _
				" DISTINCT fabricante" & _
			" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
			s_where & _
			" ORDER BY" & _
				" fabricante"
	Set rsFabr = cn.Execute(s_sql)
	do while Not rsFabr.Eof
	'	MUDOU FABRICANTE?
		if (Trim("" & rsFabr("fabricante"))<>fabricante_a) then
			if n_reg_total > 0 then
			  ' FECHA TABELA DO FABRICANTE ANTERIOR
				x = x & "</TABLE>" & chr(13)
				Response.Write x
				x="<BR>"
				end if

		  ' INICIA NOVA TABELA P/ O NOVO FABRICANTE
			if n_reg_total > 0 then x = x & "<BR>"
			s = Trim("" & rsFabr("fabricante"))
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table & _
					"	<TR>" & chr(13) & _
					"		<TD COLSPAN='5' align='center' class='ME MD MC MB' style='background:azure;'>" & _
						"<P class='F'>" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)

			x = x & chr(13) & cab 
			n_reg_sub_total = 0
			fabricante_a = Trim("" & rsFabr("fabricante"))
			end if
			
	'	COEFICIENTES P/ OPÇÃO 'COM ENTRADA'
		s_sql = "SELECT " & _
					"*" & _
				" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
				" WHERE" & _
					" (fabricante = '" & Trim("" & rsFabr("fabricante")) & "')" & _
					" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA & "')" & _
				" ORDER BY" & _
					" qtde_parcelas"
		Set rsCE = cn.Execute(s_sql)

	'	COEFICIENTES P/ OPÇÃO 'SEM ENTRADA'
		s_sql = "SELECT " & _
					"*" & _
				" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
				" WHERE" & _
					" (fabricante = '" & Trim("" & rsFabr("fabricante")) & "')" & _
					" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA & "')" & _
				" ORDER BY" & _
					" qtde_parcelas"
		Set rsSE = cn.Execute(s_sql)
		
	'	MONTA TABELA P/ EXIBIÇÃO (NAVEGA ATÉ O FINAL DOS 2 RECORDSETS)
		intIndiceNumParcela = 0
		do while (Not rsCE.Eof) Or (Not rsSE.Eof)
		
		 ' CONTAGEM
			n_reg_sub_total = n_reg_sub_total + 1
			n_reg_total = n_reg_total + 1

			intIndiceNumParcela = intIndiceNumParcela + 1
			
			x = x & "	<TR NOWRAP>" & chr(13)

		'	COM ENTRADA
		'	===========
			strQtdeParcelas = "&nbsp;"
			strCoeficiente = "&nbsp;"
			if Not rsCE.Eof then
				if CInt(intIndiceNumParcela) = CInt(rsCE("qtde_parcelas")) then
					strQtdeParcelas = decodificaCustoFinancFornecQtdeParcelas(rsCE("tipo_parcelamento"), rsCE("qtde_parcelas"))
					strCoeficiente = formata_coeficiente_custo_financ_fornecedor(rsCE("coeficiente"))
					end if
				end if
				
		'	QTDE PARCELAS
			x = x & "		<TD valign='bottom' class='ME MD MB'><P class='Cc'>&nbsp;" & strQtdeParcelas & "</P></TD>" & chr(13)
	 
		'	COEFICIENTE
			x = x & "		<TD valign='bottom' class='MD MB'><P class='Cd'>&nbsp;" & strCoeficiente & "</P></TD>" & chr(13)

		'	COLUNA SEPARADORA
		'	=================
			x = x & "		<TD>&nbsp;</TD>" & chr(13)
			
		'	SEM ENTRADA
		'	===========
			strQtdeParcelas = "&nbsp;"
			strCoeficiente = "&nbsp;"
			if Not rsSE.Eof then
				if CInt(intIndiceNumParcela) = CInt(rsSE("qtde_parcelas")) then
					strQtdeParcelas = decodificaCustoFinancFornecQtdeParcelas(rsSE("tipo_parcelamento"), rsSE("qtde_parcelas"))
					strCoeficiente = formata_coeficiente_custo_financ_fornecedor(rsSE("coeficiente"))
					end if
				end if
				
		'	QTDE PARCELAS
			x = x & "		<TD valign='bottom' class='ME MD MB'><P class='Cc'>&nbsp;" & strQtdeParcelas & "</P></TD>" & chr(13)
	 
		'	COEFICIENTE
			x = x & "		<TD valign='bottom' class='MD MB'><P class='Cd'>&nbsp;" & strCoeficiente & "</P></TD>" & chr(13)
			
			x = x & "	</TR>" & chr(13)
			
		'	AVANÇA RECORDSETS?
			if Not rsCE.Eof then
				if CInt(rsCE("qtde_parcelas")) <= CInt(intIndiceNumParcela) then rsCE.MoveNext
				end if
				
			if Not rsSE.Eof then
				if CInt(rsSE("qtde_parcelas")) <= CInt(intIndiceNumParcela) then rsSE.MoveNext
				end if
			loop

		rsFabr.movenext
		loop
		
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table
		if c_fabricante <> "" then
			s = c_fabricante
			s_aux = x_fabricante(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & "	<TR>" & chr(13) & _
					"		<TD COLSPAN='5' align='center' class='ME MTB' style='background:azure;'><P class='F'>" & s & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
			
		x = x & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='5' class='ME MB'><P class='ALERTA'>&nbsp;NÃO HÁ COEFICIENTES&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO FABRICANTE
	x = x & "</TABLE>" & chr(13)
	
	Response.write x
end sub




' _____________________________________
' TABELA IMPRESSAO MONTA
'
sub tabela_impressao_monta
dim rsFabr, rsCE, rsSE
dim x, s_sql, s_where, fabricante_a, nome_fabricante, n_reg, msg_erro
dim intIndiceNumParcela
dim strQtdeParcelasCE, strCoeficienteCE, strQtdeParcelasSE, strCoeficienteSE

	x = "<script language='JavaScript'>" & chr(13) & _
		"var data_emissao = '" & formata_data_hora(Now) & "';" & chr(13) & _
		"var Pd = new Array();" & chr(13) & _
		"Pd[0] = new oPd('','','','','','');" & chr(13)

	fabricante_a = "XXXXXX"
	nome_fabricante = ""
	
	n_reg = 0

	s_where = ""
	if c_fabricante <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (fabricante = '" & c_fabricante & "')"
		end if

	if s_where <> "" then s_where = " WHERE" & s_where
	
	s_sql = "SELECT" & _
				" DISTINCT fabricante" & _
			" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
			s_where & _
			" ORDER BY" & _
				" fabricante"
	Set rsFabr = cn.Execute(s_sql)
	do while Not rsFabr.Eof
		n_reg = n_reg + 1
		
	'	MUDOU FABRICANTE?
		if (Trim("" & rsFabr("fabricante"))<>fabricante_a) then
			nome_fabricante = x_fabricante(Trim("" & rsFabr("fabricante")))
			fabricante_a = Trim("" & rsFabr("fabricante"))
			end if

	'	COEFICIENTES P/ OPÇÃO 'COM ENTRADA'
		s_sql = "SELECT " & _
					"*" & _
				" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
				" WHERE" & _
					" (fabricante = '" & Trim("" & rsFabr("fabricante")) & "')" & _
					" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA & "')" & _
				" ORDER BY" & _
					" qtde_parcelas"
		Set rsCE = cn.Execute(s_sql)

	'	COEFICIENTES P/ OPÇÃO 'SEM ENTRADA'
		s_sql = "SELECT " & _
					"*" & _
				" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
				" WHERE" & _
					" (fabricante = '" & Trim("" & rsFabr("fabricante")) & "')" & _
					" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA & "')" & _
				" ORDER BY" & _
					" qtde_parcelas"
		Set rsSE = cn.Execute(s_sql)
		
	'	MONTA TABELA P/ EXIBIÇÃO (NAVEGA ATÉ O FINAL DOS 2 RECORDSETS)
		intIndiceNumParcela = 0
		do while (Not rsCE.Eof) Or (Not rsSE.Eof)
		
			intIndiceNumParcela = intIndiceNumParcela + 1
		
		'	COM ENTRADA
		'	===========
			strQtdeParcelasCE = ""
			strCoeficienteCE = ""
			if Not rsCE.Eof then
				if CInt(intIndiceNumParcela) = CInt(rsCE("qtde_parcelas")) then
					strQtdeParcelasCE = decodificaCustoFinancFornecQtdeParcelas(rsCE("tipo_parcelamento"), rsCE("qtde_parcelas"))
					strCoeficienteCE = formata_coeficiente_custo_financ_fornecedor(rsCE("coeficiente"))
					end if
				end if

		'	SEM ENTRADA
		'	===========
			strQtdeParcelasSE = ""
			strCoeficienteSE = ""
			if Not rsSE.Eof then
				if CInt(intIndiceNumParcela) = CInt(rsSE("qtde_parcelas")) then
					strQtdeParcelasSE = decodificaCustoFinancFornecQtdeParcelas(rsSE("tipo_parcelamento"), rsSE("qtde_parcelas"))
					strCoeficienteSE = formata_coeficiente_custo_financ_fornecedor(rsSE("coeficiente"))
					end if
				end if
		
		'> MONTA LINHA
			x = x & "Pd[Pd.length]=new oPd('" & Trim("" & rsFabr("fabricante")) & "'" & _
					",'" & nome_fabricante & "'" & _
					",'" & strQtdeParcelasCE & "'" & _
					",'" & strCoeficienteCE & "'" & _
					",'" & strQtdeParcelasSE & "'" & _
					",'" & strCoeficienteSE & "'" & _
					");" & chr(13)
		
			nome_fabricante = ""

			if (n_reg mod 200) = 0 then
				Response.Write x
				x = ""
				end if

		'	AVANÇA RECORDSETS?
			if Not rsCE.Eof then
				if CInt(rsCE("qtde_parcelas")) <= CInt(intIndiceNumParcela) then rsCE.MoveNext
				end if
				
			if Not rsSE.Eof then
				if CInt(rsSE("qtde_parcelas")) <= CInt(intIndiceNumParcela) then rsSE.MoveNext
				end if
			loop

		rsFabr.movenext
		loop

	x = x & "</script>" & chr(13)
	Response.write x
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
		s = "Componente necessário para impressão não foi carregado corretamente!!";
		alert(s);
		return false;
		}
	return true;
}

function fTabCustoFinancFornecImpressora ( f ) {
	if (!ja_carregou) return;
	if (!impressora_carregada()) return;
	printer.Initialize;
	if (printer.printing) printer.EndDoc();
	printer.seleciona_impressora();
}

function fTabCustoFinancFornecMargens( f ) {
	if (!ja_carregou) return;
	if (!impressora_carregada()) return;
	printer.configura_margens();
}

function oPd( fabricante, nome_fabricante, parcelasCE, coeficienteCE, parcelasSE, coeficienteSE ) {
	this.fabricante = fabricante;
	this.nome_fabricante = nome_fabricante;
	this.parcelasCE = parcelasCE;
	this.coeficienteCE = coeficienteCE;
	this.parcelasSE = parcelasSE;
	this.coeficienteSE = coeficienteSE;
}

function fTabCustoFinancFornecImprime( f ) {
var s, cx, cy, h, margemx, margemy, altura, iv, fabricante_a;
var ix_parcelasCE, wx_parcelasCE, ix_coeficienteCE, wx_coeficienteCE;
var ix_colSeparadora, wx_colSeparadora;
var ix_parcelasSE, wx_parcelasSE, ix_coeficienteSE, wx_coeficienteSE;
var imprime_cabecalho, titulo, titulo_base, pagina, tam_listagem;
	tam_listagem=10;
	if (!ja_carregou) return;
	if (!impressora_carregada(printer)) return;
	printer.Initialize;
	if (printer.printing) printer.EndDoc();
	printer.landscape=false;
	printer.setpapersizeletter();
	printer.job_title='TABELA DE CUSTO FINANCEIRO';
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
	
	ix_parcelasCE = margemx+25;
	wx_parcelasCE = 20;
	ix_coeficienteCE = ix_parcelasCE + wx_parcelasCE;
	wx_coeficienteCE = 24;
	ix_colSeparadora = ix_coeficienteCE + wx_coeficienteCE;
	wx_colSeparadora = 20;
	ix_parcelasSE = ix_colSeparadora + wx_colSeparadora;
	wx_parcelasSE = wx_parcelasCE;
	ix_coeficienteSE = ix_parcelasSE + wx_parcelasSE;
	wx_coeficienteSE = wx_coeficienteCE;
	
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
			printer.fontsize=14;
			printer.fontbold=true;
			s='Tabela de Custo Financeiro';
			cx=ix_parcelasCE+(ix_coeficienteSE+wx_coeficienteSE-ix_parcelasCE-printer.texto_largura(s))/2;
			printer.imprime(cx, cy, s);
			cy=cy+printer.texto_altura('X')+5;
			printer.fontsize=12;
			printer.fontbold=true;
			h=printer.texto_altura('X');
		//  NOME DA EMPRESA
			printer.linha(ix_parcelasCE-1, cy, ix_coeficienteSE+wx_coeficienteSE+1, cy);
			printer.linha(ix_parcelasCE-1, cy+h, ix_coeficienteSE+wx_coeficienteSE+1, cy+h);
			printer.linha(ix_parcelasCE-1,cy, ix_parcelasCE-1,cy+h);
			printer.linha(ix_coeficienteSE+wx_coeficienteSE+1, cy, ix_coeficienteSE+wx_coeficienteSE+1, cy+h);
			cx=ix_parcelasCE+(ix_coeficienteSE+wx_coeficienteSE-ix_parcelasCE-printer.texto_largura(titulo))/2;
			printer.imprime(cx, cy, titulo);
			cy=cy+printer.texto_altura('X')+2;
			printer.fontsize=tam_listagem;
			printer.fontbold=true;
			pagina=pagina+1;
			s=formata_inteiro(pagina);
			printer.imprime(ix_parcelasCE, altura+2.5*printer.texto_altura('X'), data_emissao);
			printer.imprime(ix_coeficienteSE+wx_coeficienteSE-printer.texto_largura(s), altura+2.5*printer.texto_altura('X'), s);
			h=printer.texto_altura('X');
		//	1ª LINHA DE TÍTULO
		//	COM ENTRADA
			printer.linha(ix_parcelasCE-1, cy, ix_coeficienteCE+wx_coeficienteCE+1, cy);
			printer.linha(ix_parcelasCE-1, cy+h, ix_coeficienteCE+wx_coeficienteCE+1, cy+h);
			printer.linha(ix_parcelasCE-1,cy, ix_parcelasCE-1,cy+h);
			printer.linha(ix_coeficienteCE+wx_coeficienteCE+1, cy, ix_coeficienteCE+wx_coeficienteCE+1, cy+h);
			s='Com Entrada';
			printer.imprime(ix_parcelasCE+(wx_parcelasCE+wx_coeficienteCE-printer.texto_largura(s))/2,cy,s);
		//  SEM ENTRADA
			printer.linha(ix_parcelasSE-1, cy, ix_coeficienteSE+wx_coeficienteSE+1, cy);
			printer.linha(ix_parcelasSE-1, cy+h, ix_coeficienteSE+wx_coeficienteSE+1, cy+h);
			printer.linha(ix_parcelasSE-1,cy, ix_parcelasSE-1,cy+h);
			printer.linha(ix_coeficienteSE+wx_coeficienteSE+1, cy, ix_coeficienteSE+wx_coeficienteSE+1, cy+h);
			s='Sem Entrada';
			printer.imprime(ix_parcelasSE+(wx_parcelasSE+wx_coeficienteSE-printer.texto_largura(s))/2,cy,s);
			cy=cy+h;
		//	2ª LINHA DE TÍTULO
		//	COM ENTRADA
			printer.linha(ix_parcelasCE-1, cy, ix_coeficienteCE+wx_coeficienteCE+1, cy);
			printer.linha(ix_parcelasCE-1, cy+h, ix_coeficienteCE+wx_coeficienteCE+1, cy+h);
			printer.linha(ix_parcelasCE-1,cy, ix_parcelasCE-1,cy+h);
			printer.linha(ix_coeficienteCE-1, cy, ix_coeficienteCE-1, cy+h);
			printer.linha(ix_coeficienteCE+wx_coeficienteCE+1, cy, ix_coeficienteCE+wx_coeficienteCE+1, cy+h);
			s='Parcelas';
			printer.imprime(ix_parcelasCE+(wx_parcelasCE-printer.texto_largura(s))/2,cy,s);
			s='Coeficiente';
			printer.imprime(ix_coeficienteCE+wx_coeficienteCE-printer.texto_largura(s),cy,s);
		//	SEM ENTRADA
			printer.linha(ix_parcelasSE-1, cy, ix_coeficienteSE+wx_coeficienteSE+1, cy);
			printer.linha(ix_parcelasSE-1, cy+h, ix_coeficienteSE+wx_coeficienteSE+1, cy+h);
			printer.linha(ix_parcelasSE-1,cy, ix_parcelasSE-1,cy+h);
			printer.linha(ix_coeficienteSE-1, cy, ix_coeficienteSE-1, cy+h);
			printer.linha(ix_coeficienteSE+wx_coeficienteSE+1, cy, ix_coeficienteSE+wx_coeficienteSE+1, cy+h);
			s='Parcelas';
			printer.imprime(ix_parcelasSE+(wx_parcelasSE-printer.texto_largura(s))/2,cy,s);
			s='Coeficiente';
			printer.imprime(ix_coeficienteSE+wx_coeficienteSE-printer.texto_largura(s),cy,s);
			
			cy=cy+h;
			printer.fontbold=false;
			}

		h=printer.texto_altura('X');
		
	//	COM ENTRADA
		printer.linha(ix_parcelasCE-1,cy,ix_parcelasCE-1,cy+h);
		printer.linha(ix_coeficienteCE-1, cy, ix_coeficienteCE-1, cy+h);
		printer.linha(ix_coeficienteCE+wx_coeficienteCE+1, cy, ix_coeficienteCE+wx_coeficienteCE+1, cy+h);
		printer.linha(ix_parcelasCE-1, cy+h, ix_coeficienteCE+wx_coeficienteCE+1, cy+h);
		
		cx=ix_parcelasCE+(wx_parcelasCE-printer.texto_largura(Pd[iv].parcelasCE))/2;
		printer.imprime_campo(cx,cy,wx_parcelasCE,Pd[iv].parcelasCE);
		
		cx=ix_coeficienteCE+wx_coeficienteCE-printer.texto_largura(Pd[iv].coeficienteCE);
		printer.imprime_campo(cx,cy,wx_coeficienteCE,Pd[iv].coeficienteCE);
		
	//	SEM ENTRADA
		printer.linha(ix_parcelasSE-1,cy,ix_parcelasSE-1,cy+h);
		printer.linha(ix_coeficienteSE-1, cy, ix_coeficienteSE-1, cy+h);
		printer.linha(ix_coeficienteSE+wx_coeficienteSE+1, cy, ix_coeficienteSE+wx_coeficienteSE+1, cy+h);
		printer.linha(ix_parcelasSE-1, cy+h, ix_coeficienteSE+wx_coeficienteSE+1, cy+h);
		
		cx=ix_parcelasSE+(wx_parcelasSE-printer.texto_largura(Pd[iv].parcelasSE))/2;
		printer.imprime_campo(cx,cy,wx_parcelasSE,Pd[iv].parcelasSE);
		
		cx=ix_coeficienteSE+wx_coeficienteSE-printer.texto_largura(Pd[iv].coeficienteSE);
		printer.imprime_campo(cx,cy,wx_coeficienteSE,Pd[iv].coeficienteSE);
		
		cy=cy+h;
		}
		
	printer.EndDoc();
	alert('Tabela de custo financeiro foi impressa!!');
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

<form id="fTabCustoFinancFornec" name="fTabCustoFinancFornec" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Tabela de Custo Financeiro</span>
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


<!--  TABELA DE CUSTO FINANCEIRO  -->
<% tabela_impressao_monta %>

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
		<a name="bIMPRESSORA" id="bIMPRESSORA" href="javascript:fTabCustoFinancFornecImpressora(fTabCustoFinancFornec)" title="seleciona a impressora">
		<img src="../botao/impressora.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="center"><div name="dMARGENS" id="dMARGENS">
		<a name="bMARGENS" id="bMARGENS" href="javascript:fTabCustoFinancFornecMargens(fTabCustoFinancFornec)" title="configura as margens de impressão">
		<img src="../botao/margens.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dIMPRIME" id="dIMPRIME">
		<a name="bIMPRIME" id="bIMPRIME" href="javascript:fTabCustoFinancFornecImprime(fTabCustoFinancFornec)" title="imprime a listagem em formulário contínuo">
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
