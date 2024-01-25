<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  O R D E M S E R V I C O E D I T A . A S P
'     =============================================================
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

	dim s, usuario, i, n, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim url_back
	url_back = Trim(request("url_back"))

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim max_qtde_itens
	max_qtde_itens = obtem_parametro_OrdemServico_Volumes_MaxQtdeItens

'	OBTÉM O NÚMERO DA ORDEM DE SERVIÇO
	dim s_num_OS
	s_num_OS = Ucase(Trim(Request.Form("c_num_OS")))

	dim s_descricao_volume, s_num_serie, s_tipo_vol, s_obs_problema
	
	dim alerta
	alerta=""

'	CONSISTE DADOS DIGITADOS
	if s_num_OS = "" then
		alerta = "Nº da ordem de serviço não foi informado."
		end if

	dim r_OS, r_OS_item
	if alerta = "" then
		if Not le_ordem_servico(s_num_OS, r_OS, msg_erro) then 
			alerta = msg_erro
		else
			if Not le_ordem_servico_item(s_num_OS, r_OS_item, msg_erro) then alerta = msg_erro
			'Assegura que dados cadastrados anteriormente sejam exibidos corretamente, mesmo se o parâmetro da quantidade máxima de itens tiver sido reduzido
			if VectorLength(r_OS_item) > max_qtde_itens then max_qtde_itens = VectorLength(r_OS_item)
			end if
		end if

	dim r_cliente, s_nome_contato
	set r_cliente = New cl_CLIENTE
	s_nome_contato = ""
	if alerta = "" then
		if r_OS.id_cliente <> "" then
			if x_cliente_bd(r_OS.id_cliente, r_cliente) then
				s_nome_contato = Trim(r_cliente.contato)
				if s_nome_contato <> "" then 
					s_nome_contato = "  (contato: " & s_nome_contato & ")"
					end if
				end if
			end if
		end if
	
	dim r_orcamentista_e_indicador
	dim s_telefone_indicador, s_tel_aux_1, s_tel_aux_2
	s_telefone_indicador = ""
	if alerta = "" then
		if r_OS.indicador <> "" then
			if le_orcamentista_e_indicador(r_OS.indicador, r_orcamentista_e_indicador, msg_erro) then
				with r_orcamentista_e_indicador
					s_tel_aux_1 = formata_ddd_telefone_ramal(.ddd, .telefone, "")
					s_tel_aux_2 = formata_ddd_telefone_ramal(.ddd_cel, .tel_cel, "")
					if (s_tel_aux_1 <> "") And (s_tel_aux_2 <> "") then
						s_telefone_indicador = s_tel_aux_1 & " / " & s_tel_aux_2
					else
						s_telefone_indicador = s_tel_aux_1 & s_tel_aux_2
						end if
					if s_telefone_indicador <> "" then s_telefone_indicador = "  (Tel: " & s_telefone_indicador & ")"
					end with
				end if
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



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fOSConfirma( f ) {
var i,blnTemItem,blnTemInfo;
	blnTemItem=false;
	for (i = 0; i < f.c_descricao_volume.length; i++) {
		blnTemInfo = false;
		if (trim(f.c_descricao_volume[i].value) != "") blnTemInfo = true;
		if (trim(f.c_num_serie[i].value) != "") blnTemInfo = true;
		if (trim(f.c_tipo[i].value) != "") blnTemInfo = true;
		if (trim(f.c_obs_problema[i].value) != "") blnTemInfo = true;

		if (blnTemInfo) {
			blnTemItem = true;
			if (trim(f.c_descricao_volume[i].value) == "") {
				alert("Informe a descrição do volume!!");
				f.c_descricao_volume[i].focus();
				return;
			}
			if (trim(f.c_num_serie[i].value) == "") {
				alert("Informe o número de série do volume!!");
				f.c_num_serie[i].focus();
				return;
			}
			if (trim(f.c_obs_problema[i].value) == "") {
				alert("Informe o problema!!");
				f.c_obs_problema[i].focus();
				return;
			}
			if (f.c_obs_problema[i].value.length > MAX_TAM_OS_OBS_PROBLEMA) {
				alert("A descrição do problema excede o tamanho máximo de " + MAX_TAM_OS_OBS_PROBLEMA + " caracteres!\nConteúdo possui " + f.c_obs_problema[i].value.length + " caracteres!");
				f.c_obs_problema[i].focus();
				return;
			}
		}
	}
	
	if (!blnTemItem) {
		alert("Não foi informado nenhum volume na lista!!");
		f.c_descricao_volume[0].focus();
		return;
		}

	if (f.c_obs_pecas_necessarias.value.length > MAX_TAM_OS_OBS_PECAS_NECESSARIAS) {
		alert("A descrição de peças necessárias excede o tamanho máximo de " + MAX_TAM_OS_OBS_PECAS_NECESSARIAS + " caracteres!\nConteúdo possui " + f.c_obs_pecas_necessarias.value.length + " caracteres!");
		f.c_obs_pecas_necessarias.focus();
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
<!-- *************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS DE CONFIRMAÇÃO  ********** -->
<!-- *************************************************************** -->
<body onload="focus();">
<center>

<form id="fOP" name="fOP" method="post" action="OrdemServicoAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_num_OS" id="c_num_OS" value="<%=s_num_OS%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="749" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td valign="bottom" width="33%"><p class="STP" style="color:<%=x_OS_status_cor(r_OS.situacao_status)%>;"
	><%=Ucase(x_OS_status(r_OS.situacao_status))%>
<%	s = ""
	if r_OS.situacao_status <> ST_OS_EM_ANDAMENTO then
		s=formata_data(r_OS.situacao_data)
		if s<>"" then s="  (" & s & ")"
		end if
%>
	<%=s%></p></td>
	<td valign="bottom" align="center" width="33%">
		<p class="STP"><%=formata_data(r_OS.data)%>
		<span class="HoraPed"><%=formata_hhnnss_para_hh_nn(r_OS.hora)%></span>
		</p></td>
	<td valign="bottom" align="right" NOWRAP><p class="PEDIDO">Ordem de Serviço nº&nbsp;<%=formata_num_OS_tela(s_num_OS)%></p></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0">
	<!--  TÍTULO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>ORDEM DE SERVIÇO</span></td>
	</tr>
<!--  CADASTRADO POR  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Cadastrado por</span></td>
	<td class="MDB" colspan="4">
		<input name="c_cadastrado_por" id="c_cadastrado_por" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=r_OS.usuario%>"></td>
	</tr>
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Estoque origem</span></td>
	<td class="MDB" colspan="4">
		<input name="c_estoque_origem_aux" id="c_estoque_origem_aux" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=x_estoque(r_OS.cod_estoque_origem)%>"></td>
	</tr>
	<% if r_OS.loja_estoque_origem <> "" then %>
		<%	s = r_OS.loja_estoque_origem & " - " & x_loja(r_OS.loja_estoque_origem) %>
		<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Loja origem</span></td>
		<td class="MDB" colspan="4">
			<input name="c_loja_aux" id="c_loja_aux" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
					value="<%=s%>"></td>
		</tr>
	<% end if %>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Empresa (CD)</span></td>
		<td class="MDB" colspan="4">
			<input name="c_id_nfe_emitente_aux" id="c_id_nfe_emitente_aux" tabindex="-1" class="PLLe" style="width:460px;margin-left:2pt;" value="<%=obtem_apelido_empresa_NFe_emitente(r_OS.id_nfe_emitente)%>" />
		</td>
	</tr>

<tr><td colspan="5">&nbsp;</td></tr>

<% if r_OS.pedido <> "" then %>

	<!--  TÍTULO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>DADOS DO PEDIDO</span></td>
	</tr>
<!--  PEDIDO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Pedido</span></td>
	<td class="MDB" colspan="4">
		<input name="c_pedido_aux" id="c_pedido_aux" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=r_OS.pedido%>"></td>
	</tr>
<!--  NF  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">NF</span></td>
	<td class="MDB" colspan="4">
		<input name="c_nf" id="c_nf" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=r_OS.nf%>"></td>
	</tr>
<!--  NOME DO CLIENTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Cliente</span></td>
	<td class="MDB" colspan="4">
		<input name="c_nome_cliente" id="c_nome_cliente" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=r_OS.nome_cliente & s_nome_contato%>"></td>
	</tr>
<!--  ENDEREÇO  -->
	<%
	with r_OS
		s = formata_endereco(.endereco, .endereco_numero, .endereco_complemento, .bairro, .cidade, .uf, .cep)
		end with
	%>
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right" valign="top"><span class="PLTe" style="vertical-align:middle;">Endereço</span></td>
	<td class="MDB" colspan="4">
		<textarea rows="<%=Cstr(MAX_LINHAS_OS_ENDERECO)%>" name="c_endereco" id="c_endereco" READONLY tabindex=-1 class="PLLe" style="width:100%;margin-left:2pt;"><%=s%></textarea></td>
	</tr>
<% if r_OS.tipo_cliente = ID_PF then %>
<!--  TELEFONE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Tel Res</span></td>
	<td class="MDB" colspan="4">
		<input name="c_tel_res" id="c_tel_res" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=formata_ddd_telefone_ramal(r_OS.ddd_res, r_OS.tel_res, "")%>"></td>
	</tr>
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Tel Com</span></td>
	<td class="MDB" colspan="4">
		<input name="c_tel_com" id="c_tel_com" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=formata_ddd_telefone_ramal(r_OS.ddd_com, r_OS.tel_com, r_OS.ramal_com)%>"></td>
	</tr>
<% else %>
<!--  TELEFONE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Telefone</span></td>
	<td class="MDB" colspan="4">
		<input name="c_telefone" id="c_telefone" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=formata_ddd_telefone_ramal(r_OS.ddd_com, r_OS.tel_com, r_OS.ramal_com)%>"></td>
	</tr>
<%end if%>
<!--  INDICADOR  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Indicador</span></td>
	<td class="MDB" colspan="4">
		<input name="c_indicador" id="c_indicador" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=r_OS.indicador & s_telefone_indicador%>"></td>
	</tr>

<tr><td colspan="5">&nbsp;</td></tr>
<% end if %>

<!--  P R O D U T O  -->
	<!--  TÍTULO DA TABELA  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>PRODUTO</span></td>
	</tr>
	<!--  TÍTULO DAS COLUNAS  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE"><p class="PLTe">Fabr</p></td>
	<td class="MDB"><p class="PLTe">Produto</p></td>
	<td class="MDB"><p class="PLTe">EAN</p></td>
	<td class="MDB"><p class="PLTe">Descrição</p></td>
	<td class="MDB"><p class="PLTd">Qtde</p></td>
	</tr>

<%	i=1 %>
	<tr>
	<td class="MDBE">
		<input name="c_fabricante" id="c_fabricante" READONLY tabindex=-1 class="PLLe" style="width:30px;"
			value="<%=r_OS.fabricante%>"></td>
	<td class="MDB">
		<input name="c_produto" id="c_produto" READONLY tabindex=-1 class="PLLe" style="width:55px;"
			value="<%=r_OS.produto%>"></td>
	<td class="MDB">
		<input name="c_ean" id="c_ean" READONLY tabindex=-1 class="PLLe" style="width:85px;"
			value="<%=r_OS.ean%>"></td>
	<td class="MDB" style="width:277px;">
		<span class="PLLe"><%=produto_formata_descricao_em_html(r_OS.descricao_html)%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value="<%=r_OS.descricao%>">
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" READONLY tabindex=-1 class="PLLd" style="width:35px;"
			value="<%=Cstr(r_OS.qtde)%>"></td>
	</tr>


<tr><td colspan="5">&nbsp;</td></tr>
	<!--  TÍTULO DA TABELA  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>VOLUMES</span></td>
	</tr>

<!--  R E L A Ç Ã O   D E   V O L U M E S  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE"><p class="PLTe">Volume</p></td>
	<td class="MDB"><p class="PLTe">Tipo</p></td>
	<td class="MDB"><p class="PLTe">Nº Série</p></td>
	<td class="MDB" colspan="2"><p class="PLTe">Problema</p></td>
	</tr>
<%  n = Lbound(r_OS_item)-1
	for i=1 to max_qtde_itens
		n = n+1
		if n <= Ubound(r_OS_item) then
			with r_OS_item(n)
				s_descricao_volume = .descricao_volume
				s_tipo_vol = .tipo
				s_num_serie = .num_serie
				s_obs_problema = .obs_problema
				end with
		else
			s_descricao_volume = ""
			s_tipo_vol = ""
			s_num_serie = ""
			s_obs_problema = ""
			end if
%>
	<tr>
	<td class="MDBE" valign="top"><input name="c_descricao_volume" id="c_descricao_volume" 
		class="PLLe" maxlength="12" 
		style="width:100px;" onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(i)%>!=1))) if (trim(this.value)=='') bCONFIRMA.focus(); else fOP.c_tipo[<%=Cstr(i-1)%>].focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
		value='<%=s_descricao_volume%>'></td>
	<td class="MDB" valign="top"><input name="c_tipo" id="c_tipo" 
		class="PLLe" maxlength="12" 
		style="width:100px;" onkeypress="if (digitou_enter(true)) fOP.c_num_serie[<%=Cstr(i-1)%>].focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
		value='<%=s_tipo_vol%>'></td>
	<td class="MDB" valign="top"><input name="c_num_serie" id="c_num_serie" 
		class="PLLe" maxlength="20" 
		style="width:130px;" onkeypress="if (digitou_enter(true)) fOP.c_obs_problema[<%=Cstr(i-1)%>].focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
		value='<%=s_num_serie%>'></td>
	<td class="MDB" colspan="2" align="right" style="width:344px;"><textarea name="c_obs_problema" id="c_obs_problema" rows="<%=Cstr(MAX_LINHAS_OS_OBS_PROBLEMA)%>" 
		class="PLLe" style="width:340px;" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fOP.c_obs_problema.length) bCONFIRMA.focus(); else fOP.c_descricao_volume[<%=Cstr(i)%>].focus();} filtra_nome_identificador();"
		><%=s_obs_problema%></textarea></td>

	<input type="hidden" name="c_descricao_volume_original" id="c_descricao_volume_original" value='<%=s_descricao_volume%>'>
	<input type="hidden" name="c_tipo_original" id="c_tipo_original" value='<%=s_tipo_vol%>'>
	<input type="hidden" name="c_num_serie_original" id="c_num_serie_original" value='<%=s_num_serie%>'>
	<input type="hidden" name="c_obs_problema_original" id="c_obs_problema_original" value='<%=s_obs_problema%>'>
	</tr>
<% next %>


<tr><td colspan="5">&nbsp;</td></tr>

<!--  PEÇAS NECESSÁRIAS  -->
	<!--  TÍTULO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>PEÇAS NECESSÁRIAS</span></td>
	</tr>
	<tr>
	<td colspan="5" class="MDBE" align="right" style="width:685px;"><textarea name="c_obs_pecas_necessarias" id="c_obs_pecas_necessarias" rows="<%=Cstr(MAX_LINHAS_OS_OBS_PECAS_NECESSARIAS)%>" 
		class="PLLe" style="width:685px;" onkeypress="filtra_nome_identificador();"
		><%=r_OS.obs_pecas_necessarias%></textarea></td>
	</tr>
</table>



<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="749" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="749" cellSpacing="0">
<tr>
	<%	if url_back <> "" then 
			s="resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
		else 
			s="javascript:history.back()"
			end if
	%>
	<td><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fOSConfirma(fOP)" title="grava as alterações na Ordem de Serviço">
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