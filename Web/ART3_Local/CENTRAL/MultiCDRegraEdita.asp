<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================
'	  MultiCDRegraEdita.asp
'     ===============================
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
	
'	OBTEM O ID
	dim s, s_aux, usuario, operacao_selecionada, id_selecionado, apelido_selecionado
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	REGRA A EDITAR
	operacao_selecionada = Trim(request("operacao_selecionada"))
	id_selecionado = Ucase(Trim(request("id_selecionado")))
	apelido_selecionado = Trim(request("apelido_selecionado"))
	
	if operacao_selecionada=OP_INCLUI then
		apelido_selecionado = filtra_nome_identificador(apelido_selecionado)
		if apelido_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_MULTI_CD_REGRA_APELIDO_NAO_INFORMADO)
	else
		if id_selecionado="" then Response.Redirect("aviso.asp?id=" & ERR_MULTI_CD_REGRA_NAO_ESPECIFICADA)
		end if
		
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if Not operacao_permitida(OP_CEN_MULTI_CD_CADASTRO_REGRAS_CONSUMO_ESTOQUE, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA COM O BANCO DE DADOS
	dim cn, tNE, tRegra, tRegraUf, tRegraUfPessoa, tRegraUfPessoaCd, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(tNE, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tRegra, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tRegraUf, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tRegraUfPessoa, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tRegraUfPessoaCd, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_cd_apelido, id_nfe_emitente_aux, s_table_id, s_row_id, s_id
	dim vCD, vCDAux, iCD, s_lista_cd
	s_lista_cd = ""
	redim vCD(0)
	set vCD(UBound(vCD)) = new cl_TRES_COLUNAS
	vCD(UBound(vCD)).c1 = 0

	s = "SELECT * FROM t_NFe_EMITENTE WHERE (st_habilitado_ctrl_estoque = 1) ORDER BY ordem, id"
	if tNE.State <> 0 then tNE.Close
	tNE.open s, cn
	do while Not tNE.Eof
		if vCD(UBound(vCD)).c1 <> 0 then
			redim preserve vCD(UBound(vCD)+1)
			set vCD(UBound(vCD)) = new cl_TRES_COLUNAS
			end if
		
		if s_lista_cd <> "" then s_lista_cd = s_lista_cd & "|"
		s_lista_cd = s_lista_cd & Trim("" & tNE("id"))
		vCD(UBound(vCD)).c1 = CLng(tNE("id"))
		vCD(UBound(vCD)).c2 = Trim("" & tNE("apelido"))
		vCD(UBound(vCD)).c3 = CLng(tNE("st_ativo"))

		tNE.MoveNext
		loop

	dim qtde_produtos_associados
	qtde_produtos_associados = 0
	if operacao_selecionada=OP_CONSULTA then
		s = "SELECT Count(*) AS qtde FROM t_PRODUTO_X_WMS_REGRA_CD WHERE (id_wms_regra_cd = " & id_selecionado & ")"
		if tRegra.State <> 0 then tRegra.Close
		tRegra.open s, cn
		if Not tRegra.Eof then
			qtde_produtos_associados = tRegra("qtde")
			end if
		end if

	if operacao_selecionada=OP_INCLUI then
		s = "SELECT * FROM t_WMS_REGRA_CD WHERE (apelido = '" & QuotedStr(apelido_selecionado) & "')"
	else
		s = "SELECT * FROM t_WMS_REGRA_CD WHERE (id = " & id_selecionado & ")"
		end if
	if tRegra.State <> 0 then tRegra.Close
	tRegra.open s, cn
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

	if operacao_selecionada=OP_INCLUI then
		if Not tRegra.EOF then Response.Redirect("aviso.asp?id=" & ERR_MULTI_CD_REGRA_JA_CADASTRADA)
	elseif operacao_selecionada=OP_CONSULTA then
		if tRegra.EOF then Response.Redirect("aviso.asp?id=" & ERR_MULTI_CD_REGRA_NAO_CADASTRADA)
		apelido_selecionado = Trim("" & tRegra("apelido"))
		end if

	dim vUF, iUF
	vUF = UF_get_array

	dim vPessoa, iPessoa
	vPessoa = TipoPessoa_get_array

	dim s_checked
	dim idxOrdem
	dim spe_id_nfe_emitente
	dim idRegraUf, idRegraUfPessoa, idRegraUfPessoaCd

%>


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>


<%
'		C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE
'
%>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	var clip_st_conteudo = false;
	var clip_uf_base;
	var clip_ckb_uf_desativada;
	var clip_vPessoa = [];

	function cria_pessoa(p_tipo_pessoa, p_ckb_pessoa_desativada, p_spe_id_nfe_emitente, p_vCD) {
		this.tipo_pessoa = p_tipo_pessoa;
		this.ckb_pessoa_desativada = p_ckb_pessoa_desativada;
		this.spe_id_nfe_emitente = p_spe_id_nfe_emitente;
		this.vCD = p_vCD;
	}

	function cria_CD(p_id_nfe_emitente, p_ckb_cd_desativado, p_ordem_prioridade){
		this.id_nfe_emitente = p_id_nfe_emitente;
		this.ckb_cd_desativado = p_ckb_cd_desativado;
		this.ordem_prioridade = p_ordem_prioridade;
	}

	function clip_copy(uf) {
		var p_ckb_pessoa_desativada;
		var p_spe_id_nfe_emitente;
		var s_id;
		var s_table_id;
		var ckb;
		var c_spe_list;
		var c_tbl_cd;

		try {
			// A VARIÁVEL ARRAY É UM PONTEIRO, SE NÃO FOR CRIADA UMA P/ CADA TIPO DE PESSOA, TODAS IRÃO APONTAR P/ O MESMO ARRAY
			<% for iPessoa=LBound(vPessoa) to UBound(vPessoa)%>
			var p_vCD_<%=vPessoa(iPessoa)%> = [];
			<% next %>
		
			clip_st_conteudo = true;
			clip_uf_base = uf;

			s_id = "ckb_UF_" + uf;
			ckb = document.getElementById(s_id);
			clip_ckb_uf_desativada = ckb.checked;

			<% for iPessoa=LBound(vPessoa) to UBound(vPessoa)%>
			// CHECK BOX "PESSOA DESATIVADA"
			s_id = "ckb_pessoa_" + uf + "_" + "<%=vPessoa(iPessoa)%>";
			ckb = document.getElementById(s_id);
			p_ckb_pessoa_desativada = ckb.checked;
			// CD SELECIONADO P/ SPE
			s_id = "c_cd_spe_" + uf + "_" + "<%=vPessoa(iPessoa)%>";
			c_spe_list = document.getElementById(s_id);
			p_spe_id_nfe_emitente = c_spe_list.value;
			// TABELA DE PRIORIZAÇÃO DOS CD'S P/ PRODUTOS DISPONÍVEIS
			s_table_id = "#tbl_" + uf + "_" + "<%=vPessoa(iPessoa)%>";
			$(s_table_id).find(".TableRow").each(function (i) {
				var p_id_nfe_emitente;
				var p_ckb_cd_desativado;
				var p_ordem_prioridade;
				var c;
				// CÓDIGO DE ID_NFE_EMITENTE DA LINHA EM QUESTÃO
				s_id = $(this).find(".CampoIdNfeEmitente").attr("id");
				c = document.getElementById(s_id);
				p_id_nfe_emitente = c.value;
				// CHECK BOX INDICANDO SE O CD ESTÁ ATIVADO/DESATIVADO
				s_id = "ckb_cd_" + uf + "_" + "<%=vPessoa(iPessoa)%>" + "_" + p_id_nfe_emitente;
				c = document.getElementById(s_id);
				p_ckb_cd_desativado = c.checked;
				// Nº DE ORDEM DA PRIORIDADE DA LINHA
				s_id = "c_ordem_" + uf + "_" + "<%=vPessoa(iPessoa)%>" + "_" + p_id_nfe_emitente;
				c = document.getElementById(s_id);
				p_ordem_prioridade = c.value;
				// ARMAZENA OS DADOS DA LINHA EM UMA POSIÇÃO DO VETOR vCD[]
				p_vCD_<%=vPessoa(iPessoa)%>[i] = new cria_CD(p_id_nfe_emitente, p_ckb_cd_desativado, p_ordem_prioridade);
			});

			// ARMAZENA OS DADOS DA 'PESSOA' NO VETOR vPessoa[]
			clip_vPessoa["<%=vPessoa(iPessoa)%>"] = new cria_pessoa("<%=vPessoa(iPessoa)%>", p_ckb_pessoa_desativada, p_spe_id_nfe_emitente, p_vCD_<%=vPessoa(iPessoa)%>);
			<% next %>

			alert("Configurações da UF '" + uf + "' copiadas com sucesso!");
		} catch (e) {
			// NOP
		}
	}

	function clip_paste(uf) {
		var s_id;

		try {
			if (!clip_st_conteudo) {
				alert("Não há dados copiados no clipboard!");
				return;
			}

			if (clip_uf_base == uf) {
				alert("A UF de origem e destino da cópia é a mesma!")
				return;
			}

			if (!confirm("Copiar as configurações da UF '" + clip_uf_base + "' para '" + uf + "'?")) return;

			// CHECK BOX "UF DESATIVADA"
			s_id = "#ckb_UF_"+uf;
			$(s_id).prop("checked", clip_ckb_uf_desativada);
			$(s_id).change();

			// vPessoa é um 'associative array'
			for (var keyP in clip_vPessoa){
				// CHECK BOX "PESSOA DESATIVADA"
				s_id = "#ckb_pessoa_" + uf + "_" + clip_vPessoa[keyP].tipo_pessoa;
				$(s_id).prop("checked", clip_vPessoa[keyP].ckb_pessoa_desativada);
				// GARANTE QUE A COR DO LABEL FIQUE DE ACORDO
				$(s_id).change();

				// ID_NFE_EMITENTE SELECIONADO NA LISTA DE CD'S P/ O SPE
				s_id = "#c_cd_spe_" + uf + "_" + clip_vPessoa[keyP].tipo_pessoa;
				$(s_id).val(clip_vPessoa[keyP].spe_id_nfe_emitente);

				// TABELA DE PRIORIZAÇÃO DOS CD'S P/ PRODUTOS DISPONÍVEIS
				for (var keyCD in clip_vPessoa[keyP].vCD){
					var s_table_id = "#tbl_" + uf + "_" + clip_vPessoa[keyP].tipo_pessoa;
					$(s_table_id).find(".TableRow").each(function (i) {
						var s_row_id = "#" + $(this).attr("id");
						var s_id = $(this).find(".CampoIdNfeEmitente").attr("id");
						var c = document.getElementById(s_id);
						var id_nfe_emitente_aux = c.value;
						if (id_nfe_emitente_aux.toString() == clip_vPessoa[keyP].vCD[keyCD].id_nfe_emitente.toString())
						{
							// CHECK BOX "DESATIVADO"
							s_id = "#ckb_cd_" + uf + "_" + clip_vPessoa[keyP].tipo_pessoa + "_" + id_nfe_emitente_aux.toString();
							$(s_id).prop("checked", clip_vPessoa[keyP].vCD[keyCD].ckb_cd_desativado);
							// GARANTE QUE A COR DO LABEL FIQUE DE ACORDO
							$(s_id).change();
							// CONSIDERANDO QUE AS LINHAS ARMAZENADAS NO CLIPBOARD ESTÃO ORDENADAS, COLOCA CADA LINHA NA ÚLTIMA POSIÇÃO DO DESTINO, DESTA FORMA, O RESULTADO FINAL IRÁ FICAR ORDENADO
							var s_last_row_id = "#" + $(s_table_id + " tr:last").attr("id");
							$(s_last_row_id).after($(s_row_id));
							return;
						}
					});
					// Reorganiza nº ordenação
					$(s_table_id).find(".TableRow").each(function (i) {
						var s_id = $(this).find(".CampoOrdenacao").attr("id");
						var c = document.getElementById(s_id);
						c.value = (i+1).toString();
					});
				}
			}

			alert("Configurações coladas com sucesso!");
		} catch (e) {
			// NOP
		}
	}
</script>

<script type="text/javascript">
	$(function () {
		$(".CkbRegra").each(function () {
			var s_id_span = "#spn_regra";
			if ($(this).is(":checked")) {
				$(s_id_span).addClass("ColorRed");
			}
			else {
				$(s_id_span).removeClass("ColorRed");
			}
		});

		$(".CkbRegra").change(function () {
			var s_id_span = "#spn_regra";
			if ($(this).is(":checked")) {
				$(s_id_span).addClass("ColorRed");
			}
			else {
				$(s_id_span).removeClass("ColorRed");
			}
		});

		$(".CkbUF").each(function () {
			var s_ckb_value = $(this).val();
			var s_id_span = "#spn_UF_" + s_ckb_value;
			if ($(this).is(":checked")) {
				$(s_id_span).addClass("ColorRed");
			}
			else {
				$(s_id_span).removeClass("ColorRed");
			}
		});

		$(".CkbUF").change(function () {
			var s_ckb_value = $(this).val();
			var s_id_span = "#spn_UF_" + s_ckb_value;
			if ($(this).is(":checked")) {
				$(s_id_span).addClass("ColorRed");
			}
			else {
				$(s_id_span).removeClass("ColorRed");
			}
		});

		$(".CkbPessoa").each(function () {
			var s_ckb_value = $(this).val();
			var s_id_span = "#spn_pessoa_" + s_ckb_value;
			if ($(this).is(":checked")) {
				$(s_id_span).addClass("ColorRed");
			}
			else {
				$(s_id_span).removeClass("ColorRed");
			}
		});

		$(".CkbPessoa").change(function () {
			var s_ckb_value = $(this).val();
			var s_id_span = "#spn_pessoa_" + s_ckb_value;
			if ($(this).is(":checked")) {
				$(s_id_span).addClass("ColorRed");
			}
			else {
				$(s_id_span).removeClass("ColorRed");
			}
		});

		$(".CkbCd").each(function () {
			var s_ckb_value = $(this).val();
			var s_id_span = "#spn_cd_" + s_ckb_value;
			if ($(this).is(":checked")) {
				$(s_id_span).addClass("ColorRed");
			}
			else {
				$(s_id_span).removeClass("ColorRed");
			}
		});

		$(".CkbCd").change(function () {
			var s_ckb_value = $(this).val();
			var s_id_span = "#spn_cd_" + s_ckb_value;
			if ($(this).is(":checked")) {
				$(s_id_span).addClass("ColorRed");
			}
			else {
				$(s_id_span).removeClass("ColorRed");
			}
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
var qtde_produtos_associados = <%=qtde_produtos_associados%>;

function CDMoverParaBaixo(table_id, row_id) {
	var s_table_id = "#" + table_id;
	var s_row_id = "#" + row_id;
	$(s_row_id).next().after($(s_row_id));
	// Reorganiza nº ordenação
	$(s_table_id).find(".TableRow").each(function (i) {
		var s_id = $(this).find(".CampoOrdenacao").attr("id");
		var c = document.getElementById(s_id);
		c.value = (i+1).toString();
	});
}

function CDMoverParaCima(table_id, row_id)
{
	var s_table_id = "#" + table_id;
	var s_row_id = "#" + row_id;
	$(s_row_id).prev().before($(s_row_id));
	// Reorganiza nº ordenação
	$(s_table_id).find(".TableRow").each(function (i) {
		var s_id = $(this).find(".CampoOrdenacao").attr("id");
		var c = document.getElementById(s_id);
		c.value = (i+1).toString();
	});
}

function expandirTudo() {
	$(".TrBody").each(function () {
		var s_row_id = $(this).attr("id");
		// OBTÉM OS 2 ÚLTIMOS CARACTERES
		var uf = s_row_id.substr(s_row_id.length - 2);
		var row_MORE_INFO;
		row_MORE_INFO = document.getElementById(s_row_id);
		row_MORE_INFO.style.display = "";
		$(".TdHeader_" + uf).removeClass("MB");
		$(".TdBody_" + uf).addClass("MB");
	});
}

function recolherTudo() {
	$(".TrBody").each(function () {
		var s_row_id = $(this).attr("id");
		// OBTÉM OS 2 ÚLTIMOS CARACTERES
		var uf = s_row_id.substr(s_row_id.length - 2);
		var row_MORE_INFO;
		row_MORE_INFO = document.getElementById(s_row_id);
		row_MORE_INFO.style.display = "none";
		$(".TdHeader_" + uf).addClass("MB");
		$(".TdBody_" + uf).removeClass("MB");
	});
}

function fExibeOcultaCampos(uf) {
var row_MORE_INFO;
	row_MORE_INFO = document.getElementById("row_body_" + uf);
	if (row_MORE_INFO.style.display.toString() == "none") {
		row_MORE_INFO.style.display = "";
		$(".TdHeader_" + uf).removeClass("MB");
		$(".TdBody_" + uf).addClass("MB");
	}
	else {
		row_MORE_INFO.style.display = "none";
		$(".TdHeader_" + uf).addClass("MB");
		$(".TdBody_" + uf).removeClass("MB");
	}
}

function RemoveRegra( f ) {
var b, s;
	s = "Confirma a exclusão da regra?";
	if (qtde_produtos_associados > 0) s += "\n\nImportante: há " + qtde_produtos_associados.toString() +" produto(s) associado(s) a esta regra. Esses produtos ficarão sem definição de regra e, consequentemente, não poderão ser vendidos!";
	b=window.confirm(s);
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaRegra( f ) {
	if (trim(f.c_apelido.value)=="") {
		alert('Preencha o apelido da regra de consumo do estoque!');
		f.c_descricao.focus();
		return;
	}

	<% for iUF=LBound(vUF) to UBound(vUF) %>
	if (!f.ckb_UF_<%=vUF(iUF)%>.checked) {
		<% for iPessoa=LBound(vPessoa) to UBound(vPessoa)%>
		if (!f.ckb_pessoa_<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>.checked) {
			if (trim(f.c_cd_spe_<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>.value) == "") {
				alert("Selecione o CD a ser usado para os produtos 'Sem Presença no Estoque' para a UF '<%=vUF(iUF)%>' no caso de '<%=descricao_multi_CD_regra_tipo_pessoa(vPessoa(iPessoa))%>'!");
				return;
			}
		}
		<% next %>
	}
	<% next %>

	dATUALIZA.style.visibility="hidden";
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

<style TYPE="text/css">
#ckb_UF {
	margin: 0pt 2pt 1pt 15pt;
	}
#ckb_PESSOA {
	margin: 0pt 2pt 1pt 6pt;
	}
#ckb_CD {
	margin: 0pt 2pt 1pt 6pt;
	}
.ColorRed
{
	color:red;
}
.TdSeqCd
{
	min-width:120px;
}
.TdSeqCkb
{
	min-width: 90px;
}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.c_descricao.focus();"
	else
		s = "focus();"
		end if
%>
<body onload="<%=s%>">
<center>



<!--  CADASTRO DA REGRA DE CONSUMO DO ESTOQUE -->

<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Nova Regra de Consumo do Estoque"
	else
		s = "Consulta/Edição de Regra de Consumo do Estoque"
		end if
%>
	<td align="center" valign="bottom"><span class="PEDIDO"><%=s%></span></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="MultiCDRegraAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<input type="hidden" name="id_selecionado" id="id_selecionado" value='<%=id_selecionado%>'>
<input type="hidden" name="c_lista_cd" id="c_lista_cd" value="<%=s_lista_cd%>" />

<!-- ************   APELIDO   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td class="MD" width="50%">
			<p class="R">APELIDO</p><p class="C"><input type="text" id="c_apelido" name="c_apelido" class="TA" value="<%=apelido_selecionado%>" maxlength="30" size="30" style="text-align:left;"></p>
		</td>
		<td width="50%">
			<p class="R">STATUS</p>
			<%
				s_checked = ""
				if operacao_selecionada<>OP_INCLUI then
					if CLng(tRegra("st_inativo")) = 1 then s_checked = " checked"
					end if
			%>
			<input type="checkbox" id="ckb_regra_st_inativo" name="ckb_regra_st_inativo" value="ON" class="TA CkbRegra"<%=s_checked%>><span id="spn_regra" class="C" style="cursor:default;" onclick="fCAD.ckb_regra_st_inativo.click();">Regra Desativada</span>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="MC">
			<%
				s=""
				if operacao_selecionada<>OP_INCLUI then
					s = Trim("" & tRegra("descricao"))
					end if
			%>
			<p class="R">DESCRIÇÃO</p>
			<textarea id="c_descricao" name="c_descricao" class="PLLe" style="width:642px;margin-left:2pt;" maxlength="800" rows="6"><%=s%></textarea>
		</td>
	</tr>
</table>

<br />

<!-- ************   DEFINIÇÕES P/ CADA UF   ************ -->
<%
	for iUF=LBound(vUF) to UBound(vUF)
		if iUF > LBound(vUF) then Response.Write "<br />"
			s_checked = ""
			idRegraUf = 0
			if operacao_selecionada <> OP_INCLUI then
				s = "SELECT * FROM t_WMS_REGRA_CD_X_UF WHERE (id_wms_regra_cd = " & id_selecionado & ") AND (uf = '" & vUF(iUF) & "')"
				if tRegraUf.State <> 0 then tRegraUf.Close
				tRegraUf.open s, cn
				if Not tRegraUf.Eof then
					if CLng(tRegraUf("st_inativo")) = CLng(1) then s_checked = " checked"
					idRegraUf = tRegraUf("id")
					end if
				end if

%>
<table width="649" cellspacing="0">
	<tr>
		<td style="width:16px;">
			<a name='bExibeOcultaUF' id='bExibeOcultaUF_<%=vUF(iUF)%>' style="margin-right:4px;" href='javascript:fExibeOcultaCampos("<%=vUF(iUF)%>");' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>
		</td>
		<td style="width:20px;padding-left:8px;" class="MC ME TdHeader_<%=vUF(iUF)%>">
			<p class="R" style="color:black;font-size:14pt;"><%=vUF(iUF)%></p>
		</td>
		<td width="80%" class="MC TdHeader_<%=vUF(iUF)%>" style="padding-left:60px;">
			<a href="javascript:clip_copy('<%=vUF(iUF)%>')" title="Copia as configurações desta UF" tabindex="-1"><img src="../IMAGEM/edit_copy.png" /></a>
			&nbsp;&nbsp;
			<a href="javascript:clip_paste('<%=vUF(iUF)%>')" title="Cola as configurações nesta UF" tabindex="-1"><img src="../IMAGEM/edit_paste.png" /></a>
		</td>
		<td width="20%" align="right" class="MC MD TdHeader_<%=vUF(iUF)%>">
			<% s_id = "ckb_UF_" & vUF(iUF) %>
			<input type="checkbox" name="<%=s_id%>" id="<%=s_id%>" value="<%=vUF(iUF)%>" class="TA CkbUF"<%=s_checked%> /><span id="spn_UF_<%=vUF(iUF)%>" class="C SpnUF" style="cursor:default;margin-right:4px;" onclick="fCAD.ckb_UF_<%=vUF(iUF)%>.click();">UF Desativada</span>
		</td>
	</tr>
	<tr id="row_body_<%=vUF(iUF)%>" class="TrBody">
		<td style="width:16px;">&nbsp;</td>
		<td class="ME MB" style="width:20px;">&nbsp;</td>
		<td colspan="2" class="MD MB TdBody_<%=vUF(iUF)%>" style="padding-right:8px;padding-top:8px;padding-bottom:8px;">
		<!-- Lista de tipos de pessoa -->
		<% for iPessoa=LBound(vPessoa) to UBound(vPessoa)%>
		<table cellspacing="0" cellpadding="0" width="100%">
			<tr>
				<td width="30%" class="ME MC">
					<p class="R" style="color:black;font-size:10pt;"><%=descricao_multi_CD_regra_tipo_pessoa(vPessoa(iPessoa))%></p>
				</td>
				<td width="45%" class="MC">&nbsp;</td>
				<td width="25%" class="MD MC" align="right" style="padding-right:8px;">
					<%	spe_id_nfe_emitente = 0
						s_checked = ""
						s_id = "ckb_pessoa_" & vUF(iUF) & "_" & vPessoa(iPessoa)
						if operacao_selecionada <> OP_INCLUI then
							idRegraUfPessoa = 0
							s = "SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA WHERE (id_wms_regra_cd_x_uf = " & idRegraUf & ") AND (tipo_pessoa = '" & vPessoa(iPessoa) & "')"
							if tRegraUfPessoa.State <> 0 then tRegraUfPessoa.Close
							tRegraUfPessoa.open s, cn
							if Not tRegraUfPessoa.Eof then
								idRegraUfPessoa = tRegraUfPessoa("id")
								if CLng(tRegraUfPessoa("st_inativo")) = CLng(1) then s_checked = " checked"
								spe_id_nfe_emitente = tRegraUfPessoa("spe_id_nfe_emitente")
								end if
							end if
					%>
					<input type="checkbox" name="<%=s_id%>" id="<%=s_id%>" value="<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>" class="TA CkbPessoa CkbPessoa_<%=vUF(iUF)%>"<%=s_checked%> /><span id="spn_pessoa_<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>" class="C SpnPessoa_<%=vUF(iUF)%>" style="cursor:default;margin-right:4px;" onclick="fCAD.ckb_pessoa_<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>.click();">Pessoa Desativada</span>
				</td>
			</tr>
			<tr>
				<td colspan="3" class="ME MB MD" style="padding-top:8px;padding-bottom:8px;padding-right:8px;">
					<table cellspacing="0" cellpadding="0" width="100%">
						<tr>
							<!-- MARGEM -->
							<td style="min-width:8px;">&nbsp;</td>
							<td width="48%" class="ME MC MD" valign="top">
								<!-- Lista de CD's para os produtos disponíveis -->
								<p class="Rf">Produtos Disponíveis</p>
							</td>
							<!-- Espaçamento -->
							<td style="width:30px;">&nbsp;</td>
							<td width="48%" class="ME MC MD" valign="top">
								<!-- Seleção do CD para 'Sem Presença no Estoque' -->
								<p class="Rf">Sem Presença no Estoque</p>
							</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td class="ME MB MD" style="padding:8px;">
								<% s_table_id = "tbl_" & vUF(iUF) & "_" & vPessoa(iPessoa) %>
								<table id="<%=s_table_id%>" cellspacing="0" cellpadding="0" width="100%">
									<% 
										idxOrdem = 0
										redim vCDAux(0)
										set vCDAux(UBound(vCDAux)) = new cl_TRES_COLUNAS
										vCDAux(UBound(vCDAux)).c1 = 0
										
										if operacao_selecionada = OP_INCLUI then
											for iCD=LBound(vCD) to UBound(vCD)
												if vCDAux(UBound(vCDAux)).c1 <> 0 then
													redim preserve vCDAux(UBound(vCDAux)+1)
													set vCDAux(UBound(vCDAux)) = new cl_TRES_COLUNAS
													end if
												vCDAux(UBound(vCDAux)).c1 = vCD(iCD).c1
												vCDAux(UBound(vCDAux)).c2 = vCD(iCD).c2
												vCDAux(UBound(vCDAux)).c3 = vCD(iCD).c3
												next
										else
											s = "SELECT tRegraCd.*, tEmitente.apelido, tEmitente.st_ativo AS emitente_st_ativo FROM t_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD tRegraCd LEFT JOIN t_NFe_EMITENTE tEmitente ON (tRegraCd.id_nfe_emitente = tEmitente.id) WHERE (id_wms_regra_cd_x_uf_x_pessoa = " & idRegraUfPessoa & ") ORDER BY ordem_prioridade"
											if tRegraUfPessoaCd.State <> 0 then tRegraUfPessoaCd.Close
											tRegraUfPessoaCd.open s, cn
											do while Not tRegraUfPessoaCd.Eof
												if vCDAux(UBound(vCDAux)).c1 <> 0 then
													redim preserve vCDAux(UBound(vCDAux)+1)
													set vCDAux(UBound(vCDAux)) = new cl_TRES_COLUNAS
													end if

												vCDAux(UBound(vCDAux)).c1 = tRegraUfPessoaCd("id_nfe_emitente")
												vCDAux(UBound(vCDAux)).c2 = Trim("" & tRegraUfPessoaCd("apelido"))
												vCDAux(UBound(vCDAux)).c3 = tRegraUfPessoaCd("emitente_st_ativo")
												tRegraUfPessoaCd.MoveNext
												loop

											'VERIFICA SE HÁ CD NOVO CADASTRADO APÓS O CADASTRAMENTO DA REGRA
											s_aux = ""
											for iCD=LBound(vCDAux) to UBound(vCDAux)
												if vCDAux(iCD).c1 <> 0 then
													if s_aux <> "" then s_aux = s_aux & ","
													s_aux = s_aux & vCDAux(iCD).c1
													end if
												next
											if s_aux <> "" then
												s = "SELECT * FROM t_NFe_EMITENTE WHERE (st_habilitado_ctrl_estoque = 1) AND (id NOT IN (" & s_aux & ")) ORDER BY ordem, id"
												if tNE.State <> 0 then tNE.Close
												tNE.open s, cn
												do while Not tNE.Eof
													if vCDAux(UBound(vCDAux)).c1 <> 0 then
														redim preserve vCDAux(UBound(vCDAux)+1)
														set vCDAux(UBound(vCDAux)) = new cl_TRES_COLUNAS
														end if
													vCDAux(UBound(vCDAux)).c1 = CLng(tNE("id"))
													vCDAux(UBound(vCDAux)).c2 = Trim("" & tNE("apelido"))
													'O NOVO CD SERÁ CONSIDERADO INICIALMENTE COMO 'DESATIVADO' POR PRECAUÇÃO, EXIGINDO QUE SEJA ALTERADO P/ 'ATIVADO' P/ QUE POSSA SER USADO
													vCDAux(UBound(vCDAux)).c3 = 0
													tNE.MoveNext
													loop
												end if
											end if

										for iCD=LBound(vCDAux) to UBound(vCDAux)
											'VERIFICA SE O CD ESTÁ HABILITADO OU NÃO NO CADASTRO BASE
											if vCDAux(iCD).c3 = 1 then
												s_checked = ""
											else
												s_checked = " checked"
												end if
											id_nfe_emitente_aux = vCDAux(iCD).c1
											s_cd_apelido = vCDAux(iCD).c2
											s_row_id = "row_" & vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & id_nfe_emitente_aux
											if operacao_selecionada <> OP_INCLUI then
												s = "SELECT * FROM t_WMS_REGRA_CD_X_UF_X_PESSOA_X_CD WHERE (id_wms_regra_cd_x_uf_x_pessoa = " & idRegraUfPessoa & ") AND (id_nfe_emitente = " & id_nfe_emitente_aux & ")"
												if tRegraUfPessoaCd.State <> 0 then tRegraUfPessoaCd.Close
												tRegraUfPessoaCd.open s, cn
												if Not tRegraUfPessoaCd.Eof then
													idRegraUfPessoaCd = tRegraUfPessoaCd("id")
													if CLng(tRegraUfPessoaCd("st_inativo")) = CLng(1) then s_checked = " checked"
													end if
												end if
									%>
									<tr id="<%=s_row_id%>" class="TableRow">
										<!-- Armazena posição/prioridade do CD na tabela -->
										<%
											s_id = "c_ordem_" & vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & id_nfe_emitente_aux
											idxOrdem = idxOrdem + 1
											%>
										<input type="hidden" name="<%=s_id%>" id="<%=s_id%>" class="CampoOrdenacao" value="<%=idxOrdem%>" />
										<!-- Armazena campo id_nfe_emitente -->
										<% s_id = "c_id_nfe_emitente_" & vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & id_nfe_emitente_aux %>
										<input type="hidden" name="<%=s_id%>" id="<%=s_id%>" class="CampoIdNfeEmitente" value="<%=id_nfe_emitente_aux%>" />
										<!-- CD -->
										<td class="TdSeqCd">
										<% s_id = "c_cd_" & vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & id_nfe_emitente_aux %>
											<input type="text" name="<%=s_id%>" id="<%=s_id%>" class="TA TxtCd TxtCd_<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>" readonly tabindex="-1" value="<%=s_cd_apelido%>" />
										</td>
										<!-- CHECKBOX: Ativado/Desativado -->
										<td class="TdSeqCkb">
											<% s_id = "ckb_cd_" & vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & id_nfe_emitente_aux %>
											<input type="checkbox" name="<%=s_id%>" id="<%=s_id%>" value="<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>_<%=id_nfe_emitente_aux%>" class="TA CkbCd CkbCd_<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>"<%=s_checked%> /><span id="spn_cd_<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>_<%=id_nfe_emitente_aux%>" class="C SpnCd_<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>" style="cursor:default;margin-right:4px;" onclick="fCAD.ckb_cd_<%=vUF(iUF)%>_<%=vPessoa(iPessoa)%>_<%=id_nfe_emitente_aux%>.click();">Desativado</span>
										</td>
										<td>
											<!-- A ÚLTIMA LINHA TAMBÉM POSSUI O BOTÃO DE MOVER P/ BAIXO DEVIDO À ROTINA QUE REALIZA SWAP DA LINHA DA TABELA, OU SEJA, TODAS AS LINHAS PRECISAM TER O MESMO PADRÃO -->
											<a name="bSetaBaixo" id="bSetaBaixo_<%=vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & id_nfe_emitente_aux%>" href="javascript:CDMoverParaBaixo('<%=s_table_id%>','<%=s_row_id%>')" title="move para baixo"
												tabindex=-1>
												<img src="../botao/SetaBaixo.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
										</td>
										<td>
											<!-- A PRIMEIRA LINHA TAMBÉM POSSUI O BOTÃO DE MOVER P/ CIMA DEVIDO À ROTINA QUE REALIZA SWAP DA LINHA DA TABELA, OU SEJA, TODAS AS LINHAS PRECISAM TER O MESMO PADRÃO -->
											<a name="bSetaCima" id="bSetaCima_<%=vUF(iUF) & "_" & vPessoa(iPessoa) & "_" & id_nfe_emitente_aux%>" href="javascript:CDMoverParaCima('<%=s_table_id%>','<%=s_row_id%>')" title="move para cima"
												tabindex=-1>
												<img src="../botao/SetaCima.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
										</td>
									</tr>
									<% next %>
								</table>
							</td>
							<td style="min-width:20px;">&nbsp;</td>
							<td valign="top">
								<table class="ME MB MD" cellspacing="0" cellpadding="0" width="100%">
									<tr>
										<td valign="top">
											<% s_id = "c_cd_spe_" & vUF(iUF) & "_" & vPessoa(iPessoa) %>
											<select id="<%=s_id%>" name="<%=s_id%>" style="margin:4px 8px 8px 8px;">
												<%=wms_apelido_empresa_nfe_emitente_monta_itens_select(spe_id_nfe_emitente) %>
											</select>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		
		<%if iPessoa < UBound(vPessoa) then %>
		<br />
		<% end if %>

		<% next %>
		</td>
	</tr>
</table>

<%
		next
%>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width='649' cellpadding='0' cellspacing='0' border='0' style="margin-top:5px;">
<tr>
	<td width="60%" align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkExpandirTudo" href="javascript:expandirTudo();"><p class="Button BtnAll" style="margin-bottom:0px;">Expandir Tudo</p></a></td>
	<td align="left" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkRecolherTudo" href="javascript:recolherTudo();"><p class="Button BtnAll" style="margin-bottom:0px;">Recolher Tudo</p></a></td>
</tr>
</table>

<br />

<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a href="javascript:history.back()" title="cancela as alterações da regra">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='CENTER'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveRegra(fCAD)' "
		s =s + "title='remove a regra cadastrada'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaRegra(fCAD)" title="atualiza o cadastro da regra">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	if tRegra.State <> 0 then tRegra.Close
	set tRegra = nothing
	
	if tRegraUf.State <> 0 then tRegraUf.Close
	set tRegraUf = nothing
	
	if tRegraUfPessoa.State <> 0 then tRegraUfPessoa.Close
	set tRegraUfPessoa = nothing
	
	if tRegraUfPessoaCd.State <> 0 then tRegraUfPessoaCd.Close
	set tRegraUfPessoaCd = nothing
	
	cn.Close
	set cn = nothing
%>