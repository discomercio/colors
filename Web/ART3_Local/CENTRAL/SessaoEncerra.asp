<%@ language=VBScript%>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'	REVISADO P/ IE10

	On Error GoTo 0
	Err.Clear

'	ATUALIZA BANCO DE DADOS
	if Trim(Session("usuario_atual")) <> "" then
		dim cn
		dim strSQL
		if bdd_conecta(cn) then
			strSQL = "UPDATE t_USUARIO SET" & _
						" SessionCtrlTicket = NULL," & _
						" SessionCtrlLoja = NULL," & _
						" SessionCtrlModulo = NULL," & _
						" SessionCtrlDtHrLogon = NULL," & _
						" SessionTokenModuloCentral = NULL," & _
						" DtHrSessionTokenModuloCentral = NULL" & _
					" WHERE" & _
						" usuario = '" & QuotedStr(Trim(Session("usuario_atual"))) & "'"
			cn.Execute(strSQL)
			
			strSQL = "UPDATE t_SESSAO_HISTORICO SET" & _
						" DtHrTermino = " & bd_formata_data_hora(Now) & _
					 " WHERE" & _
						" usuario = '" & QuotedStr(Trim("" & Session("usuario_atual"))) & "'" & _
						" AND DtHrInicio >= " & bd_formata_data_hora(Now-1) & _
						" AND SessionCtrlTicket = '" & Trim(Session("SessionCtrlTicket")) & "'"
			cn.Execute(strSQL)

        '   LIMPA EVENTUAIS LOCKS REMANESCENTES NOS RELATÓRIOS
            strSQL = "UPDATE tCRUP SET" & _
                    " locked = 0," & _
                    " cod_motivo_lock_released = " & CTRL_RELATORIO_CodMotivoLockReleased_SessaoEncerradaCentral & "," & _
                    " dt_hr_lock_released = getdate()" & _
                " FROM t_CTRL_RELATORIO_USUARIO_X_PEDIDO tCRUP INNER JOIN t_CTRL_RELATORIO tCR ON (tCRUP.id_relatorio = tCR.id)" & _
                " WHERE" & _
                    " (tCR.modulo = 'CENTRAL')" & _
                    " AND (tCRUP.usuario = '" & QuotedStr(Trim(Session("usuario_atual"))) & "')" & _
                    " AND (locked = 1)"
            cn.Execute(strSQL)

			cn.Close
			end if
		set cn = nothing
		end if
		
'	ENCERRA A SESSÃO
	Session("usuario_atual") = " "
	Session("senha_atual") = " "
	Session.Abandon

%>


<html>

<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script language="JavaScript" type="text/javascript">
	window.close();
</script>

</html>
