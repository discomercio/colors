<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =================================================================
'	  A J A X P E D I D O D E V O L U C A O R E M O V E F O T O . A S P
'     =================================================================
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
	dim strResp, blnErro, msg_erro
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s, i, cont, id_devolucao, id_upload_file, usuario
    
	id_devolucao = Trim(Request("id_devolucao"))
	id_upload_file = Trim(Request("id_upload_file"))
    usuario = Trim(Request("usuario"))

    dim serverVariablesUrl, serverVariablesServerName
    dim x, full_url_file_src, full_url_file_href, file_attr_title, file_extension
	'Esta página foi ajustada para usar a função getProtocoloEmUsoHttpOrHttps() na montagem da URL
    serverVariablesServerName = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
    serverVariablesUrl = Request.ServerVariables("URL")
    serverVariablesUrl = Ucase(serverVariablesUrl)
    serverVariablesUrl = Mid(serverVariablesUrl, 1, CInt(InStr(serverVariablesUrl, "GLOBAL")-1))
    serverVariablesUrl = serverVariablesServerName & serverVariablesUrl
	
'	SETA O CAMPO st_delete_file PARA COLOCAR O ARQUIVO NA FILA PARA SER REMOVIDO NA PRÓXIMA ROTINA AUTOMÁTICA DO SERVIDOR
	s = "SELECT * FROM t_UPLOAD_FILE WHERE (id = '" & id_upload_file & "')"

'	EXECUTA A CONSULTA
	strResp = ""
    blnErro = False
	
    cn.BeginTrans

	if rs.State <> 0 then rs.Close
	rs.open s, cn

	if Not rs.Eof then
		rs("st_delete_file") = 1
        rs("usuario_delete_file") = usuario
        rs("dt_delete_file") = Date
        rs("dt_hr_delete_file") = Now
        rs("dt_delete_file_scheduled_date") = Date
        rs.Update
        if Err <> 0 then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
            blnErro = True
		    end if
        end if
		
	if Not blnErro then
        s = "DELETE FROM t_PEDIDO_DEVOLUCAO_IMAGEM WHERE (id_pedido_devolucao = '" & id_devolucao & "' AND id_upload_file = '" & id_upload_file & "')"
        cn.Execute(s)
        if Err <> 0 then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
            blnErro = True
		    end if 
        end if

    if Not blnErro then
        '	~~~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~~~
        s = "SELECT * FROM t_UPLOAD_FILE" & _
            " INNER JOIN t_PEDIDO_DEVOLUCAO_IMAGEM ON (t_UPLOAD_FILE.id=t_PEDIDO_DEVOLUCAO_IMAGEM.id_upload_file)" & _
            " WHERE (" & _
                "id_pedido_devolucao = '" & id_devolucao & "'" & _
                " AND st_file_deleted = 0" & _
            ") ORDER BY dt_hr_cadastro"
        
        if rs.State <> 0 then rs.Close
	    rs.open s, cn
        i = 0
        strResp = strResp & "<tr>"
        do while Not rs.Eof
            i = i + 1
            x = rs("stored_file_name")
            file_extension = Mid(x, Instr(x, ".")+1, Len(x))
            full_url_file_href = getProtocoloEmUsoHttpOrHttps & "://"
            full_url_file_href = full_url_file_href & serverVariablesUrl
            full_url_file_href = full_url_file_href & "FileServer/"
            full_url_file_href = full_url_file_href & rs("stored_relative_path")
            full_url_file_href = full_url_file_href & "/" & rs("stored_file_name")
            select case file_extension
                case "jpeg", "jpg", "png", "bmp", "gif", "tif", "tiff"
                    full_url_file_src = full_url_file_href
                    file_attr_title = "clique para visualizar a imagem no tamanho original"
                case "pdf"
                    full_url_file_src = "../IMAGEM/file_pdf_150x150.png"
                    file_attr_title = "clique para visualizar o conteúdo do PDF"
                case else
                    full_url_file_src = "../IMAGEM/file_150x150.png"
                    file_attr_title = "clique para visualizar o conteúdo do arquivo"
            end select
            if i = 4 then
                strResp = strResp & "</tr>" & chr(13) & _
                            "<tr>" & chr(13)
                end if

            strResp = strResp & "<td style='width: 220px;' valign='top'>" & _
                            "   <a href='" & full_url_file_href & "' target='_blank' title='" & file_attr_title & "'>" & chr(13) & _
                            "   <img src='" & full_url_file_src & "' style='width: 150px; height: 150px;border:1px dashed black;' /></a>" & chr(13) & _
                            "   <a href='javascript:fPEDRemoverFoto(" & chr(34) & CStr(rs("id_upload_file")) & chr(34) & ")' title='remover arquivo'>" & chr(13) & _
                            "   <img src='../BOTAO/botao_X_red.gif' style='margin-left: 0px;vertical-align: top' /></a>" & chr(13) & _
                        "</td>" & chr(13)
        
            rs.MoveNext
            loop
        if i < PEDIDO_DEVOLUCAO_QTDE_FOTO then
            for cont = i to PEDIDO_DEVOLUCAO_QTDE_FOTO -1
                i = i + 1
                if i = 4 then
                strResp = strResp & "</tr>" & chr(13) & _
                            "<tr>" & chr(13)
                end if
                strResp = strResp & _
                    "<td>" & chr(13) & _
                    "   <input type='file' name='arquivo" & Cstr(i) & "' id='arquivo" & Cstr(i) & "' onchange='reloadFileReader(" & chr(34) & CStr(i) & chr(34) & ")' class='PLLd' style='font-weight: normal; width: 180px' /><br />" & chr(13) & _
                    "   <div id='image-holder-arquivo" & Cstr(i) & "' style='width: 180px; height: 140px; border: 1px dashed #000; margin-top: 3px' onclick='fPED.arquivo" & Cstr(i) & ".click();'></div>" & chr(13) & _
                    "</td>" & chr(13)
                next
            end if

        strResp = strResp & "</tr>"
        end if
	
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
