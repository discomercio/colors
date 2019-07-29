#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
#endregion

namespace Financeiro
{
	class BoletoHtml
	{
		#region [ Getter ]
		private String _textoBoletoHtml;
		public String textoBoletoHtml
		{
			get { return _textoBoletoHtml; }
		}
		#endregion

		#region [ Construtor ]
		public BoletoHtml(	String dataVencimento,
							String valorDocumento,
							String dataProcessamento,
							String dataDocumento,
							String numeroDocumento,
							String nomeCedente,
							String carteira,
							String agenciaECodigoCedente,
							String nossoNumero,
							String nomeSacado,
                            String numInscricaoSacado,
							String enderecoSacado,
							String cepCidadeUfSacado,
							String linhaDigitavel,
							String codigoBarras,
							String reciboSacadoInstrucoesLinha1,
							String reciboSacadoInstrucoesLinha2,
							String reciboSacadoInstrucoesLinha3,
							String reciboSacadoInstrucoesLinha4,
							String reciboSacadoInstrucoesLinha5,
							String reciboSacadoInstrucoesLinha6,
							String fichaCompensacaoInstrucoesLinha1,
							String fichaCompensacaoInstrucoesLinha2,
							String fichaCompensacaoInstrucoesLinha3,
							String fichaCompensacaoInstrucoesLinha4,
							String fichaCompensacaoInstrucoesLinha5,
							String fichaCompensacaoInstrucoesLinha6)
		{
			_textoBoletoHtml = _textoBaseBoletoHtml;
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|DATA_VENCIMENTO|]", dataVencimento);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|VALOR_DOCUMENTO|]", valorDocumento.PadLeft(17, ' '));
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|DATA_PROCESSAMENTO|]", dataProcessamento);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|DATA_DOCUMENTO|]", dataDocumento);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|NUMERO_DOCUMENTO|]", numeroDocumento);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|NOME_CEDENTE|]", nomeCedente);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|CARTEIRA|]", carteira);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|AGENCIA_E_CODIGO_CEDENTE|]", agenciaECodigoCedente);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|NOSSO_NUMERO|]", nossoNumero);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|NOME_SACADO|]", nomeSacado);
            _textoBoletoHtml = _textoBoletoHtml.Replace("[|NUMERO_INSCRICAO_SACADO|]", Global.formataCnpjCpf(numInscricaoSacado));
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|ENDERECO_SACADO|]", enderecoSacado);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|CEP_CIDADE_UF_SACADO|]", cepCidadeUfSacado);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|LINHA_DIGITAVEL|]", linhaDigitavel);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|CODIGO_BARRAS|]", codificaCodigoBarrasIwts(codigoBarras));
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|RECIBO_SACADO_INSTRUCOES_LINHA_1|]", reciboSacadoInstrucoesLinha1);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|RECIBO_SACADO_INSTRUCOES_LINHA_2|]", reciboSacadoInstrucoesLinha2);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|RECIBO_SACADO_INSTRUCOES_LINHA_3|]", reciboSacadoInstrucoesLinha3);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|RECIBO_SACADO_INSTRUCOES_LINHA_4|]", reciboSacadoInstrucoesLinha4);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|RECIBO_SACADO_INSTRUCOES_LINHA_5|]", reciboSacadoInstrucoesLinha5);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|RECIBO_SACADO_INSTRUCOES_LINHA_6|]", reciboSacadoInstrucoesLinha6);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_1|]", fichaCompensacaoInstrucoesLinha1);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_2|]", fichaCompensacaoInstrucoesLinha2);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_3|]", fichaCompensacaoInstrucoesLinha3);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_4|]", fichaCompensacaoInstrucoesLinha4);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_5|]", fichaCompensacaoInstrucoesLinha5);
			_textoBoletoHtml = _textoBoletoHtml.Replace("[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_6|]", fichaCompensacaoInstrucoesLinha6);

			if (File.Exists(Global.Cte.Imagens.PathImagens + "\\" + Global.Cte.Imagens.ArqLogoBradesco))
			{
				_textoBoletoHtml = _textoBoletoHtml.Replace("http://central85.com.br/images/drd_pb_g.gif", "file://" + Global.Cte.Imagens.PathImagens.Replace("\\", "/") + "/" + Global.Cte.Imagens.ArqLogoBradesco);
			}
		}
		#endregion

		#region [ Métodos ]

		#region [ codificaCodigoBarrasIwts ]
		public static String codificaCodigoBarrasIwts(String codigoBarras)
		{
			#region [ Declarações ]
			String strResposta;
			#endregion

			#region [ Consistência ]
			if (codigoBarras == null) return "";
			if (codigoBarras.Trim().Length == 0) return "";
			#endregion

			strResposta = Global.digitos(codigoBarras);
			strResposta = strResposta.Replace('0', 'n');
			strResposta = strResposta.Replace('1', 'O');
			strResposta = strResposta.Replace('2', 's');
			strResposta = strResposta.Replace('3', 'E');
			strResposta = strResposta.Replace('4', 'c');
			strResposta = strResposta.Replace('5', 'Q');
			strResposta = strResposta.Replace('6', 't');
			strResposta = strResposta.Replace('7', 'D');
			strResposta = strResposta.Replace('8', 'u');
			strResposta = strResposta.Replace('9', 'Z');

			return strResposta;
		}
		#endregion

		#endregion

		#region [ textoBaseBoletoHtml ]
		private const String _textoBaseBoletoHtml = @"
<html xmlns:v=""urn:schemas-microsoft-com:vml""
xmlns:o=""urn:schemas-microsoft-com:office:office""
xmlns:w=""urn:schemas-microsoft-com:office:word""
xmlns=""http://www.w3.org/TR/REC-html40"">
<head>
<script LANGUAGE=""JavaScript"">
<!--
   function click()
   { if (event.button==2||event.button==3) { alert('Não é possível editar este boleto!'); } }
   document.onmousedown=click
// -->
</script>
<meta HTTP-EQUIV=""Expires"" CONTENT=""Thu, 01 Jan 1970 00:00:00 GMT"">
<meta HTTP-EQUIV=""Cache-Control"" content=""no-store"">
<meta HTTP-EQUIV=""Pragma"" content=""no-cache"">
<meta http-equiv=Content-Type content=""text/html; charset=windows-1252"">
<meta name=ProgId content=Word.Document>
<meta name=Generator content=""Microsoft Word 9"">
<meta name=Originator content=""Microsoft Word 9"">
<link rel=File-List href=""./teste_arquivos/filelist.xml"">
<link rel=Edit-Time-Data href=""./teste_arquivos/editdata.mso"">
<!--[if !mso]>
<style>
v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style>
<![endif]-->
<title>Boleto Bradesco</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author></o:Author>
  <o:Template>Normal</o:Template>
  <o:LastAuthor></o:LastAuthor>
  <o:Revision>14</o:Revision>
  <o:TotalTime>377</o:TotalTime>
  <o:LastPrinted>2001-01-23T18:01:00Z</o:LastPrinted>
  <o:Created>2001-02-20T17:36:00Z</o:Created>
  <o:LastSaved>2001-03-06T18:01:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Words>389</o:Words>
  <o:Characters>2219</o:Characters>
  <o:Company></o:Company>
  <o:Lines>18</o:Lines>
  <o:Paragraphs>4</o:Paragraphs>
  <o:CharactersWithSpaces>2725</o:CharactersWithSpaces>
  <o:Version>9.2812</o:Version>
 </o:DocumentProperties>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:HyphenationZone>21</w:HyphenationZone>
  <w:DrawingGridHorizontalSpacing>9,35 pt</w:DrawingGridHorizontalSpacing>
  <w:DisplayVerticalDrawingGridEvery>2</w:DisplayVerticalDrawingGridEvery>
 </w:WordDocument>
</xml><![endif]-->
<style>
<!--
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"""";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:""Times New Roman"";
	mso-fareast-font-family:""Times New Roman"";}
h1
	{mso-style-next:Normal;
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:8.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Arial;
	mso-font-kerning:0pt;}
@page Section1
	{size:612.0pt 792.0pt;
	margin:70.9pt 3.0cm 70.9pt 3.0cm;
	mso-header-margin:35.45pt;
	mso-footer-margin:35.45pt;
	mso-vertical-page-align:bottom;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
<!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext=""edit"" spidmax=""1100""/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext=""edit"">
  <o:idmap v:ext=""edit"" data=""1""/>
 </o:shapelayout></xml><![endif]-->
</head>
<body lang=PT-BR style='tab-interval:35.4pt'>
<div class=Section1>
<table border=1 cellspacing=0 cellpadding=0 width=706 style='width:529.35pt;
 margin-left:-3.5pt;border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;
 mso-padding-alt:0cm 3.5pt 0cm 3.5pt'>
 <tr style='height:25.3pt'>
  <td width=138 colspan=7 valign=top style='width:103.35pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:25.3pt'>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><!--[if gte vml 1]><v:shapetype id=""_x0000_t75""
   coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe""
   filled=""f"" stroked=""f"">
   <v:stroke joinstyle=""miter""/>
   <v:formulas>
    <v:f eqn=""if lineDrawn pixelLineWidth 0""/>
    <v:f eqn=""sum @0 1 0""/>
    <v:f eqn=""sum 0 0 @1""/>
    <v:f eqn=""prod @2 1 2""/>
    <v:f eqn=""prod @3 21600 pixelWidth""/>
    <v:f eqn=""prod @3 21600 pixelHeight""/>
    <v:f eqn=""sum @0 0 1""/>
    <v:f eqn=""prod @6 1 2""/>
    <v:f eqn=""prod @7 21600 pixelWidth""/>
    <v:f eqn=""sum @8 21600 0""/>
    <v:f eqn=""prod @7 21600 pixelHeight""/>
    <v:f eqn=""sum @10 21600 0""/>
   </v:formulas>
   <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>
   <o:lock v:ext=""edit"" aspectratio=""t""/>
  </v:shapetype><v:shape id=""_x0000_i1025"" type=""#_x0000_t75"" style='width:92.25pt;
   height:24.75pt'>
   <v:imagedata src=""http://central85.com.br/images/drd_pb_g.gif"" width=""123"" height=""33"" o:title=""Bradesco""/>
  </v:shape><![endif]--><![if !vml]><img width=123 height=33
  src=""http://central85.com.br/images/drd_pb_g.gif"" width=""123"" height=""33""><br><![endif]><o:p></o:p></span></p>
  </td>
  <td width=87 colspan=3 style='width:65.4pt;border-top:none;border-left:none;
  border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:25.3pt'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='font-size:20.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>237-2<o:p></o:p></span></b></p>
  </td>
  <td width=265 colspan=7 valign=bottom style='width:198.4pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:25.3pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=216 colspan=5 valign=bottom style='width:162.2pt;border:none;
  border-bottom:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:25.3pt'>
  <h1>Recibo do Sacado</h1>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:25.3pt;border:none' width=0 height=34></td>
  <![endif]>
 </tr>
 <tr style='height:3.5pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:3.5pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Local de pagamento</span><b><span style='font-size:8.0pt;
  mso-bidi-font-size:12.0pt;font-family:""Courier New""'><o:p></o:p></span></b></p>
  </td>
  <td width=216 colspan=5 rowspan=7 style='width:162.2pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:3.5pt'>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><!--[if gte vml 1]><v:shape id=""_x0000_i1095"" type=""#_x0000_t75""
   style='width:92.25pt;height:24.75pt'>
   <v:imagedata src=""http://central85.com.br/images/drd_pb_g.gif"" width=""123"" height=""33"" o:title=""Bradesco""/>
  </v:shape><![endif]--><![if !vml]><img width=123 height=33
  src=""http://central85.com.br/images/drd_pb_g.gif"" width=""123"" height=""33""><![endif]><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:3.5pt;border:none' width=0 height=5></td>
  <![endif]>
 </tr>
 <tr style='height:15.75pt'>
  <td width=48 colspan=2 valign=top style='width:36.35pt;border:none;
  border-bottom:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:15.75pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><b><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New""'><o:p></o:p></span></b></p>
  </td>
  <td width=441 colspan=15 valign=top style='width:330.8pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:15.75pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>Banco Bradesco S.A.<o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>Pagável preferencialmente nas Agências Bradesco</span></b><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:15.75pt;border:none' width=0 height=21></td>
  <![endif]>
 </tr>
 <tr style='height:18.15pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:18.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Cedente<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  </span><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>[|NOME_CEDENTE|]<o:p></o:p></span></b></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:18.15pt;border:none' width=0 height=24></td>
  <![endif]>
 </tr>
 <tr style='height:18.15pt'>
  <td width=125 colspan=6 valign=top style='width:93.95pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:18.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Data do Documento<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>[|DATA_DOCUMENTO|]<o:p></o:p></span></p>
  </td>
  <td width=124 colspan=6 valign=top style='width:93.0pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:18.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Número do Documento<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  </span><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>[|NUMERO_DOCUMENTO|]<o:p></o:p></span></b></p>
  </td>
  <td width=90 colspan=2 valign=top style='width:67.15pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:18.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Espécie Documento<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
  <h1><span style='font-family:""Courier New"";font-weight:normal'>DM<o:p></o:p></span></h1>
  </td>
  <td width=63 colspan=2 valign=top style='width:47.05pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:18.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Aceite<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'> <o:p></o:p></span></p>
  </td>
  <td width=88 valign=top style='width:66.0pt;border-top:none;border-left:none;
  border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:18.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Data Processamento<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>[|DATA_PROCESSAMENTO|]<o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:18.15pt;border:none' width=0 height=24></td>
  <![endif]>
 </tr>
 <tr style='height:20.7pt'>
  <td width=66 colspan=3 valign=top style='width:49.4pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:20.7pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Uso do Banco<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  <br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>8600<o:p></o:p></span></p>
  </td>
  <td width=59 colspan=3 valign=top style='width:44.55pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:20.7pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Cip<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  <br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>000</span><span style='font-size:5.0pt;mso-bidi-font-size:
  12.0pt;font-family:""Courier New""'><o:p></o:p></span></p>
  </td>
  <td width=50 colspan=3 valign=top style='width:37.45pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:20.7pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Carteira<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  <br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>[|CARTEIRA|]<o:p></o:p></span></p>
  </td>
  <td width=74 colspan=3 valign=top style='width:55.55pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:20.7pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Espécie Moeda<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  <br>
  </span><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>REAL<o:p></o:p></span></b></p>
  </td>
  <td width=152 colspan=4 valign=top style='width:114.2pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:20.7pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Quantidade<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  <br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'><o:p></o:p></span></p>
  </td>
  <td width=88 valign=top style='width:66.0pt;border-top:none;border-left:none;
  border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:20.7pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Valor<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:20.7pt;border:none' width=0 height=28></td>
  <![endif]>
 </tr>
 <tr style='height:6.75pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.75pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Instruções: (Texto de responsabilidade do Cedente)</span><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:6.75pt;border:none' width=0 height=9></td>
  <![endif]>
 </tr>
 <tr style='height:14.95pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:14.95pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|RECIBO_SACADO_INSTRUCOES_LINHA_1|]<o:p></o:p></span></b></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:14.95pt;border:none' width=0 height=20></td>
  <![endif]>
 </tr>
 <tr style='height:14.2pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:14.2pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|RECIBO_SACADO_INSTRUCOES_LINHA_2|]<o:p></o:p></span></b></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;background:#CCCCCC;padding:0cm 3.5pt 0cm 3.5pt;
  height:14.2pt'>
  <p class=MsoNormal><b><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Vencimento<o:p></o:p></span></b></p>
  <p class=MsoNormal align=right style='text-align:right'><b><span
  style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>       [|DATA_VENCIMENTO|]<o:p></o:p></span></b></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border-top:solid windowtext .5pt;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:none;
  background:#CCCCCC;padding:0cm 3.5pt 0cm 3.5pt;height:14.2pt'>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:14.2pt;border:none' width=0 height=19></td>
  <![endif]>
 </tr>
 <tr style='height:15.95pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:15.95pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|RECIBO_SACADO_INSTRUCOES_LINHA_3|]<o:p></o:p></span></b></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:15.95pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Agência / Cód. Cedente<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>[|AGENCIA_E_CODIGO_CEDENTE|]<o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:15.95pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:15.95pt;border:none' width=0 height=21></td>
  <![endif]>
 </tr>
 <tr style='height:8.8pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:8.8pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|RECIBO_SACADO_INSTRUCOES_LINHA_4|]<o:p></o:p></span></b></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:8.8pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Nosso Número<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>[|NOSSO_NUMERO|]<o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:8.8pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:8.8pt;border:none' width=0 height=12></td>
  <![endif]>
 </tr>
 <tr style='height:8.6pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:8.6pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|RECIBO_SACADO_INSTRUCOES_LINHA_5|]<o:p></o:p></span></b></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;background:#CCCCCC;padding:0cm 3.5pt 0cm 3.5pt;
  height:8.6pt'>
  <p class=MsoNormal><b><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>1 ( = ) Valor Documento<o:p></o:p></span></b></p>
  <p class=MsoNormal align=right style='text-align:right'><b><span
  style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>[|VALOR_DOCUMENTO|]<o:p></o:p></span></b></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;background:
  #CCCCCC;padding:0cm 3.5pt 0cm 3.5pt;height:8.6pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:8.6pt;border:none' width=0 height=11></td>
  <![endif]>
 </tr>
 <tr style='height:14.05pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:14.05pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|RECIBO_SACADO_INSTRUCOES_LINHA_6|]<o:p></o:p></span></b></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:14.05pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>2 ( - ) Desconto / Abatimento<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:14.05pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:14.05pt;border:none' width=0 height=19></td>
  <![endif]>
 </tr>
 <tr style='height:7.55pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:7.55pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'> <o:p></o:p></span></b></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:7.55pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>3 ( - ) Outras Deduções<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:7.55pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:7.55pt;border:none' width=0 height=10></td>
  <![endif]>
 </tr>
 <tr style='height:10.2pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:10.2pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'> <o:p></o:p></span></b></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:10.2pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>4 ( + ) Moras / Multa<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:10.2pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:10.2pt;border:none' width=0 height=14></td>
  <![endif]>
 </tr>
 <tr style='height:3.25pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:3.25pt'>
  <p class=MsoNormal><sup><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'></span></sup><b><span style='font-size:8.0pt;
  mso-bidi-font-size:12.0pt;font-family:""Courier New""'><o:p></o:p></span></b></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:3.25pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>5 ( + ) Outros Acrécimos<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:3.25pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:3.25pt;border:none' width=0 height=4></td>
  <![endif]>
 </tr>
 <tr style='height:6.4pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.4pt'>
  <p class=MsoNormal><sup><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>.</span></sup><span style='font-size:5.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'> Recebimento através do cheque nº<span
  style=""mso-spacerun: yes"">                                                  
  </span>do Banco<span style=""mso-spacerun:
  yes"">                                      </span>.<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Quitação válida somente após liquidação do cheque.</span><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.4pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>6 ( = ) Valor Cobrado</span><span style='font-size:8.0pt;
  mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.4pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:6.4pt;border:none' width=0 height=9></td>
  <![endif]>
 </tr>
 <tr style='height:29.85pt'>
  <td width=83 colspan=2 valign=top style='width:62.45pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:29.85pt'>
  <p class=MsoNormal><span style='font-size:6.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Sacado:</span><span style='font-size:8.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=623 colspan=20 valign=top style='width:466.9pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:29.85pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>[|NOME_SACADO|]&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[|NUMERO_INSCRICAO_SACADO|]<o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>[|ENDERECO_SACADO|]<o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>[|CEP_CIDADE_UF_SACADO|]</span></b><span style='font-size:8.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:29.85pt;border:none' width=0 height=40></td>
  <![endif]>
 </tr>
 <tr style='height:9.4pt'>
  <td width=618 colspan=19 valign=bottom style='width:463.5pt;border:none;
  border-bottom:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:9.4pt'>
  <p class=MsoNormal><span style='font-size:6.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Sacador / Avalista: <b></b><o:p></o:p></span></p>
  </td>
  <td width=88 colspan=3 valign=bottom style='width:65.85pt;border:none;
  border-bottom:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:9.4pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:9.4pt;border:none' width=0 height=13></td>
  <![endif]>
 </tr>
 <tr style='height:14.85pt'>
  <td width=654 colspan=21 valign=top style='width:490.25pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:14.85pt'>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:4.5pt;mso-bidi-font-size:12.0pt;font-family:Arial'>Autenticação
  Mecânica</span><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=52 valign=top style='width:39.1pt;border:none;mso-border-top-alt:
  solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:14.85pt'>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:14.85pt;border:none' width=0 height=20></td>
  <![endif]>
 </tr>
 <tr style='height:8.3pt'>
  <td width=706 colspan=22 valign=top style='width:529.35pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:8.3pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>.................................................................................................................................................................................................................................................................................................................................................<o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:8.3pt;border:none' width=0 height=11></td>
  <![endif]>
 </tr>
 <tr style='height:20.55pt'>
  <td width=138 colspan=7 valign=top style='width:103.35pt;border:none;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:20.55pt'>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><!--[if gte vml 1]><v:shape id=""_x0000_i1096"" type=""#_x0000_t75""
   style='width:92.25pt;height:24.75pt'>
   <v:imagedata src=""http://central85.com.br/images/drd_pb_g.gif"" width=""123"" height=""33"" o:title=""Bradesco""/>
  </v:shape><![endif]--><![if !vml]>
  <img src=""http://central85.com.br/images/drd_pb_g.gif"" width=""123"" height=""33""><br>
  <![endif]><o:p></o:p></span></p>
  </td>
  <td width=87 colspan=3 style='width:65.4pt;border:none;border-right:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:20.55pt'>
  <p class=MsoNormal align=center style='text-align:center'><b><span
  style='font-size:20.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>237-2</span></b><b><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></b></p>
  </td>
  <td width=415 colspan=10 style='width:311.3pt;border:none;mso-border-left-alt:
  solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:20.55pt'>
  <h1 align=center style='text-align:center'><span style='font-size:10.0pt;
  mso-bidi-font-size:12.0pt'>[|LINHA_DIGITAVEL|]<o:p></o:p></span></h1>
  </td>
  <td width=66 colspan=2 valign=bottom style='width:49.3pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:20.55pt'>
  <h1 align=center style='text-align:center'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-weight:normal'><o:p></o:p></span></h1>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:20.55pt;border:none' width=0 height=27></td>
  <![endif]>
 </tr>
 <tr style='height:3.7pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border-top:solid windowtext .5pt;
  border-left:none;border-bottom:none;border-right:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:3.7pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Local de Pagamento<o:p></o:p></span></p>
  </td>
  <td width=151 colspan=3 rowspan=2 valign=top style='width:112.9pt;border-top:
  solid windowtext .5pt;border-left:none;border-bottom:solid windowtext .5pt;
  border-right:none;mso-border-left-alt:solid windowtext .5pt;background:#CCCCCC;
  padding:0cm 3.5pt 0cm 3.5pt;height:3.7pt'>
  <p class=MsoNormal><b><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Vencimento<o:p></o:p></span></b></p>
  <p class=MsoNormal align=right style='text-align:right'><b><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
  <p class=MsoNormal align=right style='text-align:right'><b><span
  style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>       [|DATA_VENCIMENTO|]<o:p></o:p></span></b></p>
  </td>
  <td width=66 colspan=2 rowspan=2 valign=top style='width:49.3pt;border-top:
  solid windowtext .5pt;border-left:none;border-bottom:solid windowtext .5pt;
  border-right:none;background:#CCCCCC;padding:0cm 3.5pt 0cm 3.5pt;height:3.7pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:3.7pt;border:none' width=0 height=5></td>
  <![endif]>
 </tr>
 <tr style='height:18.35pt'>
  <td width=48 colspan=2 valign=top style='width:36.35pt;border:none;
  border-bottom:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:18.35pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=441 colspan=15 valign=top style='width:330.8pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:18.35pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>Banco Bradesco S.A.<o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>Pagável Preferencialmente nas Agências Bradesco.</span></b><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:18.35pt;border:none' width=0 height=24></td>
  <![endif]>
 </tr>
 <tr style='height:5.15pt'>
  <td width=490 colspan=17 valign=top style='width:367.15pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Cedente<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|NOME_CEDENTE|]</span></b><span style='font-size:5.0pt;
  mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Agência / Cód. Cedente<br style='mso-special-character:
  line-break'>
  <![if !supportLineBreakNewLine]><br style='mso-special-character:line-break'>
  <![endif]><o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>[|AGENCIA_E_CODIGO_CEDENTE|]</span><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:5.15pt;border:none' width=0 height=7></td>
  <![endif]>
 </tr>
 <tr style='height:5.15pt'>
  <td width=117 colspan=5 valign=top style='width:87.65pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Data do Documento<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|DATA_DOCUMENTO|]<o:p></o:p></span></p>
  </td>
  <td width=122 colspan=6 valign=top style='width:91.2pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Número do Documento<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial;color:white'>.</span><span style='font-size:8.0pt;
  mso-bidi-font-size:12.0pt;font-family:Arial;color:white'><o:p></o:p></span></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|NUMERO_DOCUMENTO|]</span></b><span style='font-size:5.0pt;
  mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=88 colspan=2 valign=top style='width:65.85pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Espécie Documento<br style='mso-special-character:line-break'>
  <![if !supportLineBreakNewLine]><br style='mso-special-character:line-break'>
  <![endif]><o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>DM</span><span style='font-size:5.0pt;
  mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=63 colspan=2 valign=top style='width:47.05pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Aceite<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><span style=""mso-spacerun: yes""> </span></span><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New""'> </span><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=101 colspan=2 valign=top style='width:75.4pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Data Processamento<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial;color:white'>.<o:p></o:p></span></p>
  <p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|DATA_PROCESSAMENTO|]</span><span style='font-size:5.0pt;
  mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Nosso Número<br style='mso-special-character:line-break'>
  <![if !supportLineBreakNewLine]><br style='mso-special-character:line-break'>
  <![endif]><o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>[|NOSSO_NUMERO|]</span><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:5.15pt;border:none' width=0 height=7></td>
  <![endif]>
 </tr>
 <tr style='height:16.9pt'>
  <td width=66 colspan=3 valign=top style='width:49.4pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:16.9pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Uso do Banco<br>
  <br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>8600</span><span style='font-size:5.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=51 colspan=2 valign=top style='width:38.25pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:16.9pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Cip<br>
  <br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>000</span><span style='font-size:5.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=50 colspan=3 valign=top style='width:37.4pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:16.9pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Carteira<br>
  <br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>[|CARTEIRA|]</span><span style='font-size:5.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=72 colspan=3 valign=top style='width:53.8pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:16.9pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Espécie Moeda<br>
  <br>
  </span><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'>REAL</span></b><span style='font-size:5.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=151 colspan=4 valign=top style='width:112.9pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:16.9pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Quantidade<br>
  <br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'></span><span style='font-size:5.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=101 colspan=2 valign=top style='width:75.4pt;border-top:none;
  border-left:none;border-bottom:solid windowtext .5pt;border-right:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:16.9pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Valor<br>
  <br>
  </span><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:
  ""Courier New""'></span><span style='font-size:5.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;background:#CCCCCC;padding:0cm 3.5pt 0cm 3.5pt;
  height:16.9pt'>
  <p class=MsoNormal><b><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>1 ( = ) Valor Documento<br>
  </span></b><b><span style='font-size:3.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial;color:silver'>.</span></b><b><span style='font-size:5.0pt;
  mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></b></p>
  <p class=MsoNormal align=right style='text-align:right'><b><span
  style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'>[|VALOR_DOCUMENTO|]</span></b><span
  style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;background:
  #CCCCCC;padding:0cm 3.5pt 0cm 3.5pt;height:16.9pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:16.9pt;border:none' width=0 height=23></td>
  <![endif]>
 </tr>
 <tr style='height:8.15pt'>
  <td width=15 rowspan=7 style='width:11.35pt;border:none;border-bottom:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:8.15pt'>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>I<o:p></o:p></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>N<o:p></o:p></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>S<o:p></o:p></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>T<o:p></o:p></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>R<o:p></o:p></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>U<o:p></o:p></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>Ç<o:p></o:p></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>Õ<o:p></o:p></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>E<o:p></o:p></span></p>
  <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;mso-bidi-font-size:12.0pt;font-family:""Courier New"";
  color:white'>S</span><span style='font-size:7.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial;color:white'><o:p></o:p></span></p>
  </td>
  <td width=474 colspan=16 valign=top style='width:355.8pt;border:none;
  border-right:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:8.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Instruções: (Texto de responsabilidade do Cedente)<o:p></o:p></span></p>
  </td>
  <td width=151 colspan=3 rowspan=2 valign=top style='width:112.9pt;border:
  none;border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:8.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>2 ( - ) Desconto / Abatimento<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><span
  style=""mso-spacerun: yes""> </span></span><span style='font-size:5.0pt;
  mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 rowspan=2 valign=top style='width:49.3pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:8.15pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:8.15pt;border:none' width=0 height=11></td>
  <![endif]>
 </tr>
 <tr style='height:8.15pt'>
  <td width=474 colspan=16 rowspan=6 valign=bottom style='width:355.8pt;
  border-top:none;border-left:none;border-bottom:solid windowtext .5pt;
  border-right:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:8.15pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>*** VALORES EXPRESSOS EM REAIS *** <span style='color:white'>.</span><o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_1|]<span style='color:white'>.<o:p></o:p></span></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_2|]<span style='color:white'>.</span><o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_3|]<span style='color:white'>.</span><o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_4|]<span style='color:white'>.</span><o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_5|]<span style='color:white'>.</span><o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:""Courier New""'>[|FICHA_COMPENSACAO_INSTRUCOES_LINHA_6|]<span style='color:white'>.</span><o:p></o:p></span></b></p>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>&nbsp;
  <o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:8.15pt;border:none' width=0 height=11></td>
  <![endif]>
 </tr>
 <tr style='height:8.15pt'>
  <td width=151 colspan=3 rowspan=2 valign=top style='width:112.9pt;border:
  none;border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:8.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>3 ( - ) Outras Deduções<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><span
  style=""mso-spacerun: yes""> </span></span><span style='font-size:5.0pt;
  mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 rowspan=2 valign=top style='width:49.3pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:8.15pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:8.15pt;border:none' width=0 height=11></td>
  <![endif]>
 </tr>
 <tr style='height:8.15pt'>
  <![if !supportMisalignedRows]>
  <td style='height:8.15pt;border:none' width=0 height=11></td>
  <![endif]>
 </tr>
 <tr style='height:5.15pt'>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>4 ( + ) Moras / Multa<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial;
  color:white'>.</span><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial;color:white'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:5.15pt;border:none' width=0 height=7></td>
  <![endif]>
 </tr>
 <tr style='height:5.15pt'>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>5 ( + ) Outros Acrécimos<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:8.0pt;mso-bidi-font-size:12.0pt;font-family:Arial;
  color:white'>.</span><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial;color:white'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:5.15pt;border:none' width=0 height=7></td>
  <![endif]>
 </tr>
 <tr style='height:5.15pt'>
  <td width=151 colspan=3 valign=top style='width:112.9pt;border:none;
  border-bottom:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>6 ( = ) Valor Cobrado<o:p></o:p></span></p>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=66 colspan=2 valign=top style='width:49.3pt;border:none;border-bottom:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:5.15pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:5.15pt;border:none' width=0 height=7></td>
  <![endif]>
 </tr>
 <tr style='height:25.55pt'>
  <td width=66 colspan=2 valign=top style='width:49.4pt;border:none;mso-border-top-alt:
  solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:25.55pt'>
  <p class=MsoNormal><span style='font-size:6.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Sacado:</span><span style='font-size:8.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=640 colspan=20 valign=top style='width:479.95pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:25.55pt'>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>[|NOME_SACADO|]&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[|NUMERO_INSCRICAO_SACADO|]<o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>[|ENDERECO_SACADO|]<o:p></o:p></span></b></p>
  <p class=MsoNormal><b><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>[|CEP_CIDADE_UF_SACADO|]</span></b><span style='font-size:8.0pt;mso-bidi-font-size:
  12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:25.55pt;border:none' width=0 height=34></td>
  <![endif]>
 </tr>
 <tr style='height:9.4pt'>
  <td width=617 colspan=18 valign=bottom style='width:463.1pt;border:none;
  border-bottom:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:9.4pt'>
  <p class=MsoNormal><span style='font-size:6.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>Sacador / Avalista: <b></b><o:p></o:p></span></p>
  </td>
  <td width=88 colspan=4 valign=bottom style='width:66.25pt;border:none;
  border-bottom:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:9.4pt'>
  <p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:9.4pt;border:none' width=0 height=13></td>
  <![endif]>
 </tr>
 <tr style='height:3.5pt'>
  <td width=654 colspan=21 valign=top style='width:490.25pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:3.5pt'>
  <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:4.5pt;mso-bidi-font-size:12.0pt;font-family:Arial'>Autenticação
  Mecânica</span><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=52 valign=top style='width:39.1pt;border:none;mso-border-top-alt:
  solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:3.5pt'>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:3.5pt;border:none' width=0 height=5></td>
  <![endif]>
 </tr>
 <tr style='height:3.5pt'>
  <td width=706 colspan=22 valign=top style='width:529.35pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:3.5pt'>
  <input type=hidden id=""c_codigo_barras_loaded"" value=""N"">
  <p class=MsoNormal><img width=406 height=51 id=""_x0000_i1093""
  width=406 height=51 src=""http://central85.com.br/e_x_e_c/cb.exe?code=[|CODIGO_BARRAS|]"" onload=""c_codigo_barras_loaded.value='S';""><span style='font-size:
  5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:3.5pt;border:none' width=0 height=5></td>
  <![endif]>
 </tr>
 <tr style='height:3.5pt'>
  <td width=706 colspan=22 valign=top style='width:529.35pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:3.5pt'>
  <p class=MsoNormal><span style='font-size:5.0pt;mso-bidi-font-size:12.0pt;
  font-family:Arial'>.................................................................................................................................................................................................................................................................................................................................................<o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:3.5pt;border:none' width=0 height=5></td>
  <![endif]>
 </tr>
 <tr style='height:3.5pt'>
  <td width=654 colspan=21 valign=top style='width:490.25pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:3.5pt'>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:4.5pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <td width=52 valign=top style='width:39.1pt;border:none;padding:0cm 3.5pt 0cm 3.5pt;
  height:3.5pt'>
  <p class=MsoNormal align=right style='text-align:right'><![if !supportEmptyParas]>&nbsp;<![endif]><span
  style='font-size:5.0pt;mso-bidi-font-size:12.0pt;font-family:Arial'><o:p></o:p></span></p>
  </td>
  <![if !supportMisalignedRows]>
  <td style='height:3.5pt;border:none' width=0 height=5></td>
  <![endif]>
 </tr>
 <![if !supportMisalignedColumns]>
 <tr height=0>
  <td width=15 style='border:none'></td>
  <td width=33 style='border:none'></td>
  <td width=17 style='border:none'></td>
  <td width=17 style='border:none'></td>
  <td width=34 style='border:none'></td>
  <td width=8 style='border:none'></td>
  <td width=13 style='border:none'></td>
  <td width=29 style='border:none'></td>
  <td width=8 style='border:none'></td>
  <td width=50 style='border:none'></td>
  <td width=13 style='border:none'></td>
  <td width=11 style='border:none'></td>
  <td width=77 style='border:none'></td>
  <td width=13 style='border:none'></td>
  <td width=50 style='border:none'></td>
  <td width=13 style='border:none'></td>
  <td width=88 style='border:none'></td>
  <td width=128 style='border:none'></td>
  <td width=1 style='border:none'></td>
  <td width=22 style='border:none'></td>
  <td width=14 style='border:none'></td>
  <td width=52 style='border:none'></td>
  <td style='height:3.5pt;border:none' width=0 height=5></td>
 </tr>
 <![endif]>
</table>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
</body>
</html>
";
		#endregion
	}
}
