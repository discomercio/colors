#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Media;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using mshtml;
#endregion

namespace Financeiro
{
	public partial class FBoletoHtml : Financeiro.FModelo
	{
		#region [ interface IHTMLElementRender ]
		// Replacement for mshtml imported interface, Tlbimp.exe generates wrong signatures
		[ComImport, InterfaceType((short)1), Guid("3050F669-98B5-11CF-BB82-00AA00BDCE0B")]
		private interface IHTMLElementRender
		{
			void DrawToDC(IntPtr hdc);
			void SetDocumentPrinter(string bstrPrinterName, IntPtr hdc);
		}
		#endregion

		#region [ interface IViewObject ]
		[ComVisible(true), ComImport()]
		[GuidAttribute("0000010d-0000-0000-C000-000000000046")]
		[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]

		private interface IViewObject
		{
			[return: MarshalAs(UnmanagedType.I4)]
			[PreserveSig]
			int Draw(
				//tagDVASPECT
				[MarshalAs(UnmanagedType.U4)] UInt32 dwDrawAspect,
				int lindex,
				IntPtr pvAspect,
				[In] IntPtr ptd,
				//[MarshalAs(UnmanagedType.Struct)] ref DVTARGETDEVICE ptd,
				IntPtr hdcTargetDev,
				IntPtr hdcDraw,
				[MarshalAs(UnmanagedType.Struct)] ref tagRECT lprcBounds,
				[MarshalAs(UnmanagedType.Struct)] ref tagRECT lprcWBounds,
				IntPtr pfnContinue,
				[MarshalAs(UnmanagedType.U4)] UInt32 dwContinue);
		}
		#endregion

		#region [ Atributos ]

		#region [ Diversos ]
		private bool _InicializacaoOk;
		private Form _formChamador = null;
		private String _emailCliente;
		private String _dataVencimento;
		private String _nomeCedente;
		private String _nomeSacado;
		private String _numInscricaoSacado;
		private String _enderecoSacado;
		private String _cepCidadeUfSacado;
		private String _numeroDocumento;
		private String _valorDocumento;
		private String _linhaDigitavel;
		private String _assuntoSelecionado = "";
		private String _destinatarioParaSelecionado = "";
		private String _destinatarioCopiaSelecionado = "";

		private BoletoHtml _boletoHtml;
		FConfiguracao fConfiguracao;
		FEmailParametros fEmailParametros;
		#endregion

		#endregion

		#region [ Construtor ]
		public FBoletoHtml( Form formChamador,
							String emailCliente,
							String dataVencimento,
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
							String fichaCompensacaoInstrucoesLinha6
							)
		{
			InitializeComponent();

			_formChamador = formChamador;
			_emailCliente = emailCliente;
			_dataVencimento = dataVencimento;
			_nomeCedente = nomeCedente;
			_nomeSacado = nomeSacado;
			_numInscricaoSacado = numInscricaoSacado;
			_enderecoSacado = enderecoSacado;
			_cepCidadeUfSacado = cepCidadeUfSacado;
			_numeroDocumento = numeroDocumento;
			_valorDocumento = valorDocumento;
			_linhaDigitavel = linhaDigitavel;

			_assuntoSelecionado = "Cobrança Bradesco, boleto bancário, vencimento: " + dataVencimento;
			_destinatarioParaSelecionado = _emailCliente;

			#region [ Cria instância de BoletoHtml ]
			_boletoHtml = new BoletoHtml(	dataVencimento,
											valorDocumento,
											dataProcessamento,
											dataDocumento,
											numeroDocumento,
											nomeCedente,
											carteira,
											agenciaECodigoCedente,
											nossoNumero,
											nomeSacado,
                                            numInscricaoSacado,
                                            enderecoSacado,
											cepCidadeUfSacado,
											linhaDigitavel,
											codigoBarras,
											reciboSacadoInstrucoesLinha1,
											reciboSacadoInstrucoesLinha2,
											reciboSacadoInstrucoesLinha3,
											reciboSacadoInstrucoesLinha4,
											reciboSacadoInstrucoesLinha5,
											reciboSacadoInstrucoesLinha6,
											fichaCompensacaoInstrucoesLinha1,
											fichaCompensacaoInstrucoesLinha2,
											fichaCompensacaoInstrucoesLinha3,
											fichaCompensacaoInstrucoesLinha4,
											fichaCompensacaoInstrucoesLinha5,
											fichaCompensacaoInstrucoesLinha6
											);
			#endregion
		}
		#endregion

		#region [ Métodos privados ]

		#region [ trataBotaoPrinterDialog ]
		private void trataBotaoPrinterDialog()
		{
			prnDialogConsulta.ShowDialog();
		}
		#endregion

		#region [ trataBotaoImprimir ]
		private void trataBotaoImprimir()
		{
			info(ModoExibicaoMensagemRodape.EmExecucao, "imprimindo o boleto");
			webBrowser.Print();
			info(ModoExibicaoMensagemRodape.Normal);
		}
		#endregion

		#region [ trataBotaoEnviaBoletoPorEmail ]
		private void trataBotaoEnviaBoletoPorEmail()
		{
			#region [ Declarações ]
			int intScrollRectangleHeight;
			String strDestinatarioPara;
			String strDestinatarioCopia;
			String strCorpoEmail;
			DialogResult drResultado;
			MailMessage mailMensagem;
			SmtpClient smtpCliente;
			MailAddress mailAddressFrom;
			MailAddress mailAddressTo;
			MailAddress mailAddressCc;
			MailAddress mailAddressBcc;
			Attachment attachment;
			MemoryStream msStream;
			String[] v;
			DateTime dtHrInicioEspera;
			WebBrowser wb;
			HtmlElement c_codigo_barras_loaded;
			String strOuterHtml;
			bool blnCodigoBarrasLoaded;
			int intTentativas;
			#endregion

			try
			{

				#region [ Consistência dos parâmetros de envio de e-mails ]
				if (Global.Usuario.fin_servidor_smtp_endereco.Trim().Length == 0)
				{
					if (!confirma("É necessário configurar os parâmetros para envio de e-mails!!\nDeseja configurar agora?")) return;
					if (!configuraParametrosEmail())
					{
						avisoErro("Os parâmetros para envio de e-mails não foram configurados corretamente!!");
						return;
					}
				}
				#endregion

				#region [ Obtém dados p/ enviar o e-mail ]
				fEmailParametros = new FEmailParametros(Global.Usuario.fin_email_remetente,
														Global.Usuario.fin_display_name_remetente,
														_assuntoSelecionado,
														_destinatarioParaSelecionado,
														_destinatarioCopiaSelecionado);
				fEmailParametros.StartPosition = FormStartPosition.Manual;
				fEmailParametros.Left = this.Left + (this.Width - fEmailParametros.Width) / 2;
				fEmailParametros.Top = this.Top + (this.Height - fEmailParametros.Height) / 2;
				drResultado = fEmailParametros.ShowDialog();

				if (drResultado != DialogResult.OK)
				{
					avisoErro("Envio do e-mail foi cancelado!!");
					return;
				}

				_assuntoSelecionado = fEmailParametros.assuntoEmail;
				_destinatarioParaSelecionado = fEmailParametros.destinatarioPara;
				_destinatarioCopiaSelecionado = fEmailParametros.destinatarioCopia;
				#endregion

				#region [ Prepara e-mail ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "preparando o e-mail");
				mailMensagem = new MailMessage();
				smtpCliente = new SmtpClient();

				mailAddressFrom = new MailAddress(Global.Usuario.fin_email_remetente, Global.Usuario.fin_display_name_remetente);
				mailMensagem.From = mailAddressFrom;

				#region[ Preenche destinatário ]
				strDestinatarioPara = _destinatarioParaSelecionado.Trim();
				if (strDestinatarioPara.Length > 0)
				{
					strDestinatarioPara = strDestinatarioPara.Replace("\n", " ");
					strDestinatarioPara = strDestinatarioPara.Replace("\r", " ");
					strDestinatarioPara = strDestinatarioPara.Replace(",", " ");
					strDestinatarioPara = strDestinatarioPara.Replace(";", " ");
					v = strDestinatarioPara.Split(' ');
					for (int i = 0; i < v.Length; i++)
					{
						if (v[i].Trim().Length > 0)
						{
							mailAddressTo = new MailAddress(v[i].Trim());
							mailMensagem.To.Add(mailAddressTo);
						}
					}
				}
				#endregion

				#region [ Se houver Cópia Para, preenche campos ]
				strDestinatarioCopia = _destinatarioCopiaSelecionado.Trim();
				if (strDestinatarioCopia.Length > 0)
				{
					strDestinatarioCopia = strDestinatarioCopia.Replace("\n", " ");
					strDestinatarioCopia = strDestinatarioCopia.Replace("\r", " ");
					strDestinatarioCopia = strDestinatarioCopia.Replace(",", " ");
					strDestinatarioCopia = strDestinatarioCopia.Replace(";", " ");
					v = strDestinatarioCopia.Split(' ');
					for (int i = 0; i < v.Length; i++)
					{
						if (v[i].Trim().Length > 0)
						{
							mailAddressCc = new MailAddress(v[i].Trim());
							mailMensagem.CC.Add(mailAddressCc);
						}
					}
				}
				#endregion

				#region [ Bcc para o próprio remetente: fica como um comprovante ]
				mailAddressBcc = new MailAddress(Global.Usuario.fin_email_remetente);
				mailMensagem.Bcc.Add(mailAddressBcc);
				#endregion

				strCorpoEmail = "Data do envio: " + Global.formataDataDdMmYyyyComSeparador(DateTime.Now) +
								"\n\n" +
								"Prezado sacado," +
								"\n\n" +
								_nomeSacado + " - " + Global.formataCnpjCpf(_numInscricaoSacado) +
								"\n" +
								_enderecoSacado +
								"\n" +
								_cepCidadeUfSacado +
								"\n\n" +
								"Ref: Boleto de pagamento transmitido via e-mail" +
								"\n\n" +
								"Dados do título" +
								"\n" +
								"===============" +
								"\n" +
								"N. do título...: " + _numeroDocumento +
								"\n" +
								"Vencimento.....: " + _dataVencimento +
								"\n" +
								"Valor do título: " + _valorDocumento +
								"\n" +
								"Linha digitável: " + _linhaDigitavel +
								"\n\n\n" +
								"Instruções para impressão do boleto bancário:" +
								"\n" +
								"=============================================" +
								"\n\n" +
								"1) Dê um duplo \"click\" sobre o ícone do arquivo em anexo (Arquivo de imagem JPG)" +
								"\n" +
								"2) Imprima o boleto conforme instruções abaixo:" +
								"\n" +
								"   a) Utilize impressora jato de tinta ou laser;" +
								"\n" +
								"   b) Utilize papel de tamanho A4 (210x297mm);" +
								"\n" +
								"   c) Configure a impressora para modo Normal de Impressão;" +
								"\n\n\n" +
								"Atenciosamente," +
								"\n\n" +
								_nomeCedente +
								"\n";

				mailMensagem.Subject = _assuntoSelecionado;
				mailMensagem.Priority = MailPriority.High;
				mailMensagem.BodyEncoding = Encoding.GetEncoding("Windows-1252");
				mailMensagem.Body = strCorpoEmail;

				#region [ Cria componente WebBrowser para renderizar o html ]
				wb = new WebBrowser();
				wb.ScrollBarsEnabled = false;
				wb.AllowNavigation = true;
				wb.Width = 780;
				wb.Height = 1022;
				#endregion

				#region [ Laço de tentativas para renderizar o html ]
				intTentativas = 0;
				do
				{
					intTentativas++;

					#region [ Se após várias tentativas o problema persiste, tenta recriar o componente WebBrowser ]
					if (intTentativas > 10)
					{
						if (wb != null) wb.Dispose();
						wb = new WebBrowser();
						wb.ScrollBarsEnabled = false;
						wb.AllowNavigation = true;
						wb.Width = 780;
						wb.Height = 1022;
					}
					#endregion

					#region [ Carrega o html no WebBrowser ]
					if (wb.Document != null) wb.Document.OpenNew(true);
					wb.DocumentText = _boletoHtml.textoBoletoHtml;
					Application.DoEvents();
					#endregion

					#region [ Aguarda o WebBrowser processar o html ]
					dtHrInicioEspera = DateTime.Now;
					while (wb.ReadyState != WebBrowserReadyState.Complete)
					{
						Application.DoEvents();
						Thread.Sleep(200);
						if (Global.calculaTimeSpanSegundos(DateTime.Now - dtHrInicioEspera) > 120)
						{
							throw new FinanceiroException("Falha ao gerar a imagem do boleto: timeout na renderização do código html!!");
						}
					}
					Application.DoEvents();
					wb.Update();
					Application.DoEvents();
					#endregion

					#region [ wb.Document.Body.ScrollRectangle.Height ]
					intScrollRectangleHeight = 0;
					if (wb.Document != null)
					{
						if (wb.Document.Body != null)
						{
							if (wb.Document.Body.ScrollRectangle != null) intScrollRectangleHeight = wb.Document.Body.ScrollRectangle.Height;
						}
					}
					#endregion

					#region [ c_codigo_barras_loaded ]
					blnCodigoBarrasLoaded = false;
					c_codigo_barras_loaded = wb.Document.GetElementById("c_codigo_barras_loaded");
					if (c_codigo_barras_loaded != null)
					{
						strOuterHtml = c_codigo_barras_loaded.OuterHtml;
						if (strOuterHtml != null) blnCodigoBarrasLoaded = strOuterHtml.ToUpper().Contains("VALUE=S");
					}
					#endregion

					#region [ O html foi renderizado? ]
					if ((intScrollRectangleHeight < 600)
							||
						(!blnCodigoBarrasLoaded))
					{
						if (intTentativas > 100) throw new FinanceiroException("Falha ao gerar a imagem do boleto: falha na renderização do código html!!");
						Application.DoEvents();
					}
					#endregion

				} while ((intScrollRectangleHeight < 600)
							||
						 (!blnCodigoBarrasLoaded));
				#endregion

				#region [ Ajusta o tamanho do componente WebBrowser em função do conteúdo do html renderizado ]
				wb.Width = wb.Document.Body.ScrollRectangle.Width;
				wb.Height = wb.Document.Body.ScrollRectangle.Height;
				#endregion

				#region [ Captura a imagem do html renderizado pelo componente WebBrowser ]
				// Get the view object of the browser
				IViewObject VObject = wb.Document.DomDocument as IViewObject;
				// Construct a bitmap as big as the required image.
				Bitmap bmp = new Bitmap(wb.Document.Body.ClientRectangle.Width, wb.Document.Body.ClientRectangle.Height);
				// The size of the portion of the web page to be captured.
				mshtml.tagRECT SourceRect = new tagRECT();
				SourceRect.left = 0;
				SourceRect.top = 0;
				SourceRect.right = wb.Right;
				SourceRect.bottom = wb.Bottom;

				// The size to render the target image. This can be used to shrink the image to a thumbnail.
				mshtml.tagRECT TargetRect = new tagRECT();
				TargetRect.left = 0;
				TargetRect.top = 0;
				TargetRect.right = wb.Right;
				TargetRect.bottom = wb.Bottom;

				// Draw the web page into the bitmap.
				using (Graphics gr = Graphics.FromImage(bmp))
				{
					IntPtr hdc = gr.GetHdc();
					int hr =
						VObject.Draw((int)DVASPECT.DVASPECT_CONTENT,
							(int)-1, IntPtr.Zero, IntPtr.Zero,
							IntPtr.Zero, hdc, ref TargetRect, ref SourceRect,
							IntPtr.Zero, (uint)0);
					gr.ReleaseHdc();
				}
				#endregion

				#region [ Anexa no email a imagem capturada no formato jpeg ]
				msStream = new MemoryStream();
				bmp.Save(msStream, System.Drawing.Imaging.ImageFormat.Jpeg);
				msStream.Position = 0;
				attachment = new Attachment(msStream, "BOLETO.JPG", System.Net.Mime.MediaTypeNames.Image.Jpeg);
				mailMensagem.Attachments.Add(attachment);
				#endregion

				#region [ Transmite o e-mail ]
				smtpCliente.Host = Global.Usuario.fin_servidor_smtp_endereco;
				if (Global.Usuario.fin_servidor_smtp_porta > 0) smtpCliente.Port = Global.Usuario.fin_servidor_smtp_porta;
				smtpCliente.Credentials = new System.Net.NetworkCredential(Global.Usuario.fin_usuario_smtp, Global.Usuario.fin_senha_smtp);

				info(ModoExibicaoMensagemRodape.EmExecucao, "transmitindo o e-mail");
				smtpCliente.Send(mailMensagem);
				#endregion

				info(ModoExibicaoMensagemRodape.Normal);
				SystemSounds.Exclamation.Play();
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade("Falha ao enviar o e-mail!!\n\n" + ex.ToString());
				avisoErro("Falha ao enviar o e-mail!!\n\n" + ex.ToString());
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ configuraParametrosEmail ]
		private bool configuraParametrosEmail()
		{
			DialogResult drResultado;
			fConfiguracao = new FConfiguracao();
			fConfiguracao.StartPosition = FormStartPosition.Manual;
			fConfiguracao.Left = this.Left + (this.Width - fConfiguracao.Width) / 2;
			fConfiguracao.Top = this.Top + (this.Height - fConfiguracao.Height) / 2;
			drResultado = fConfiguracao.ShowDialog();
		
			if (drResultado == DialogResult.OK)
				return true;
			else
				return false;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FBoletoHtml ]

		#region [ FBoletoHtml_Load ]
		private void FBoletoHtml_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				#region [ Inicializa browser ]
				webBrowser.Navigate("about:blank");
				#endregion

				blnSucesso = true;
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				Close();
				return;
			}
			finally
			{
				if (!blnSucesso) Close();
			}
		}
		#endregion

		#region [ FBoletoHtml_Shown ]
		private void FBoletoHtml_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Permissão de acesso ao módulo ]
					if (!Global.Acesso.operacaoPermitida(Global.Acesso.OP_CEN_FIN_APP_FINANCEIRO_BOLETO_CADASTRAR))
					{
						btnEmail.Enabled = false;
					}
					#endregion

					#region [ Carrega dados de exibição do boleto ]
					info(ModoExibicaoMensagemRodape.EmExecucao, "Gerando exibição do boleto");
					webBrowser.DocumentText = _boletoHtml.textoBoletoHtml;
					#endregion

					#region [ Posiciona foco ]
					webBrowser.Focus();
					#endregion

					_InicializacaoOk = true;
				}
				#endregion
			}
			catch (Exception ex)
			{
				avisoErro(ex.ToString());
				Close();
				return;
			}
			finally
			{
				// Se não inicializou corretamente, assegura-se de que o painel será fechado
				if (!_InicializacaoOk) Close();
			}
		}
		#endregion

		#endregion

		#region [ webBrowser ]

		#region [ webBrowser_DocumentCompleted ]
		private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
		{
			info(ModoExibicaoMensagemRodape.Normal);
		}
		#endregion

		#endregion

		#region [ btnPrinterDialog ]

		#region [ btnPrinterDialog_Click ]
		private void btnPrinterDialog_Click(object sender, EventArgs e)
		{
			trataBotaoPrinterDialog();
		}
		#endregion

		#endregion

		#region [ btnImprimir ]

		#region [ btnImprimir_Click ]
		private void btnImprimir_Click(object sender, EventArgs e)
		{
			trataBotaoImprimir();
		}
		#endregion

		#endregion

		#region [ btnEmail ]

		#region [ btnEmail_Click ]
		private void btnEmail_Click(object sender, EventArgs e)
		{
			trataBotaoEnviaBoletoPorEmail();
		}
		#endregion

		#endregion

		#endregion
	}
}
