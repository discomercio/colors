#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Media;
#endregion

namespace Financeiro
{
	public partial class FConfiguracao : Financeiro.FModelo
	{
		#region [ Atributos ]

		#region [ Diversos ]
		private bool _InicializacaoOk;
		#endregion

		#endregion

		#region [ Construtor ]
		public FConfiguracao()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos privados ]

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtServidorSmtp.Text = "";
			txtServidorSmtpPorta.Text = "";
			txtEmailRemetente.Text = "";
			txtDisplayNameRemetente.Text = "";
			txtUsuarioSmtp.Text = "";
			txtSenhaSmtp.Text = "";
		}
		#endregion

		#region [ consisteCamposEmail ]
		private bool consisteCamposEmail()
		{
			#region [ Declarações ]
			String strRelacaoEmailInvalido = "";
			#endregion

			#region [ txtServidorSmtp ]
			if (txtServidorSmtp.Text.Trim().Length == 0)
			{
				avisoErro("Informe o endereço do servidor SMTP para envio de e-mails!!");
				txtServidorSmtp.Focus();
				return false;
			}
			#endregion

			#region [ txtServidorSmtpPorta ]
			if (Global.converteInteiro(txtServidorSmtpPorta.Text) == 0)
			{
				avisoErro("Informe a porta do servidor SMTP para envio de e-mails!!");
				txtServidorSmtpPorta.Focus();
				return false;
			}
			#endregion

			#region [ txtEmailRemetente ]
			if (txtEmailRemetente.Text.Trim().Length == 0)
			{
				avisoErro("Informe o endereço de e-mail que será usado para enviar os e-mails!!");
				txtEmailRemetente.Focus();
				return false;
			}

			if (!Global.isEmailOk(txtEmailRemetente.Text, ref strRelacaoEmailInvalido))
			{
				avisoErro("O endereço de e-mail informado é inválido!!");
				txtEmailRemetente.Focus();
				return false;
			}
			#endregion

			#region [ txtDisplayNameRemetente ]
			if (txtDisplayNameRemetente.Text.Trim().Length == 0)
			{
				avisoErro("Informe o nome do remetente!!");
				txtDisplayNameRemetente.Focus();
				return false;
			}
			#endregion

			#region [ txtUsuarioSmtp ]
			if (txtUsuarioSmtp.Text.Trim().Length == 0)
			{
				avisoErro("Informe o usuário para ser usado na autenticação ao conectar com o servidor SMTP de e-mail!!");
				txtUsuarioSmtp.Focus();
				return false;
			}
			#endregion

			#region [ txtSenhaSmtp ]
			if (txtSenhaSmtp.Text.Trim().Length == 0)
			{
				avisoErro("Informe a senha para ser usada na autenticação ao conectar com o servidor SMTP de e-mail!!");
				txtSenhaSmtp.Focus();
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ trataBotaoConfirma ]
		private void trataBotaoConfirma()
		{
			#region [ Declarações ]
			String strMsgErro = "";
			#endregion

			#region [ Verifica se a conexão c/ o BD está ok ]
			if (!BD.isConexaoOk())
			{
				if (!FMain.reiniciaBancoDados())
				{
					avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
					return;
				}
			}
			#endregion

			if (!consisteCamposEmail()) return;

			if (!UsuarioDAO.atualizaFinEmail(Global.Usuario.usuario,
											Global.Usuario.usuario,
											txtEmailRemetente.Text.Trim(),
											txtDisplayNameRemetente.Text.Trim(),
											txtServidorSmtp.Text.Trim(),
											(int)Global.converteInteiro(txtServidorSmtpPorta.Text.Trim()),
											txtUsuarioSmtp.Text.Trim(),
											txtSenhaSmtp.Text.Trim(),
											ref strMsgErro))
			{
				avisoErro("Falha ao tentar atualizar o banco de dados!!\n" + strMsgErro);
				return;
			}

			#region [ Atualiza dados carregados na memória ]
			Global.Usuario.fin_email_remetente = txtEmailRemetente.Text.Trim();
			Global.Usuario.fin_display_name_remetente = txtDisplayNameRemetente.Text.Trim();
			Global.Usuario.fin_servidor_smtp_endereco = txtServidorSmtp.Text.Trim();
			Global.Usuario.fin_servidor_smtp_porta = (int)Global.converteInteiro(txtServidorSmtpPorta.Text);
			Global.Usuario.fin_usuario_smtp = txtUsuarioSmtp.Text.Trim();
			Global.Usuario.fin_senha_smtp = txtSenhaSmtp.Text.Trim();
			#endregion

			SystemSounds.Exclamation.Play();
			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FConfiguracao ]

		#region [ FConfiguracao_Load ]
		private void FConfiguracao_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				limpaCampos();

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

		#region [ FConfiguracao_Shown ]
		private void FConfiguracao_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Preenche campos ]
					txtEmailRemetente.Text = Global.Usuario.fin_email_remetente;
					txtDisplayNameRemetente.Text = Global.Usuario.fin_display_name_remetente;
					txtServidorSmtp.Text = Global.Usuario.fin_servidor_smtp_endereco;
					txtServidorSmtpPorta.Text = Global.Usuario.fin_servidor_smtp_porta.ToString();
					txtUsuarioSmtp.Text = Global.Usuario.fin_usuario_smtp;
					txtSenhaSmtp.Text = Global.Usuario.fin_senha_smtp;
					#endregion

					#region [ Posiciona foco ]
					txtServidorSmtp.Focus();
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
				info(ModoExibicaoMensagemRodape.Normal);
				// Se não inicializou corretamente, assegura-se de que o painel será fechado
				if (!_InicializacaoOk) Close();
			}
		}
		#endregion

		#endregion

		#region [ txtServidorSmtp ]

		#region [ txtServidorSmtp_Enter ]
		private void txtServidorSmtp_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtServidorSmtp_Leave ]
		private void txtServidorSmtp_Leave(object sender, EventArgs e)
		{
			txtServidorSmtp.Text = txtServidorSmtp.Text.Trim();
		}
		#endregion

		#region [ txtServidorSmtp_KeyDown ]
		private void txtServidorSmtp_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtServidorSmtpPorta);
		}
		#endregion

		#region [ txtServidorSmtp_KeyPress ]
		private void txtServidorSmtp_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtServidorSmtpPorta ]

		#region [ txtServidorSmtpPorta_Enter ]
		private void txtServidorSmtpPorta_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtServidorSmtpPorta_Leave ]
		private void txtServidorSmtpPorta_Leave(object sender, EventArgs e)
		{
			txtServidorSmtpPorta.Text = txtServidorSmtpPorta.Text.Trim();
		}
		#endregion

		#region [ txtServidorSmtpPorta_KeyDown ]
		private void txtServidorSmtpPorta_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtEmailRemetente);
		}
		#endregion

		#region [ txtServidorSmtpPorta_KeyPress ]
		private void txtServidorSmtpPorta_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtEmailRemetente ]

		#region [ txtEmailRemetente_Enter ]
		private void txtEmailRemetente_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtEmailRemetente_Leave ]
		private void txtEmailRemetente_Leave(object sender, EventArgs e)
		{
			txtEmailRemetente.Text = txtEmailRemetente.Text.Trim();
		}
		#endregion

		#region [ txtEmailRemetente_KeyDown ]
		private void txtEmailRemetente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDisplayNameRemetente);
		}
		#endregion

		#region [ txtEmailRemetente_KeyPress ]
		private void txtEmailRemetente_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDisplayNameRemetente ]

		#region [ txtDisplayNameRemetente_Enter ]
		private void txtDisplayNameRemetente_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDisplayNameRemetente_Leave ]
		private void txtDisplayNameRemetente_Leave(object sender, EventArgs e)
		{
			txtDisplayNameRemetente.Text = txtDisplayNameRemetente.Text.Trim();
		}
		#endregion

		#region [ txtDisplayNameRemetente_KeyDown ]
		private void txtDisplayNameRemetente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtUsuarioSmtp);
		}
		#endregion

		#region [ txtDisplayNameRemetente_KeyPress ]
		private void txtDisplayNameRemetente_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtUsuarioSmtp ]

		#region [ txtUsuarioSmtp_Enter ]
		private void txtUsuarioSmtp_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtUsuarioSmtp_Leave ]
		private void txtUsuarioSmtp_Leave(object sender, EventArgs e)
		{
			txtUsuarioSmtp.Text = txtUsuarioSmtp.Text.Trim();
		}
		#endregion

		#region [ txtUsuarioSmtp_KeyDown ]
		private void txtUsuarioSmtp_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtSenhaSmtp);
		}
		#endregion

		#region [ txtUsuarioSmtp_KeyPress ]
		private void txtUsuarioSmtp_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtSenhaSmtp ]

		#region [ txtSenhaSmtp_Enter ]
		private void txtSenhaSmtp_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtSenhaSmtp_Leave ]
		private void txtSenhaSmtp_Leave(object sender, EventArgs e)
		{
			txtSenhaSmtp.Text = txtSenhaSmtp.Text.Trim();
		}
		#endregion

		#region [ txtSenhaSmtp_KeyDown ]
		private void txtSenhaSmtp_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnConfirma);
		}
		#endregion

		#region [ txtSenhaSmtp_KeyPress ]
		private void txtSenhaSmtp_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ btnConfirma ]

		#region [ btnConfirma_Click ]
		private void btnConfirma_Click(object sender, EventArgs e)
		{
			trataBotaoConfirma();
		}
		#endregion

		#endregion

		#endregion
	}
}
