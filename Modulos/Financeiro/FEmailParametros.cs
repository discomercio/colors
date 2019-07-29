#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
#endregion

namespace Financeiro
{
	public partial class FEmailParametros : Form
	{
		#region [ Atributos ]

		#region [ Diversos ]
		private bool _InicializacaoOk;
		private String _enderecoRemetente = "";
		private String _displayNameRemetente = "";
		private String _assuntoDefault = "";
		private String _destinatarioParaDefault = "";
		private String _destinatarioCopiaDefault = "";
		public String assuntoEmail = "";
		public String destinatarioPara = "";
		public String destinatarioCopia = "";
		#endregion
		
		#endregion

		#region [ Construtor ]
		public FEmailParametros(String enderecoRemetente, String displayNameRemetente, String assuntoDefault, String destinatarioParaDefault, String destinatarioCopiaDefault)
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion

			_enderecoRemetente = enderecoRemetente;
			_displayNameRemetente = displayNameRemetente;
			_assuntoDefault = assuntoDefault;
			_destinatarioParaDefault = destinatarioParaDefault;
			_destinatarioCopiaDefault = destinatarioCopiaDefault;
		}
		#endregion

		#region [ Métodos privados ]

		#region[ aviso ]
		public void aviso(string mensagem)
		{
			MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
		#endregion

		#region[ avisoErro ]
		public void avisoErro(string mensagem)
		{
			MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
		#endregion

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtAssunto.Text = "";
			txtDestinatarioPara.Text = "";
			txtDestinatarioCopia.Text = "";
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Declarações ]
			String strRelacaoEmailInvalido = "";
			String[] v;
			String strMsgErro;
			#endregion

			if (txtAssunto.Text.Trim().Length == 0)
			{
				avisoErro("Informe o texto para ser usado como assunto do e-mail!!");
				txtAssunto.Focus();
				return false;
			}

			if ((txtDestinatarioPara.Text.Trim().Length == 0) && (txtDestinatarioCopia.Text.Trim().Length == 0))
			{
				avisoErro("Informe o destinatário do e-mail!!");
				txtDestinatarioPara.Focus();
				return false;
			}

			if (txtDestinatarioPara.Text.Trim().Length > 0)
			{
				if (!Global.isEmailOk(txtDestinatarioPara.Text.Trim(), ref strRelacaoEmailInvalido))
				{
					strMsgErro = "";
					v = strRelacaoEmailInvalido.Split(' ');
					for (int i = 0; i < v.Length; i++)
					{
						if (strMsgErro.Length > 0) strMsgErro += "\n";
						strMsgErro += v[i];
					}
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "E-mail inválido:" + strMsgErro;
					avisoErro(strMsgErro);
					txtDestinatarioPara.Focus();
					return false;
				}
			}

			if (txtDestinatarioCopia.Text.Trim().Length > 0)
			{
				if (!Global.isEmailOk(txtDestinatarioCopia.Text.Trim(), ref strRelacaoEmailInvalido))
				{
					strMsgErro = "";
					v = strRelacaoEmailInvalido.Split(' ');
					for (int i = 0; i < v.Length; i++)
					{
						if (strMsgErro.Length > 0) strMsgErro += "\n";
						strMsgErro += v[i];
					}
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "E-mail inválido:" + strMsgErro;
					avisoErro(strMsgErro);
					txtDestinatarioCopia.Focus();
					return false;
				}
			}

			return true;
		}
		#endregion

		#region [ trataBotaoOk ]
		private void trataBotaoOk()
		{
			if (!consisteCampos()) return;

			#region [ Preenche campos com os dados editados ]
			assuntoEmail = txtAssunto.Text.Trim();
			destinatarioPara = txtDestinatarioPara.Text.Trim();
			destinatarioCopia = txtDestinatarioCopia.Text.Trim();
			#endregion

			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#region [ trataBotaoCancela ]
		private void trataBotaoCancela()
		{
			this.DialogResult = DialogResult.Cancel;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FEmailParametros ]

		#region [ FEmailParametros_Load ]
		private void FEmailParametros_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				#region [ Limpa campos ]
				limpaCampos();
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

		#region [ FEmailParametros_Shown ]
		private void FEmailParametros_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Preenche dados ]
					lblEmailRemetente.Text = _displayNameRemetente + " (" + _enderecoRemetente + ")";
					txtAssunto.Text = _assuntoDefault;
					txtDestinatarioPara.Text = _destinatarioParaDefault;
					txtDestinatarioCopia.Text = _destinatarioCopiaDefault;
					#endregion

					#region [ Posiciona foco ]
					if (txtAssunto.Text.Length == 0)
						txtAssunto.Focus();
					else
						txtDestinatarioPara.Focus();
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

		#region [ txtAssunto ]

		#region [ txtAssunto_Enter ]
		private void txtAssunto_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtAssunto_Leave ]
		private void txtAssunto_Leave(object sender, EventArgs e)
		{
			txtAssunto.Text = txtAssunto.Text.Trim();
		}
		#endregion

		#region [ txtAssunto_KeyDown ]
		private void txtAssunto_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtDestinatarioPara);
		}
		#endregion

		#region [ txtAssunto_KeyPress ]
		private void txtAssunto_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDestinatarioPara ]

		#region [ txtDestinatarioPara_Enter ]
		private void txtDestinatarioPara_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDestinatarioPara_Leave ]
		private void txtDestinatarioPara_Leave(object sender, EventArgs e)
		{
			txtDestinatarioPara.Text = txtDestinatarioPara.Text.Trim();
		}
		#endregion

		#region [ txtDestinatarioPara_KeyPress ]
		private void txtDestinatarioPara_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtDestinatarioCopia ]

		#region [ txtDestinatarioCopia_Enter ]
		private void txtDestinatarioCopia_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDestinatarioCopia_Leave ]
		private void txtDestinatarioCopia_Leave(object sender, EventArgs e)
		{
			txtDestinatarioCopia.Text = txtDestinatarioCopia.Text.Trim();
		}
		#endregion

		#region [ txtDestinatarioCopia_KeyPress ]
		private void txtDestinatarioCopia_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ btnOk ]

		#region [ btnOk_Click ]
		private void btnOk_Click(object sender, EventArgs e)
		{
			trataBotaoOk();
		}
		#endregion

		#endregion

		#region [ btnCancela ]

		#region [ btnCancela_Click ]
		private void btnCancela_Click(object sender, EventArgs e)
		{
			trataBotaoCancela();
		}
		#endregion

		#endregion

		#endregion
	}
}
