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
	public partial class FAutorizacao : Form
	{
		#region [ Atributos ]
		private String _mensagem;
		public String mensagem
		{
			get { return _mensagem; }
			set { _mensagem = value; }
		}

		private String _senha;
		public String senha
		{
			get { return _senha; }
			set { _senha = value; }
		}
		#endregion

		#region [ Métodos ]

		#region[ aviso ]
		public void aviso(string mensagem)
		{
			MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Warning);
		}
		#endregion

		#region [ ProcessaBotaoOk ]
		private void ProcessaBotaoOk()
		{
			if (txtSenha.Text.Trim().Length == 0)
			{
				aviso("Digite a senha!!");
				txtSenha.Focus();
				return;
			}

			senha = txtSenha.Text.Trim();
			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#endregion

		#region [ Construtor ]
		public FAutorizacao(String mensagem)
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion

			_mensagem = mensagem;
		}
		#endregion

		#region [ Eventos ]

		#region [ FAutorizacao ]

		#region [ FAutorizacao_Load ]
		private void FAutorizacao_Load(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String strMensagemHtml;
			#endregion

			#region [ Inicializa browser ]
			webBrowserMensagem.Navigate("about:blank");
			#endregion

			strMensagemHtml = "<html>" +"\n" +
							  "<body>" + "\n" +
							  "<center>" + "\n" +
							  "<span style='font-family:Arial,Helvetica,sans-serif;font-size:10pt;font-weight:bold;'>" + "\n" +
							  mensagem.Replace("\n", "<br>") + "\n" +
							  "</span>" + "\n" +
							  "</center>" + "\n" +
							  "</body>" + "\n" +
							  "</html>";
			webBrowserMensagem.DocumentText = strMensagemHtml;
			senha = "";
			txtSenha.Text = "";
		}
		#endregion

		#region [ FAutorizacao_Shown ]
		private void FAutorizacao_Shown(object sender, EventArgs e)
		{
			txtSenha.Focus();
		}
		#endregion

		#endregion

		#region [ btnOk_Click ]
		private void btnOk_Click(object sender, EventArgs e)
		{
			ProcessaBotaoOk();
		}
		#endregion

		#region [ btnCancela_Click ]
		private void btnCancela_Click(object sender, EventArgs e)
		{
			this.DialogResult = DialogResult.Cancel;
		}
		#endregion

		#region [ txtSenha_KeyDown ]
		private void txtSenha_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				if (txtSenha.Text.Length == 0) return;
				ProcessaBotaoOk();
			}
		}
		#endregion

		#endregion
	}
}
