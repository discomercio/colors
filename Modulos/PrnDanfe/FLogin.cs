#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
#endregion

namespace PrnDANFE
{
	public partial class FLogin : Form
	{
		#region [ Atributos ]
		private static String _usuario;
		public static String usuario
		{
			get { return FLogin._usuario; }
			set { FLogin._usuario = value; }
		}

		private static String _senha;
		public static String senha
		{
			get { return FLogin._senha; }
			set { FLogin._senha = value; }
		}

		private bool _inicializacaoOk;
		#endregion

		#region [ Construtor ]
		public FLogin()
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion

			_inicializacaoOk = false;
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
			if (txtUsuario.Text.Length == 0)
			{
				txtUsuario.Focus();
				aviso("Informe o usuário!!");
				return;
			}
			if (txtSenha.Text.Length == 0)
			{
				txtSenha.Focus();
				aviso("Informe a senha!!");
				return;
			}

			usuario = txtUsuario.Text.Trim();
			senha = txtSenha.Text.Trim();
			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form FLogin ]

		#region [ FLogin_Load ]
		private void FLogin_Load(object sender, EventArgs e)
		{
			txtUsuario.Text = usuario;
			txtSenha.Text = "";
			usuario = "";
			senha = "";
			this.Text += " - " + Global.Cte.Aplicativo.M_ID;
		}
		#endregion

		#region [ FLogin_Shown ]
		private void FLogin_Shown(object sender, EventArgs e)
		{
			if (!_inicializacaoOk)
			{
				if (txtUsuario.Text.Trim().Length > 0) txtSenha.Focus();
				_inicializacaoOk = true;
			}
		}
		#endregion

		#endregion

		#region [ txtUsuario ]

		#region [ txtUsuario_KeyDown ]
		private void txtUsuario_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				if (txtUsuario.Text.Length == 0) return;
				if (txtSenha.Text.Length == 0)
				{
					txtSenha.Focus();
					return;
				}
				ProcessaBotaoOk();
			}
		}
		#endregion

		#region [ txtUsuario_Enter ]
		private void txtUsuario_Enter(object sender, EventArgs e)
		{
			txtUsuario.Select(0, txtUsuario.Text.Length);
		}
		#endregion

		#endregion

		#region [ txtSenha ]

		#region [ txtSenha_KeyDown ]
		private void txtSenha_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				if (txtSenha.Text.Length == 0) return;
				if (txtUsuario.Text.Length == 0)
				{
					txtUsuario.Focus();
					return;
				}
				ProcessaBotaoOk();
			}
		}
		#endregion

		#region [ txtSenha_Enter ]
		private void txtSenha_Enter(object sender, EventArgs e)
		{
			txtSenha.Select(0, txtSenha.Text.Length);
		}
		#endregion

		#endregion

		#region [ Botões ]

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

		#endregion

		#endregion
	}
}
