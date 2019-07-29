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
	public partial class FBoletoAvulsoComPedidoSelPedido : Form
	{
		#region [ Atributos ]
		private List<String> _listaPedidosSelecionados = new List<String>();

		public List<String> listaPedidosSelecionados
		{
			get { return _listaPedidosSelecionados; }
		}
		#endregion

		#region [ Construtor ]
		public FBoletoAvulsoComPedidoSelPedido()
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion
		}
		#endregion

		#region [ Métodos Privados ]

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

		#region [ normalizaCampoTxtPedido ]
		private void normalizaCampoTxtPedido()
		{
			String strListaPedidosNormalizado = "";

			for (int i = 0; i < txtPedido.Lines.Length; i++)
			{
				if (txtPedido.Lines[i].ToString().Trim().Length > 0)
				{
					if (strListaPedidosNormalizado.Length > 0) strListaPedidosNormalizado += "\n";
					strListaPedidosNormalizado += Global.normalizaNumeroPedido(txtPedido.Lines[i]);
				}
			}
			txtPedido.Text = strListaPedidosNormalizado;
		}
		#endregion

		#region [ ProcessaBotaoOk ]
		private void ProcessaBotaoOk()
		{
			normalizaCampoTxtPedido();

			for (int i = 0; i < (txtPedido.Lines.Length - 1); i++)
			{
				for (int j = i + 1; j < txtPedido.Lines.Length; j++)
				{
					if (txtPedido.Lines[i].Equals(txtPedido.Lines[j]))
					{
						avisoErro("O pedido " + txtPedido.Lines[i] + " está repetido na lista!!");
						return;
					}
				}
			}

			_listaPedidosSelecionados.Clear();
			for (int i = 0; i < txtPedido.Lines.Length; i++)
			{
				_listaPedidosSelecionados.Add(txtPedido.Lines[i]);
			}

			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FBoletoAvulsoComPedidoSelPedido ]

		#region [ FBoletoAvulsoComPedidoSelPedido_Load ]
		private void FBoletoAvulsoComPedidoSelPedido_Load(object sender, EventArgs e)
		{
			txtPedido.Text = "";
			txtPedido.Focus();
		}
		#endregion

		#endregion

		#region [ txtPedido ]

		#region [ txtPedido_KeyPress ]
		private void txtPedido_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#region [ txtPedido_Leave ]
		private void txtPedido_Leave(object sender, EventArgs e)
		{
			normalizaCampoTxtPedido();
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
