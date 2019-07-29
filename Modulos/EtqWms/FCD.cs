using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EtqWms
{
	public partial class FCD : Form
	{
		#region [ Construtor ]
		public FCD()
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion
		}
		#endregion

		#region [ Atributos ]
		private static String _usuEmit;
		public static String usuEmit
		{
			get { return FCD._usuEmit; }
			set { FCD._usuEmit = value; }
		}
		#endregion

		#region [ Métodos ]

		#region[ aviso ]
		public void aviso(string mensagem)
		{
			MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Warning);
		}
		#endregion


		#region [ carregaEmitentes ]
		private void carregaEmitentes()
		{
			#region [ Declarações ]
			int i;
			#endregion


				cbEmitente.Items.Clear();
				cbEmitente.Items.Add("");
				for (i = 0; i < Global.Usuario.listaEmitentes.Count; i++)
				{
					cbEmitente.Items.Add(Global.Usuario.listaEmitentes[i].emit);
				}
		}

		#endregion

		#region [ Processa Botão OK ]

		private void ProcessaBotaoOk()
		{
			int i = 0;

			if (cbEmitente.Text.Trim() == "")
			{
				aviso("Selecione um Emitente!!!");
				return;
			}

			i = cbEmitente.SelectedIndex - 1;

			Global.Usuario.emit = Global.Usuario.listaEmitentes[i].emit;
			Global.Usuario.emit_uf = Global.Usuario.listaEmitentes[i].emit_uf;
			Global.Usuario.emit_id = Global.Usuario.listaEmitentes[i].emit_id;
			Global.Usuario.txtEspecifico = Global.Usuario.listaEmitentes[i].emit_texto_especifico;

			this.DialogResult = DialogResult.OK;

		}

		#endregion

		#endregion

		#region [ Eventos ]

		private void FCD_Load(object sender, EventArgs e)
		{
			int i;
			bool achou;

			carregaEmitentes();

			//se houver um emitente padrão para o usuário, pré-selecionar
			if (usuEmit != "")
			{
				i = 0;
				achou = false;
				do
				{
					if (Global.Usuario.listaEmitentes[i].emit_id == usuEmit)
					{
						cbEmitente.SelectedIndex = i + 1;
						achou = true;
					}
					i = i + 1;
				} while ((!achou) && (i < Global.Usuario.listaEmitentes.Count));
			}
		}



		private void btnOk_Click(object sender, EventArgs e)
		{
			ProcessaBotaoOk();
		}

		private void btnCancela_Click(object sender, EventArgs e)
		{
			this.DialogResult = DialogResult.Cancel;
		}

		private void cbEmitente_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				ProcessaBotaoOk();
			}
		}

		#endregion

	}
}
