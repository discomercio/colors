using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EtqWms
{
	public partial class FVolumeSeleciona : Form
	{
		#region [ enum ]
		public enum eOpcaoSelecao
		{
			VOLUME_UNICO = 1,
			VOLUME_INTERVALO = 2
		}
		#endregion

		#region [ Atributos ]
		private List<EtiquetaDados> _listaEtqParcial;

		private int _numNF;
		public int numNF
		{
			get { return _numNF; }
			set { _numNF = value; }
		}

		private int _numVolumeUnico;
		public int numVolumeUnico
		{
			get { return _numVolumeUnico; }
			set { _numVolumeUnico = value; }
		}

		private int _numVolumeIntervaloInicio;
		public int numVolumeIntervaloInicio
		{
			get { return _numVolumeIntervaloInicio; }
			set { _numVolumeIntervaloInicio = value; }
		}

		private int _numVolumeIntervaloFim;
		public int numVolumeIntervaloFim
		{
			get { return _numVolumeIntervaloFim; }
			set { _numVolumeIntervaloFim = value; }
		}

		private eOpcaoSelecao _opcaoSelecionada;
		public eOpcaoSelecao opcaoSelecionada
		{
			get { return _opcaoSelecionada; }
			set { _opcaoSelecionada = value; }
		}
		#endregion

		#region [ Construtor ]
		public FVolumeSeleciona(List<EtiquetaDados> listaEtqParcial)
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion

			_listaEtqParcial = listaEtqParcial;
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
			#region [ Declarações ]
			bool blnNfOk = false;
			bool blnNumVolOk = false;
			int numNfAux;
			string strCampoObs;
			#endregion

			#region [ Consiste campos ]
			if (txtNF.Text.Trim().Length == 0)
			{
				aviso("Informe o nº da NF da etiqueta a ser impressa!!");
				txtNF.Focus();
				return;
			}

			if (Global.converteInteiro(txtNF.Text) <= 0)
			{
				aviso("Nº da NF preenchido com valor inválido!!");
				txtNF.Focus();
				return;
			}

			if ((!rbVolumeUnico.Checked) && (!rbVolumeRange.Checked))
			{
				aviso("Selecione uma opção (volume único ou intervalo de volumes)!!");
				return;
			}

			if (rbVolumeUnico.Checked)
			{
				if (txtVolumeUnico.Text.Trim().Length == 0)
				{
					aviso("Informe o nº do volume!!");
					txtVolumeUnico.Focus();
					return;
				}

				if (Global.converteInteiro(txtVolumeUnico.Text) <= 0)
				{
					aviso("Informe um nº de volume válido!!");
					txtVolumeUnico.Focus();
					return;
				}
			}

			if (rbVolumeRange.Checked)
			{
				if ((txtIntervaloInicio.Text.Trim().Length == 0) && (txtIntervaloFim.Text.Trim().Length == 0))
				{
					aviso("Informe o intervalo de volumes para impressão das etiquetas!!");
					txtIntervaloInicio.Focus();
					return;
				}

				if (txtIntervaloInicio.Text.Trim().Length > 0)
				{
					if (Global.converteInteiro(txtIntervaloInicio.Text) <= 0)
					{
						aviso("Informe um nº válido para o início do intervalo!!");
						txtIntervaloInicio.Focus();
						return;
					}
				}

				if (txtIntervaloFim.Text.Trim().Length > 0)
				{
					if (Global.converteInteiro(txtIntervaloFim.Text) <= 0)
					{
						aviso("Informe um nº válido para o final do intervalo!!");
						txtIntervaloFim.Focus();
						return;
					}
				}

				if (txtIntervaloFim.Text.Trim().Length > 0)
				{
					if (Global.converteInteiro(txtIntervaloFim.Text) < Global.converteInteiro(txtIntervaloInicio.Text))
					{
						aviso("O valor final do intervalo não pode ser menor que o valor inicial!!");
						txtIntervaloInicio.Focus();
						return;
					}
				}
			}
			#endregion

			#region [ Consiste nº NF ]
			numNfAux = (int)Global.converteInteiro(txtNF.Text);
			for (int i = 0; i < _listaEtqParcial.Count; i++)
			{
				if (_listaEtqParcial[i].obs_3.Length > 0)
				{
					strCampoObs = _listaEtqParcial[i].obs_3;
				}
				else
				{
					strCampoObs = _listaEtqParcial[i].obs_2;
				}

				if ((int)Global.converteInteiro(strCampoObs) == numNfAux)
				{
					blnNfOk = true;
					if (rbVolumeUnico.Checked)
					{
						if ((int)Global.converteInteiro(txtVolumeUnico.Text) <= _listaEtqParcial[i].qtde_volumes_pedido)
						{
							blnNumVolOk = true;
							break;
						}
					}
					else if (rbVolumeRange.Checked)
					{
						if ((int)Global.converteInteiro(txtIntervaloFim.Text) <= _listaEtqParcial[i].qtde_volumes_pedido)
						{
							blnNumVolOk = true;
							break;
						}
					}
				}
			}

			if (!blnNfOk)
			{
				aviso("O nº da NF informado não foi localizado nos dados selecionados!!");
				txtNF.Focus();
				return;
			}

			if (!blnNumVolOk)
			{
				if (rbVolumeUnico.Checked)
				{
					aviso("O nº do volume informado não existe nos dados selecionados!!");
					txtVolumeUnico.Focus();
				}
				else if (rbVolumeRange.Checked)
				{
					aviso("O nº final do intervalo de volumes não existe nos dados selecionados!!");
					txtIntervaloFim.Focus();
				}
				return;
			}
			#endregion

			#region [ Obtém os dados de resposta ]
			_numNF = (int)Global.converteInteiro(txtNF.Text);

			if (rbVolumeUnico.Checked)
			{
				_opcaoSelecionada = eOpcaoSelecao.VOLUME_UNICO;
				_numVolumeUnico = (int)Global.converteInteiro(txtVolumeUnico.Text);
			}

			if (rbVolumeRange.Checked)
			{
				_opcaoSelecionada = eOpcaoSelecao.VOLUME_INTERVALO;
				_numVolumeIntervaloInicio = (int)Global.converteInteiro(txtIntervaloInicio.Text);
				_numVolumeIntervaloFim = (int)Global.converteInteiro(txtIntervaloFim.Text);
			}
			#endregion

			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FVolumeSeleciona ]

		#region [ FVolumeSeleciona_Load ]
		private void FVolumeSeleciona_Load(object sender, EventArgs e)
		{
			#region [ Declarações ]
			string strValorAnterior = "--XXX--XXX--";
			string strCampoObs;
			int intCounter = 0;
			#endregion

			btnDummy.Top = -500;
			txtNF.Text = "";
			txtVolumeUnico.Text = "";
			txtIntervaloInicio.Text = "";
			txtIntervaloFim.Text = "";

			for (int i = 0; i < _listaEtqParcial.Count; i++)
			{
				if (_listaEtqParcial[i].obs_3.Length > 0)
				{
					strCampoObs = _listaEtqParcial[i].obs_3;
				}
				else
				{
					strCampoObs = _listaEtqParcial[i].obs_2;
				}

				if (!strCampoObs.Equals(strValorAnterior))
				{
					intCounter++;
					strValorAnterior = strCampoObs.ToString();
				}
			}

			if (intCounter == 1) txtNF.Text = strValorAnterior;
		}
		#endregion

		#region [ FVolumeSeleciona_Shown ]
		private void FVolumeSeleciona_Shown(object sender, EventArgs e)
		{
			if (txtNF.Text.Length > 0)
			{
				btnDummy.Focus();
			}
			else
			{
				txtNF.Focus();
			}
		}
		#endregion

		#endregion

		#region [ btnOk ]

		#region [ btnOk_Click ]
		private void btnOk_Click(object sender, EventArgs e)
		{
			ProcessaBotaoOk();
		}
		#endregion

		#endregion

		#region [ btnCancela ]

		#region [ btnCancela_Click ]
		private void btnCancela_Click(object sender, EventArgs e)
		{
			this.DialogResult = DialogResult.Cancel;
		}
		#endregion

		#endregion

		#region [ txtNF ]

		#region [ txtNF_Enter ]
		private void txtNF_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNF_Leave ]
		private void txtNF_Leave(object sender, EventArgs e)
		{
			txtNF.Text = txtNF.Text.Trim();
		}
		#endregion

		#region [ txtNF_KeyPress ]
		private void txtNF_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#region [ txtNF_KeyDown ]
		private void txtNF_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtVolumeUnico);
		}
		#endregion

		#endregion

		#region [ txtVolumeUnico ]

		#region [ txtVolumeUnico_Enter ]
		private void txtVolumeUnico_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtVolumeUnico_Leave ]
		private void txtVolumeUnico_Leave(object sender, EventArgs e)
		{
			txtVolumeUnico.Text = txtVolumeUnico.Text.Trim();
		}
		#endregion

		#region [ txtVolumeUnico_KeyPress ]
		private void txtVolumeUnico_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
			if (!rbVolumeUnico.Checked) rbVolumeUnico.Checked = true;
		}
		#endregion

		#region [ txtVolumeUnico_KeyDown ]
		private void txtVolumeUnico_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnOk);
		}
		#endregion

		#endregion

		#region [ txtIntervaloInicio ]
		
		#region [ txtIntervaloInicio_Enter ]
		private void txtIntervaloInicio_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtIntervaloInicio_Leave ]
		private void txtIntervaloInicio_Leave(object sender, EventArgs e)
		{
			txtIntervaloInicio.Text = txtIntervaloInicio.Text.Trim();
		}
		#endregion

		#region [ txtIntervaloInicio_KeyPress ]
		private void txtIntervaloInicio_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
			if (!rbVolumeRange.Checked) rbVolumeRange.Checked = true;
		}
		#endregion

		#region [ txtIntervaloInicio_KeyDown ]
		private void txtIntervaloInicio_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtIntervaloFim);
		}
		#endregion

		#endregion

		#region [ txtIntervaloFim ]
		
		#region [ txtIntervaloFim_Enter ]
		private void txtIntervaloFim_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtIntervaloFim_Leave ]
		private void txtIntervaloFim_Leave(object sender, EventArgs e)
		{
			txtIntervaloFim.Text = txtIntervaloFim.Text.Trim();
		}
		#endregion

		#region [ txtIntervaloFim_KeyPress ]
		private void txtIntervaloFim_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
			if (!rbVolumeRange.Checked) rbVolumeRange.Checked = true;
		}
		#endregion

		#region [ txtIntervaloFim_KeyDown ]
		private void txtIntervaloFim_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnOk);
		}
		#endregion

		#endregion

		#endregion
	}
}
