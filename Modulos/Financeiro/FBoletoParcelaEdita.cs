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
	public partial class FBoletoParcelaEdita : Financeiro.FModelo
	{
		#region [ Enum ]
		public enum eBoletoParcelaOperacao : byte
		{
			EDITAR = 1,
			INCLUIR = 2
		}
		#endregion

		#region [ Atributos ]
		private bool _InicializacaoOk;

		private eBoletoParcelaOperacao _operacaoSelecionada;

		private String _parcelaSelecionadaDadosRateio = "";
		public String parcelaSelecionadaDadosRateio
		{
			get { return _parcelaSelecionadaDadosRateio; }
			set { _parcelaSelecionadaDadosRateio = value; }
		}

		private DateTime _parcelaSelecionadaDtVencto;
		public DateTime parcelaSelecionadaDtVencto
		{
			get { return _parcelaSelecionadaDtVencto; }
			set { _parcelaSelecionadaDtVencto = value; }
		}

		private decimal _parcelaSelecionadaValor = 0;
		public decimal parcelaSelecionadaValor
		{
			get { return _parcelaSelecionadaValor; }
			set { _parcelaSelecionadaValor = value; }
		}
		#endregion

		#region [ Construtor ]
		public FBoletoParcelaEdita(eBoletoParcelaOperacao operacao)
		{
			InitializeComponent();
			_operacaoSelecionada = operacao;
		}
		#endregion

		#region [ Métodos ]

		#region [ recalculaValorParcela ]
		private void recalculaValorParcela()
		{
			decimal vlParcela = 0;
			for (int i = 0; i < grdRateio.Rows.Count; i++)
			{
				vlParcela += Global.converteNumeroDecimal(grdRateio.Rows[i].Cells["grdRateio_valor"].Value.ToString());
			}
			lblTotalGridParcelas.Text = Global.formataMoeda(vlParcela);
		}
		#endregion

		#region [ trataBotaoResultadoConfirma ]
		private void trataBotaoResultadoConfirma()
		{
			#region [ Declarações ]
			String strDadosRateio = "";
			#endregion

			#region [ Recalcula o valor da parcela ]
			recalculaValorParcela();
			#endregion

			#region [ Consistência ]
			if (txtVencto.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar a data de vencimento da parcela!!");
				txtVencto.Focus();
				return;
			}

			if (!Global.isDataOk(txtVencto.Text))
			{
				avisoErro("Data inválida!!");
				txtVencto.Focus();
				return;
			}

			if (Global.converteNumeroDecimal(lblTotalGridParcelas.Text) <= 0)
			{
				avisoErro("Valor da parcela é inválido!!");
				return;
			}
			#endregion

			#region [ Monta dados do rateio ]
			for (int i = 0; i < grdRateio.Rows.Count; i++)
			{
				if (strDadosRateio.Length > 0) strDadosRateio += "|";
				strDadosRateio += grdRateio.Rows[i].Cells["grdRateio_pedido"].Value.ToString() + "=" + grdRateio.Rows[i].Cells["grdRateio_valor"].Value.ToString();
			}
			#endregion

			#region [ Atualiza dados de retorno ]
			_parcelaSelecionadaDtVencto = Global.converteDdMmYyyyParaDateTime(Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtVencto.Text));
			_parcelaSelecionadaValor = Global.converteNumeroDecimal(lblTotalGridParcelas.Text);
			_parcelaSelecionadaDadosRateio = strDadosRateio;
			#endregion

			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FBoletoParcelaEdita ]

		#region [ FBoletoParcelaEdita_Shown ]
		private void FBoletoParcelaEdita_Shown(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String[] vRateio;
			String[] v;
			#endregion

			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Preenche campos ]
					if (_operacaoSelecionada == eBoletoParcelaOperacao.INCLUIR)
					{
						grdRateio.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
						txtVencto.Text = "";
						lblTotalGridParcelas.Text = Global.formataMoeda(_parcelaSelecionadaValor);
						txtVencto.Focus();
					}
					else
					{
						grdRateio.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
						txtVencto.Text = Global.formataDataDdMmYyyyComSeparador(_parcelaSelecionadaDtVencto);
						lblTotalGridParcelas.Text = Global.formataMoeda(_parcelaSelecionadaValor);
						txtVencto.Focus();
					}
					#endregion

					#region [ Preenche o grid de rateio ]
					vRateio = parcelaSelecionadaDadosRateio.Split('|');
					foreach (String rateio in vRateio)
					{
						if (rateio != null)
						{
							if (rateio.Trim().Length > 0)
							{
								v = rateio.Split('=');
								grdRateio.Rows.Add();
								grdRateio.Rows[grdRateio.Rows.Count - 1].Cells["grdRateio_pedido"].Value = v[0];
								grdRateio.Rows[grdRateio.Rows.Count - 1].Cells["grdRateio_valor"].Value = v[1];
							}
						}
					}
					#endregion

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < grdRateio.Rows.Count; i++)
					{
						if (grdRateio.Rows[i].Selected) grdRateio.Rows[i].Selected = false;
						for (int j = 0; j < grdRateio.Columns.Count; j++)
						{
							if (grdRateio[j, i].Selected) grdRateio[j, i].Selected = false;
						}
					}
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

		#region [ FBoletoParcelaEdita_KeyPress ]
		private void FBoletoParcelaEdita_KeyPress(object sender, KeyPressEventArgs e)
		{
			#region [ Quando o grid estiver em edição, filtra digitação ]
			if (this.ActiveControl.GetType().Equals(typeof(DataGridViewTextBoxEditingControl)))
			{
				if (grdRateio.CurrentCell != null)
				{
					if (grdRateio.IsCurrentCellInEditMode)
					{
						e.KeyChar = Global.filtraDigitacaoMoeda(e.KeyChar);
					}
				}
			}
			#endregion
		}
		#endregion

		#region [ FBoletoParcelaEdita_KeyDown ]
		private void FBoletoParcelaEdita_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				#region [ Quando terminar de editar a última célula, coloca o foco no botão de cadastrar ]
				if (grdRateio.CurrentCell != null)
				{
					if (grdRateio.CurrentCell.RowIndex == (grdRateio.Rows.Count - 1))
					{
						e.SuppressKeyPress = true;
						if (txtVencto.Text.Trim().Length == 0)
						{
							txtVencto.Focus();
							return;
						}
						trataBotaoResultadoConfirma();
						return;
					}
				}
				#endregion
			}
		}
		#endregion

		#endregion

		#region [ grdRateio ]

		#region [ grdRateio_CellEndEdit ]
		private void grdRateio_CellEndEdit(object sender, DataGridViewCellEventArgs e)
		{
			if (grdRateio.CurrentCell == null) return;
			grdRateio.CurrentCell.Value = Global.formataMoeda(Global.converteNumeroDecimal(grdRateio.CurrentCell.Value.ToString()));
			recalculaValorParcela();
		}
		#endregion

		#endregion

		#region [ txtVencto ]

		#region [ txtVencto_Enter ]
		private void txtVencto_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtVencto_Leave ]
		private void txtVencto_Leave(object sender, EventArgs e)
		{
			if (txtVencto.Text.Length == 0) return;
			txtVencto.Text = Global.formataDataDigitadaParaDDMMYYYYComSeparador(txtVencto.Text);
			if (!Global.isDataOk(txtVencto.Text))
			{
				avisoErro("Data inválida!!");
				txtVencto.Focus();
				return;
			}
		}
		#endregion

		#region [ txtVencto_KeyDown ]
		private void txtVencto_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				if (grdRateio.Rows.Count > 0)
				{
					grdRateio.Focus();
					grdRateio.Rows[0].Cells["grdRateio_valor"].Selected = true;
					if (_operacaoSelecionada == eBoletoParcelaOperacao.INCLUIR) grdRateio.BeginEdit(true);
				}
				return;
			}
		}
		#endregion

		#region [ txtVencto_KeyPress ]
		private void txtVencto_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ pnCampos ]

		#region [ pnCampos_Click ]
		private void pnCampos_Click(object sender, EventArgs e)
		{
			btnDummy.Focus();
		}
		#endregion

		#endregion

		#region [ btnCadastrar ]

		#region [ btnCadastrar_Click ]
		private void btnCadastrar_Click(object sender, EventArgs e)
		{
			trataBotaoResultadoConfirma();
		}
		#endregion

		#endregion

		#endregion
	}
}
