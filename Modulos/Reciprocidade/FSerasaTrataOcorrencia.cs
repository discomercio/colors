#region [ using ]
using System;
using System.Collections.Generic;
using System.Windows.Forms;
#endregion

namespace Reciprocidade
{
	public partial class FSerasaTrataOcorrencia : Form
	{
		#region [ Declarações ]
		private bool _InicializacaoOk;
		private int _id = 0;
		private String _numBoleto;
		private DateTime _dtEmissao;
		private DateTime _dtVencto;
		private Decimal _vlTitulo;
		private DateTime _dtPagto;
		private Decimal _vlPago;
		private String _linhaCodigosErros;
		private Dictionary<String, String> _dictErros;

		public String _numBoletoCorrigido;
		public DateTime _dtEmissaoCorrigido;
		public DateTime _dtVenctoCorrigido;
		public Decimal _vlTituloCorrigido;
		public DateTime _dtPagtoCorrigido;
		public Decimal _vlPagoCorrigido;
		#endregion

		#region [ construtor ]
		public FSerasaTrataOcorrencia(int id,
									  String numBoleto,
									  DateTime dtEmissao,
									  DateTime dtVencto,
									  Decimal vlTitulo,
									  DateTime dtPagto,
									  Decimal vlPago,
									  String linhaCodigosErros,
									  Dictionary<String, String> dictErros)
		{
			InitializeComponent();

			#region [ Define a cor de fundo de acordo com o ambiente acessado ]
			BackColor = Global.BackColorPainelPadrao;
			#endregion

			this._id = id;
			this._numBoleto = numBoleto;
			this._dtEmissao = dtEmissao;
			this._dtVencto = dtVencto;
			this._vlTitulo = vlTitulo;
			this._dtPagto = dtPagto;
			this._vlPago = vlPago;
			this._linhaCodigosErros = linhaCodigosErros;
			this._dictErros = dictErros;
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

		#region[ confirma ]
		public bool confirma(string mensagem)
		{
			return (MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes);
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Numero Titulo ]
			if (mskNumTitulo.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o número do título!!");
				mskNumTitulo.Focus();
				return false;
			}

			if (mskNumTitulo.Text.Trim().Length != 12)
			{
				avisoErro("O campo Número do Título deve conter 12 posições!!");
				mskNumTitulo.Focus();
				return false;
			}
			#endregion

			#region [ Valor do Título ]
			if (txtValor.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o valor do título!!");
				txtValor.Focus();
				return false;
			}

			if (Global.converteNumeroDecimal(txtValor.Text.Trim()) <= 0)
			{
				avisoErro("O valor do título informado está incorreto!!");
				txtValor.Focus();
				return false;
			}
			#endregion

			#region [ Data de Vencimento ]
			if (DtpDataVecimento.Value < DtpDataEmissao.Value)
			{
				avisoErro("A data de vencimento deve ser maior ou\n igual a data de emissão!!");
				DtpDataVecimento.Focus();
				return false;
			}
			#endregion

			#region [ Data de Pagamento ]
			if (DtpDataPagamento.Checked &&
					DtpDataPagamento.Value < DtpDataEmissao.Value)
			{
				avisoErro("A data de pagamento deve ser maior ou\n igual a data de emissão!!");
				DtpDataPagamento.Focus();
				return false;
			}

			if (DtpDataPagamento.Checked &&
					txtValorPago.Text.Trim().Length == 0)
			{
				avisoErro("O campo valor pago deve ser preenchido!!");
				txtValorPago.Focus();
				return false;
			}

			if (DtpDataPagamento.Checked &&
					Global.converteNumeroDecimal(txtValorPago.Text.Trim()) == 0)
			{
				avisoErro("O campo valor pago deve ser preenchido!!");
				txtValorPago.Focus();
				return false;
			}
			#endregion

			#region [ Valor Pago ]
			if (txtValorPago.Text.Trim().Length > 0 &&
					Global.converteNumeroDecimal(txtValorPago.Text.Trim()) < 0)
			{
				avisoErro("O valor pago informado está incorreto!!");
				txtValorPago.Focus();
				return false;
			}

			if (Global.converteNumeroDecimal(txtValorPago.Text.Trim()) > 0 &&
					!DtpDataPagamento.Checked)
			{
				avisoErro("É necessário informar a data de pagamento!!");
				DtpDataPagamento.Focus();
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ Decodifica ]
		private List<String> decodifica(String linhaCodigoErros)
		{
			linhaCodigoErros = linhaCodigoErros.Trim();
			List<String> mensagens = new List<String>();

			for (int i = 0; i < linhaCodigoErros.Length; i = i + 3)
			{
				String cod = linhaCodigoErros.Substring(i, 3);
				String msg = _dictErros[cod];
				mensagens.Add(msg);
			}

			return mensagens;
		}
		#endregion

		#region [ trataBotaoCancela ]
		private void trataBotaoCancela()
		{
			this.DialogResult = DialogResult.Cancel;
		}
		#endregion

		#region [ trataBotaoOk ]
		private void trataBotaoOk()
		{
			if (!consisteCampos()) return;
			if (!confirma("Confirma o tratamento desse título?")) return;

			_numBoletoCorrigido = mskNumTitulo.Text.Trim();
			_dtEmissaoCorrigido = DtpDataEmissao.Value;
			_dtVenctoCorrigido = DtpDataVecimento.Value;
			_vlTituloCorrigido = Global.converteNumeroDecimal(txtValor.Text.Trim());

			if (DtpDataPagamento.Checked)
			{
				_dtPagtoCorrigido = DtpDataPagamento.Value;
			}
			else
			{
				_dtPagtoCorrigido = DateTime.MinValue;
			}

			_vlPagoCorrigido = Global.converteNumeroDecimal(txtValorPago.Text.Trim());
			this.DialogResult = DialogResult.OK;
		}
		#endregion
		#endregion

		#region [ FSerasaTrataOcorrencia_Shown ]
		private void FSerasaTrataOcorrencia_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Preenche dados ]
					mskNumTitulo.Text = this._numBoleto;
					txtValor.Text = Global.formataMoeda(this._vlTitulo);
					txtValorPago.Text = Global.formataMoeda(this._vlPago);
					DtpDataEmissao.Value = this._dtEmissao;
					DtpDataVecimento.Value = this._dtVencto;

					if (this._dtPagto == DateTime.MinValue)
					{
						DtpDataPagamento.Checked = false;
					}
					else
					{
						DtpDataPagamento.Value = this._dtPagto;
					}

					#region [ Preenche Listbox de Erros ]
					List<String> erros = decodifica(this._linhaCodigosErros);
					foreach (String erro in erros)
					{
						LstErros.Items.Add(erro);
					}
					#endregion
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

		#region [ btnOk_Click ]
		private void btnOk_Click(object sender, EventArgs e)
		{
			trataBotaoOk();
		}
		#endregion

		#region [ btnCancela_Click ]
		private void btnCancela_Click(object sender, EventArgs e)
		{
			trataBotaoCancela();
		}
		#endregion
	}
}
