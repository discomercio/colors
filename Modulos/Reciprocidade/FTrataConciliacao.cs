#region [ using ]
using System;
using System.Windows.Forms;
#endregion

namespace Reciprocidade
{
	public partial class FTrataConciliacao : Form
	{
		#region [ Declarações ]
		private bool _InicializacaoOk;
		private int _id = 0;
		private String _numBoleto;
		private DateTime _dtEmissao;
		private DateTime _dtVencto;
		private Decimal _vlTitulo;
		private DateTime _dtPagto;
		private DateTime _dtFinalPeriodoArquivo;

		public DateTime _dtVenctoCorrigido;
		public Decimal _vlTituloCorrigido;
		public DateTime _dtPagtoCorrigido;
		public bool _blnTituloExcluido;
		#endregion

		#region [ Construtor ]
		public FTrataConciliacao(int id,
									  String numBoleto,
									  DateTime dtEmissao,
									  DateTime dtVencto,
									  Decimal vlTitulo,
									  DateTime dtPagto,
									  DateTime dtFinalPeriodoArquivo)
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
			this._dtFinalPeriodoArquivo = dtFinalPeriodoArquivo;
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
			if (DtpDataVecimento.Value < _dtEmissao)
			{
				avisoErro("A data de vencimento deve ser maior ou\n igual a data de emissão (" + Global.formataDataDdMmYyyyComSeparador(_dtEmissao) + ")!!");
				DtpDataVecimento.Focus();
				return false;
			}
			#endregion

			#region [ Data de Pagamento ]
			if (DtpDataPagamento.Checked && (DtpDataPagamento.Value < _dtEmissao))
			{
				avisoErro("A data de pagamento deve ser maior ou\n igual a data de emissão (" + Global.formataDataDdMmYyyyComSeparador(_dtEmissao) + ")!!");
				DtpDataPagamento.Focus();
				return false;
			}

			if (DtpDataPagamento.Checked && (DtpDataPagamento.Value > _dtFinalPeriodoArquivo))
			{
				avisoErro("A data de pagamento não pode ser posterior à data final do período informado no header do arquivo (" + Global.formataDataDdMmYyyyComSeparador(_dtFinalPeriodoArquivo) + ")!!");
				DtpDataPagamento.Focus();
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ trataBotaoCancela ]
		private void trataBotaoCancela()
		{
			this.DialogResult = DialogResult.Cancel;
		}
		#endregion

		#region [ trataBotaoLimpaDtPagto ]
		private void trataBotaoLimpaDtPagto()
		{
			DtpDataPagamento.Checked = false;
		}
		#endregion

		#region [ trataBotaoOk ]
		private void trataBotaoOk()
		{
			if (!consisteCampos()) return;

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

			if (chkExclusaoTitulo.Checked)
			{
				_blnTituloExcluido = true;
			}

			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FTrataConciliacao_Shown ]
		private void FTrataConciliacao_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Preenche dados ]
					txtNumTitulo.Text = this._numBoleto;
					txtDataEmissao.Text = Global.formataDataDdMmYyyyComSeparador(this._dtEmissao);
					txtValor.Text = Global.formataMoeda(this._vlTitulo);
					DtpDataVecimento.Value = this._dtVencto;

					// Obs: se a data atribuída a MaxDate for menor que hoje, o componente automaticamente seleciona a data do MaxDate e a propriedade checked fica 'true'
					DtpDataPagamento.MinDate = _dtEmissao;
					DtpDataPagamento.MaxDate = _dtFinalPeriodoArquivo;
					if (DtpDataPagamento.Checked) DtpDataPagamento.Checked = false;
					if (this._dtPagto != DateTime.MinValue)
					{
						try
						{
							DtpDataPagamento.Value = this._dtPagto;
							DtpDataPagamento.Checked = true;
						}
						catch (Exception)
						{
							// NOP
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

		#region [ btnLimpaDtPagto_Click ]
		private void btnLimpaDtPagto_Click(object sender, EventArgs e)
		{
			trataBotaoLimpaDtPagto();

		}
		#endregion

		#endregion
	}
}
