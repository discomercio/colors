#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
#endregion

namespace Financeiro
{
	public partial class FCepPesquisa : Financeiro.FModelo
	{
		#region [ Atributos ]
		private bool _InicializacaoOk;
		private bool _blnCarregandoComboUf;
		private String _ultUfConsultadaPesquisaLocalidades = "";
		#endregion

		#region [ Getters/Setters ]
		private String _cepDefault = "";
		public String cepDefault
		{
			get { return _cepDefault; }
			set { _cepDefault = Global.digitos(value); }
		}

		private String _cepSelecionado = "";
		public String cepSelecionado
		{
			get { return _cepSelecionado; }
			set { _cepSelecionado = value; }
		}

		private String _ufSelecionado = "";
		public String ufSelecionado
		{
			get { return _ufSelecionado; }
			set { _ufSelecionado = value; }
		}

		private String _cidadeSelecionada = "";
		public String cidadeSelecionada
		{
			get { return _cidadeSelecionada; }
			set { _cidadeSelecionada = value; }
		}

		private String _bairroSelecionado = "";
		public String bairroSelecionado
		{
			get { return _bairroSelecionado; }
			set { _bairroSelecionado = value; }
		}

		private String _logradouroSelecionado = "";
		public String logradouroSelecionado
		{
			get { return _logradouroSelecionado; }
			set { _logradouroSelecionado = value; }
		}

		private String _complementoSelecionado = "";
		public String complementoSelecionado
		{
			get { return _complementoSelecionado; }
			set { _complementoSelecionado = value; }
		}

		private String _numeroOuComplementoSelecionado = "";
		public String numeroOuComplementoSelecionado
		{
			get { return _numeroOuComplementoSelecionado; }
			set { _numeroOuComplementoSelecionado = value; }
		}
		#endregion

		#region [ Construtor ]
		public FCepPesquisa()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos ]

		#region [ trataAlteracaoUf ]
		private void trataAlteracaoUf()
		{
			String strUfSelecionada = "";

			if (_blnCarregandoComboUf) return;

			if (cbUF.SelectedIndex > -1) strUfSelecionada = cbUF.Items[cbUF.SelectedIndex].ToString();
			if (_ultUfConsultadaPesquisaLocalidades.Equals(strUfSelecionada)) return;

			_ultUfConsultadaPesquisaLocalidades = strUfSelecionada;

			cbLocalidade.DataSource = null;
			if (cbUF.SelectedIndex == -1) return;

			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");
				cbLocalidade.DataSource = CepDAO.getLocalidades(cbUF.Items[cbUF.SelectedIndex].ToString());
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoPesquisaPorCep ]
		private void trataBotaoPesquisaPorCep()
		{
			#region [ Declarações ]
			List<Cep> listaCep;
			Cep cep;
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

			#region [ Consistência ]
			if (txtCep.Text.Length == 0)
			{
				avisoErro("Informe o CEP a ser pesquisado!!");
				txtCep.Focus();
				return;
			}

			if (!Global.isCepOk(txtCep.Text))
			{
				avisoErro("CEP em formato inválido!!");
				txtCep.Focus();
				return;
			}
			#endregion

			#region [ Limpa o grid ]
			grdResultado.Rows.Clear();
			#endregion

			#region [ Executa consulta no BD ]
			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");
				listaCep = CepDAO.getCep(txtCep.Text);
			}
			catch (Exception ex)
			{
				avisoErro(ex.Message);
				return;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
			#endregion

			#region [ Processa o resultado ]
			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "carregando resultado no grid");

				if (listaCep.Count > 0) grdResultado.Rows.Add(listaCep.Count);

				#region [ Carrega os dados no grid ]
				for (int i = 0; i < listaCep.Count; i++)
				{
					cep = listaCep[i];
					grdResultado.Rows[i].Cells["cep"].Value = Global.formataCep(cep.cep);
					grdResultado.Rows[i].Cells["uf"].Value = Global.formataCep(cep.uf);
					grdResultado.Rows[i].Cells["cidade"].Value = Global.formataCep(cep.cidade);
					grdResultado.Rows[i].Cells["bairro"].Value = Global.formataCep(cep.bairro);
					grdResultado.Rows[i].Cells["logradouro"].Value = Global.formataCep(cep.logradouro);
					grdResultado.Rows[i].Cells["complemento"].Value = Global.formataCep(cep.complemento);
				}
				#endregion

				#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
				for (int i = 0; i < grdResultado.Rows.Count; i++)
				{
					if (grdResultado.Rows[i].Selected) grdResultado.Rows[i].Selected = false;
				}
				#endregion

				grdResultado.Focus();
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
			#endregion
		}
		#endregion

		#region [ trataBotaoPesquisaPorEndereco ]
		private void trataBotaoPesquisaPorEndereco()
		{
			#region [ Declarações ]
			String strUf, strLocalidade, strEndereco;
			List<Cep> listaCep;
			Cep cep;
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

			#region [ Consistências ]
			if (cbUF.SelectedIndex == -1)
			{
				avisoErro("É necessário selecionar a UF!!");
				cbUF.Focus();
				return;
			}
			if (cbLocalidade.SelectedIndex == -1)
			{
				avisoErro("É necessário selecionar a localidade!!");
				cbLocalidade.Focus();
				return;
			}
			#endregion

			#region [ Limpa o grid ]
			grdResultado.Rows.Clear();
			#endregion

			#region [ Executa consulta no BD ]
			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");
				strUf = cbUF.Items[cbUF.SelectedIndex].ToString();
				strLocalidade = cbLocalidade.Items[cbLocalidade.SelectedIndex].ToString();
				strEndereco = txtEndereco.Text.Trim();
				listaCep = CepDAO.getCep(strUf, strLocalidade, strEndereco);
			}
			catch (Exception ex)
			{
				avisoErro(ex.Message);
				return;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
			#endregion

			#region [ Processa o resultado ]
			try
			{
				info(ModoExibicaoMensagemRodape.EmExecucao, "carregando resultado no grid");

				if (listaCep.Count > 0) grdResultado.Rows.Add(listaCep.Count);

				#region [ Carrega os dados no grid ]
				for (int i = 0; i < listaCep.Count; i++)
				{
					cep = listaCep[i];
					grdResultado.Rows[i].Cells["cep"].Value = Global.formataCep(cep.cep);
					grdResultado.Rows[i].Cells["uf"].Value = Global.formataCep(cep.uf);
					grdResultado.Rows[i].Cells["cidade"].Value = Global.formataCep(cep.cidade);
					grdResultado.Rows[i].Cells["bairro"].Value = Global.formataCep(cep.bairro);
					grdResultado.Rows[i].Cells["logradouro"].Value = Global.formataCep(cep.logradouro);
					grdResultado.Rows[i].Cells["complemento"].Value = Global.formataCep(cep.complemento);
				}
				#endregion

				#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
				for (int i = 0; i < grdResultado.Rows.Count; i++)
				{
					if (grdResultado.Rows[i].Selected) grdResultado.Rows[i].Selected = false;
				}
				#endregion

				grdResultado.Focus();
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
			#endregion
		}
		#endregion

		#region [ trataBotaoResultadoCancela ]
		private void trataBotaoResultadoCancela()
		{
			this.DialogResult = DialogResult.Cancel;
		}
		#endregion

		#region [ trataBotaoResultadoConfirma ]
		private void trataBotaoResultadoConfirma()
		{
			int intLinhaSelecionada = -1;

			#region [ Consistência ]
			for (int i = 0; i < grdResultado.Rows.Count; i++)
			{
				if (grdResultado.Rows[i].Selected)
				{
					intLinhaSelecionada = i;
					break;
				}
			}

			if (intLinhaSelecionada == -1)
			{
				avisoErro("Nenhum CEP do resultado da pesquisa foi selecionado!!");
				grdResultado.Focus();
				return;
			}
			#endregion

			#region [ Preencheu o nº/complemento? ]
			if (txtNumeroOuComplemento.Text.Trim().Length == 0)
			{
				if (grdResultado.Rows[intLinhaSelecionada].Cells["logradouro"].Value.ToString().Length > 0)
				{
					if (!confirma("O campo Nº/Complemento não foi preenchido!!\nNão se esqueça de completar o endereço corretamente antes de cadastrar o boleto!!\n\nContinua?"))
					{
						txtNumeroOuComplemento.Focus();
						return;
					}
				}
			}
			#endregion

			#region [ Atualiza dados de retorno ]
			this.cepSelecionado = grdResultado.Rows[intLinhaSelecionada].Cells["cep"].Value.ToString();
			this.ufSelecionado = grdResultado.Rows[intLinhaSelecionada].Cells["uf"].Value.ToString();
			this.cidadeSelecionada = grdResultado.Rows[intLinhaSelecionada].Cells["cidade"].Value.ToString();
			this.bairroSelecionado = grdResultado.Rows[intLinhaSelecionada].Cells["bairro"].Value.ToString();
			this.logradouroSelecionado = grdResultado.Rows[intLinhaSelecionada].Cells["logradouro"].Value.ToString();
			this.complementoSelecionado = grdResultado.Rows[intLinhaSelecionada].Cells["complemento"].Value.ToString();
			this.numeroOuComplementoSelecionado = txtNumeroOuComplemento.Text;
			#endregion

			this.DialogResult = DialogResult.OK;
		}
		#endregion

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtCep.Text = "";
			cbUF.SelectedIndex = -1;
			cbLocalidade.Items.Clear();
			txtEndereco.Text = "";
			grdResultado.Rows.Clear();
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FCepPesquisa ]

		#region [ FCepPesquisa_Load ]
		private void FCepPesquisa_Load(object sender, EventArgs e)
		{
			String strListaUf = "AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO";
			String[] vUf;
			bool blnSucesso = false;

			try
			{
				limpaCampos();

				#region [ Combo UF ]
				_blnCarregandoComboUf = true;
				vUf = strListaUf.Split(' ');
				foreach (String uf in vUf)
				{
					cbUF.Items.Add(uf);
				}
				cbUF.SelectedIndex = -1;
				_blnCarregandoComboUf = false;
				#endregion

				#region [ Cep default ]
				if (_cepDefault.Length > 0)
				{
					txtCep.Text = Global.formataCep(_cepDefault);
				}
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

		#region [ FCepPesquisa_Shown ]
		private void FCepPesquisa_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					txtCep.Focus();

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

		#region [ txtCep ]

		#region [ txtCep_Enter ]
		private void txtCep_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtCep_Leave ]
		private void txtCep_Leave(object sender, EventArgs e)
		{
			if (txtCep.Text.Length == 0) return;

			if (!Global.isCepOk(txtCep.Text))
			{
				avisoErro("CEP inválido!!");
				txtCep.Focus();
				return;
			}
		}
		#endregion

		#region [ txtCep_KeyDown ]
		private void txtCep_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				if (txtCep.Text.Length == 0)
				{
					cbUF.Focus();
					return;
				}
				if (Global.isCepOk(txtCep.Text)) trataBotaoPesquisaPorCep();
				return;
			}
		}
		#endregion

		#region [ txtCep_KeyPress ]
		private void txtCep_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoCep(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ cbUF ]

		#region [ cbUF_KeyDown ]
		private void cbUF_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, cbLocalidade);
		}
		#endregion

		#region [ cbUF_SelectionChangeCommitted ]
		private void cbUF_SelectionChangeCommitted(object sender, EventArgs e)
		{
			trataAlteracaoUf();
		}
		#endregion

		#region [ cbUF_Leave ]
		private void cbUF_Leave(object sender, EventArgs e)
		{
			trataAlteracaoUf();
		}
		#endregion

		#endregion

		#region [ cbLocalidade ]

		#region [ cbLocalidade_KeyDown ]
		private void cbLocalidade_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtEndereco);
		}
		#endregion

		#endregion

		#region [ txtEndereco ]

		#region [ txtEndereco_Enter ]
		private void txtEndereco_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtEndereco_Leave ]
		private void txtEndereco_Leave(object sender, EventArgs e)
		{
			txtEndereco.Text = txtEndereco.Text.Trim();
		}
		#endregion

		#region [ txtEndereco_KeyDown ]
		private void txtEndereco_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				
				if ((cbUF.SelectedIndex == -1) || (cbLocalidade.SelectedIndex == -1))
				{
					btnPesquisarPorEndereco.Focus();
					return;
				}

				if (txtEndereco.Text.Trim().Length > 0)
					trataBotaoPesquisaPorEndereco();
				else
					btnPesquisarPorEndereco.Focus();

				return;
			}
		}
		#endregion

		#region [ txtEndereco_KeyPress ]
		private void txtEndereco_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtNumeroOuComplemento ]

		#region [ txtNumeroOuComplemento_Enter ]
		private void txtNumeroOuComplemento_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNumeroOuComplemento_Leave ]
		private void txtNumeroOuComplemento_Leave(object sender, EventArgs e)
		{
			txtNumeroOuComplemento.Text = txtNumeroOuComplemento.Text.Trim();
		}
		#endregion

		#region [ txtNumeroOuComplemento_KeyDown ]
		private void txtNumeroOuComplemento_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnConfirma);
		}
		#endregion

		#region [ txtNumeroOuComplemento_KeyPress ]
		private void txtNumeroOuComplemento_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ btnPesquisaPorCep ]

		#region [ btnPesquisaPorCep_Click ]
		private void btnPesquisaPorCep_Click(object sender, EventArgs e)
		{
			trataBotaoPesquisaPorCep();
		}
		#endregion

		#endregion

		#region [ btnPesquisarPorEndereco ]

		#region [ btnPesquisarPorEndereco_Click ]
		private void btnPesquisarPorEndereco_Click(object sender, EventArgs e)
		{
			trataBotaoPesquisaPorEndereco();
		}
		#endregion

		#endregion

		#region [ btnCancela ]

		#region [ btnCancela_Click ]
		private void btnCancela_Click(object sender, EventArgs e)
		{
			trataBotaoResultadoCancela();
		}
		#endregion

		#endregion

		#region [ btnConfirma ]

		#region [ btnConfirma_Click ]
		private void btnConfirma_Click(object sender, EventArgs e)
		{
			trataBotaoResultadoConfirma();
		}
		#endregion

		#endregion

		#endregion
	}
}
