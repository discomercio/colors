#region [ using ]
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Media;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	#region [ Delegate ]
	public delegate void RegistroGravadoEventHandler();
	public delegate void RegistroAnuladoEventHandler();
	#endregion

	public partial class FBoletoCadastraDetalhe : Financeiro.FModelo
	{
		#region [ Eventos Customizados ]
		public event RegistroGravadoEventHandler evtRegistroGravado;
		public event RegistroAnuladoEventHandler evtRegistroAnulado;
		#endregion

		#region [ Atributos ]
		private Form _formChamador = null;

		private bool _InicializacaoOk;
		public bool inicializacaoOk
		{
			get { return _InicializacaoOk; }
		}

		private bool _OcorreuExceptionNaInicializacao;
		public bool ocorreuExceptionNaInicializacao
		{
			get { return _OcorreuExceptionNaInicializacao; }
		}

		private bool _blnRegistroFoiGravado = false;
		private bool _blnRegistroFoiAnulado = false;
		private List<Pedido> _listaPedidos = new List<Pedido>();

		private decimal _vlTotalFamiliaPagoPedidos = 0;
		private decimal _vlTotalDevolucoesPedidos = 0;
		private decimal _vlTotalPrecoNfPedidos = 0;
		private decimal _vlTotalBoletoPedidos = 0;
		private decimal _vlTotalFormaPagtoPedidos = 0;
		
		private int _idBoletoPreCadastradoSelecionado;
		public int idBoletoPreCadastradoSelecionado
		{
			get { return _idBoletoPreCadastradoSelecionado; }
			set { _idBoletoPreCadastradoSelecionado = value; }
		}

		BoletoPreCadastrado boletoPreCadastradoSelecionado;
		Cliente clienteSelecionado;
		BoletoCedente boletoCedenteSelecionado;

		ToolStripMenuItem menuBoleto;
		ToolStripMenuItem menuBoletoCadastrar;
		ToolStripMenuItem menuBoletoAnular;

		FCepPesquisa fCepPesquisa;
		FPedido fPedido;
		FBoletoParcelaEdita fBoletoParcelaEdita;
		#endregion

		#region [ Construtor ]
		public FBoletoCadastraDetalhe(Form formChamador)
		{
			InitializeComponent();

			_formChamador = formChamador;

			#region [ Menu Boleto ]
			// Menu principal de Boleto
			menuBoleto = new ToolStripMenuItem("&Boleto");
			menuBoleto.Name = "menuBoleto";
			// Cadastrar
			menuBoletoCadastrar = new ToolStripMenuItem("&Cadastrar", null, menuBoletoCadastrar_Click);
			menuBoletoCadastrar.Name = "menuBoletoCadastrar";
			menuBoleto.DropDownItems.Add(menuBoletoCadastrar);
			// Anular
			menuBoletoAnular = new ToolStripMenuItem("&Anular", null, menuBoletoAnular_Click);
			menuBoletoAnular.Name = "menuBoletoAnular";
			menuBoleto.DropDownItems.Add(menuBoletoAnular);
			// Adiciona o menu Boleto ao menu principal
			menuPrincipal.Items.Insert(1, menuBoleto);
			#endregion
		}
		#endregion

		#region [ Métodos ]

		#region [ atualizaExibicaoValorTotalPedidoSelecionado ]
		private void atualizaExibicaoValorTotalPedidoSelecionado()
		{
			String strPedidoSelecionado;

			if (lbPedido.Items.Count == 0)
			{
				lblTitTotalPedidoSelecionado.Text = "";
				lblTotalPedidoSelecionado.Text = "";
				return;
			}

			if (lbPedido.SelectedIndex == -1)
				strPedidoSelecionado = lbPedido.Items[0].ToString();
			else
				strPedidoSelecionado = lbPedido.Items[lbPedido.SelectedIndex].ToString();

			for (int i = 0; i < _listaPedidos.Count; i++)
			{
				if (_listaPedidos[i].pedido.Equals(strPedidoSelecionado))
				{
					lblTitTotalPedidoSelecionado.Text = "Pedido " + strPedidoSelecionado;
					lblTotalPedidoSelecionado.Text = Global.formataMoeda(_listaPedidos[i].vlTotalPrecoNfDestePedido);
					break;
				}
			}
		}
		#endregion

		#region [ ajustaPosicaoLblTotalGridParcelas ]
		private void ajustaPosicaoLblTotalGridParcelas()
		{
			lblTotalGridParcelas.Left = grdParcelas.Left + grdParcelas.Width - lblTotalGridParcelas.Width - 3;
			if (Global.isVScrollBarVisible(grdParcelas)) lblTotalGridParcelas.Left -= Global.getVScrollBarWidth(grdParcelas);
		}
		#endregion

		#region [ recalculaValorTotalParcelas ]
		private void recalculaValorTotalParcelas()
		{
			decimal vlTotal = 0;
			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				vlTotal += Global.converteNumeroDecimal(grdParcelas.Rows[i].Cells["grdParcelas_valor"].Value.ToString());
			}
			lblTotalGridParcelas.Text = Global.formataMoeda(vlTotal);
		}
		#endregion

		#region [ atualizaNumeracaoParcelas ]
		private void atualizaNumeracaoParcelas()
		{
			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				grdParcelas.Rows[i].Cells["grdParcelas_num_parcela"].Value = (i + 1).ToString() + " / " + grdParcelas.Rows.Count.ToString();
			}
		}
		#endregion

		#region [ ajustaPosicaoLblTotalGridParcelasNF ]
		private void ajustaPosicaoLblTotalGridParcelasNF()
		{
			lblTotalGridParcelasNF.Left = grdParcelasNF.Left + grdParcelasNF.Width - lblTotalGridParcelasNF.Width - 3;
			if (Global.isVScrollBarVisible(grdParcelasNF)) lblTotalGridParcelasNF.Left -= Global.getVScrollBarWidth(grdParcelasNF);
		}
		#endregion

		#region [ limpaCampos ]
		void limpaCampos()
		{
			cbBoletoCedente.SelectedIndex = -1;
			txtClienteNome.Text = "";
			txtEndereco.Text = "";
			txtClienteCnpjCpf.Text = "";
			txtBairro.Text = "";
			txtCep.Text = "";
			txtCidade.Text = "";
			txtUF.Text = "";
			txtEmail.Text = "";
			txtSegundaMensagem.Text = "";
			txtMensagem1.Text = "";
			txtMensagem2.Text = "";
			txtMensagem3.Text = "";
			txtMensagem4.Text = "";
			grdParcelas.Rows.Clear();
			lblTotalGridParcelasNF.Text = "";
			lblTotalGridParcelas.Text = "";
			lbPedido.Items.Clear();
			txtNumeroNF.Text = "";
			txtJurosMora.Text = "";
			txtPercMulta.Text = "";
			txtProtestarApos.Text = "";
			lblTitTotalPedidoSelecionado.Text = "";
			lblTotalPedidoSelecionado.Text = "";
			lblTotalTodosPedidos.Text = "";
			lblTotalValorPagoTodosPedidos.Text = "";
			lblTotalDevolucoesTodosPedidos.Text = "";
		}
		#endregion

		#region [ limpaCamposVinculadosBoletoCedente ]
		private void limpaCamposVinculadosBoletoCedente()
		{
			txtSegundaMensagem.Text = "";
			txtMensagem1.Text = "";
			txtMensagem2.Text = "";
			txtMensagem3.Text = "";
			txtMensagem4.Text = "";
			txtJurosMora.Text = "";
			txtPercMulta.Text = "";
			txtProtestarApos.Text = "";
		}
		#endregion

		#region [ preencheCamposVinculadosBoletoCedente ]
		private void preencheCamposVinculadosBoletoCedente()
		{
			txtSegundaMensagem.Text = boletoCedenteSelecionado.segunda_mensagem_padrao;
			txtMensagem1.Text = boletoCedenteSelecionado.mensagem_1_padrao;
			txtMensagem2.Text = boletoCedenteSelecionado.mensagem_2_padrao;
			txtMensagem3.Text = boletoCedenteSelecionado.mensagem_3_padrao;
			txtMensagem4.Text = boletoCedenteSelecionado.mensagem_4_padrao;
			txtJurosMora.Text = Global.formataPercentual(boletoCedenteSelecionado.juros_mora);
			txtPercMulta.Text = Global.formataPercentual(boletoCedenteSelecionado.perc_multa);
			txtProtestarApos.Text = boletoCedenteSelecionado.qtde_dias_protestar_apos_padrao.ToString();
		}
		#endregion

		#region [ trataSelecaoBoletoCedente ]
		private void trataSelecaoBoletoCedente()
		{
			if (cbBoletoCedente.SelectedIndex == -1)
			{
				boletoCedenteSelecionado = null;
				limpaCamposVinculadosBoletoCedente();
			}
			else
			{
				boletoCedenteSelecionado = BoletoCedenteDAO.getBoletoCedente((int)Global.converteInteiro(cbBoletoCedente.SelectedValue.ToString()));
				preencheCamposVinculadosBoletoCedente();
			}
		}
		#endregion

		#region [ comboBoletoCedentePosicionaDefault ]
		private bool comboBoletoCedentePosicionaDefault()
		{
			#region [ Declarações ]
			bool blnHaDefault = false;
			bool blnPedidosComCedentesDiferentes = false;
			int intIdBoletoCedente = 0;
			int intIdBoletoCedenteAux;
			NFeEmitente nfeEmitente;
			DsDataSource.DtbBoletoCedenteComboRow rowBoletoCedente;
			#endregion

			for (int i = 0; i < _listaPedidos.Count; i++)
			{
				intIdBoletoCedenteAux = 0;

				#region [ Estrutura nova: obtém o cedente do boleto através do emitente da NFe definido no pedido ]
				nfeEmitente = NFeEmitenteDAO.getNFeEmitenteById(_listaPedidos[i].id_nfe_emitente);
				if (nfeEmitente != null)
				{
					intIdBoletoCedenteAux = nfeEmitente.id_boleto_cedente;
				}
				#endregion

				if (intIdBoletoCedenteAux > 0)
				{
					if (intIdBoletoCedente == 0)
					{
						intIdBoletoCedente = intIdBoletoCedenteAux;
					}
					else
					{
						if (intIdBoletoCedente != intIdBoletoCedenteAux)
						{
							blnPedidosComCedentesDiferentes = true;
							break;
						}
					}
				}
			} // for

			if (blnPedidosComCedentesDiferentes) intIdBoletoCedente = 0;

			foreach (System.Data.DataRowView item in cbBoletoCedente.Items)
			{
				rowBoletoCedente = (DsDataSource.DtbBoletoCedenteComboRow)item.Row;
				if (rowBoletoCedente.id == intIdBoletoCedente)
				{
					cbBoletoCedente.SelectedIndex = cbBoletoCedente.Items.IndexOf(item);
					blnHaDefault = true;
					break;
				}
			}

			return blnHaDefault;
		}
		#endregion

		#region [ obtemDadosBoletoCamposTela ]
		/// <summary>
		/// Carrega os dados dos campos na tela em um objeto da classe Boleto
		/// </summary>
		/// <returns>
		/// Retorna um objeto Boleto com os dados dos campos da tela
		/// </returns>
		private Boleto obtemDadosBoletoCamposTela()
		{
			#region [ Declarações ]
			Boleto boletoEditado = new Boleto();
			BoletoItem boletoItem;
			BoletoItemRateio boletoItemRateio;
			byte byteNumParcela;
			String strDadosRateio;
			String[] vRateio;
			String[] v;
			#endregion

			#region [ Dados do registro principal ]
			if (cbBoletoCedente.SelectedValue != null) boletoEditado.id_boleto_cedente = (short)Global.converteInteiro(cbBoletoCedente.SelectedValue.ToString());
			boletoEditado.nome_sacado = txtClienteNome.Text;
			boletoEditado.endereco_sacado = txtEndereco.Text;
			
			boletoEditado.num_inscricao_sacado = Global.digitos(txtClienteCnpjCpf.Text);
			if (boletoEditado.num_inscricao_sacado.Length == 11)
				boletoEditado.tipo_sacado = Global.Cte.FIN.BoletoBradesco.CodTipoSacado.CPF;
			else if (boletoEditado.num_inscricao_sacado.Length == 14)
				boletoEditado.tipo_sacado = Global.Cte.FIN.BoletoBradesco.CodTipoSacado.CNPJ;
			else if (boletoEditado.num_inscricao_sacado.Length == 0)
				boletoEditado.tipo_sacado = Global.Cte.FIN.BoletoBradesco.CodTipoSacado.NAO_TEM;
			else
				boletoEditado.tipo_sacado = Global.Cte.FIN.BoletoBradesco.CodTipoSacado.OUTROS;
			
			boletoEditado.bairro_sacado = txtBairro.Text;
			boletoEditado.cep_sacado = Global.digitos(txtCep.Text);
			boletoEditado.cidade_sacado = txtCidade.Text;
			boletoEditado.uf_sacado = txtUF.Text;
			boletoEditado.email_sacado = txtEmail.Text;
			boletoEditado.segunda_mensagem = txtSegundaMensagem.Text;
			boletoEditado.mensagem_1 = txtMensagem1.Text;
			boletoEditado.mensagem_2 = txtMensagem2.Text;
			boletoEditado.mensagem_3 = txtMensagem3.Text;
			boletoEditado.mensagem_4 = txtMensagem4.Text;
			boletoEditado.numero_NF = boletoPreCadastradoSelecionado.numero_NF;
			boletoEditado.id_cliente = clienteSelecionado.id;
			boletoEditado.id_nf_parcela_pagto = boletoPreCadastradoSelecionado.id;
			boletoEditado.tipo_vinculo = Global.Cte.FIN.CodBoletoTipoVinculo.BOLETO_COM_PEDIDO_EMISSAO_AUTOMATICA;
			if (boletoCedenteSelecionado != null)
			{
				boletoEditado.codigo_empresa = boletoCedenteSelecionado.codigo_empresa;
				boletoEditado.nome_empresa = boletoCedenteSelecionado.nome_empresa;
				boletoEditado.num_banco = boletoCedenteSelecionado.num_banco;
				boletoEditado.nome_banco = boletoCedenteSelecionado.nome_banco;
				boletoEditado.agencia = boletoCedenteSelecionado.agencia;
				boletoEditado.digito_agencia = boletoCedenteSelecionado.digito_agencia;
				boletoEditado.conta = boletoCedenteSelecionado.conta;
				boletoEditado.digito_conta = boletoCedenteSelecionado.digito_conta;
				boletoEditado.carteira = boletoCedenteSelecionado.carteira;
			}
			boletoEditado.juros_mora = (double)Global.converteNumeroDecimal(txtJurosMora.Text);
			boletoEditado.perc_multa = (double)Global.converteNumeroDecimal(txtPercMulta.Text);
			boletoEditado.qtde_dias_protesto = (byte)Global.converteInteiro(txtProtestarApos.Text);

			// Outros campos
			boletoEditado.qtde_dias_decurso_prazo = 0;
			if (boletoEditado.qtde_dias_protesto == 0)
			{
				boletoEditado.primeira_instrucao = "00";
				boletoEditado.segunda_instrucao = "00";
			}
			else
			{
				boletoEditado.primeira_instrucao = "06";
				boletoEditado.segunda_instrucao = boletoEditado.qtde_dias_protesto.ToString().PadLeft(2, '0');
			}
			boletoEditado.qtde_parcelas = (byte)grdParcelas.Rows.Count;
			#endregion

			#region [ Dados das parcelas ]
			byteNumParcela = 0;
			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				boletoItem = new BoletoItem();
				byteNumParcela++;
				boletoItem.num_parcela = byteNumParcela;
				boletoItem.tipo_vencimento = Global.Cte.FIN.CodBoletoItemTipoVencto.VENCTO_DEFINIDO;
				boletoItem.dt_vencto = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString());
				boletoItem.valor = Global.converteNumeroDecimal(grdParcelas.Rows[i].Cells["grdParcelas_valor"].Value.ToString());
				boletoItem.dt_limite_desconto = boletoItem.dt_vencto;
				boletoItem.valor_desconto = 0m;
				boletoItem.bonificacao_por_dia = 0m;
				if (boletoEditado.juros_mora == 0d)
				{
					boletoItem.valor_por_dia_atraso = 0m;
				}
				else
				{
					boletoItem.valor_por_dia_atraso = boletoItem.valor * (decimal)(boletoEditado.juros_mora / 100) / 30;
					// Arredondamento
					boletoItem.valor_por_dia_atraso = Global.converteNumeroDecimal(Global.formataMoeda(boletoItem.valor_por_dia_atraso));
				}
				boletoItem.nosso_numero = boletoItem.nosso_numero.PadLeft(11, '0');
				boletoItem.digito_nosso_numero = "0";
				boletoItem.primeira_mensagem = boletoItem.num_parcela.ToString().PadLeft(2, '0') +
											   boletoEditado.qtde_parcelas.ToString().PadLeft(2, '0');
				boletoItem.numero_documento = boletoEditado.numero_NF.ToString() +
											  "/" +
											  boletoItem.num_parcela.ToString().PadLeft(2, '0');

				#region [ Instrução de protesto ]
				// Devido ao custo do cartório, apenas algumas parcelas serão geradas c/ instrução de protesto
				if (boletoEditado.primeira_instrucao.Equals("06"))
				{
					boletoItem.st_instrucao_protesto = Global.calculaBoletoItemFlagInstrucaoProtesto(boletoItem.num_parcela, boletoEditado.qtde_parcelas);
				}
				else
				{
					boletoItem.st_instrucao_protesto = 0;
				}
				#endregion

				#region [ Rateio ]
				strDadosRateio = grdParcelas.Rows[i].Cells["grdParcelas_dados_rateio"].Value.ToString();
				vRateio = strDadosRateio.Split('|');
				foreach (String rateio in vRateio)
				{
					if (rateio != null)
					{
						if (rateio.Trim().Length > 0)
						{
							boletoItemRateio = new BoletoItemRateio();
							v = rateio.Split('=');
							boletoItemRateio.pedido = v[0];
							boletoItemRateio.valor = Global.converteNumeroDecimal(v[1]);
							boletoItem.listaBoletoItemRateio.Add(boletoItemRateio);
						}
					}
				}
				#endregion

				boletoEditado.listaBoletoItem.Add(boletoItem);
			}
			#endregion

			return boletoEditado;
		}
		#endregion

		#region [ isBoletoEditado ]
		/// <summary>
		/// Compara os dados dos dois objetos para verificar se o usuário fez alguma edição
		/// </summary>
		/// <param name="boletoOriginal">
		/// Objeto contendo os dados originais
		/// </param>
		/// <param name="boletoEditado">
		/// Objeto contendo os dados atuais, de acordo com o que está nos campos na tela
		/// </param>
		/// <returns>
		/// true: houve edição nos dados
		/// false: não houve nenhuma edição
		/// </returns>
		private bool isBoletoEditado(BoletoPreCadastrado boletoOriginal, Boleto boletoEditado)
		{
			#region [ Declarações ]
			int indiceBoletoEditado = 0;
			String strEndereco;
			#endregion

			strEndereco = clienteSelecionado.endereco;
			if (clienteSelecionado.endereco_numero.Length > 0) strEndereco += ", " + clienteSelecionado.endereco_numero;
			if (clienteSelecionado.endereco_complemento.Length > 0) strEndereco += " " + clienteSelecionado.endereco_complemento;

			if (!clienteSelecionado.nome.ToUpper().Equals(boletoEditado.nome_sacado)) return true;
			if (!strEndereco.ToUpper().Equals(boletoEditado.endereco_sacado)) return true;
			if (!clienteSelecionado.cnpj_cpf.Equals(boletoEditado.num_inscricao_sacado)) return true;
			if (!clienteSelecionado.cep.Equals(boletoEditado.cep_sacado)) return true;
			if (!clienteSelecionado.email.Equals(boletoEditado.email_sacado)) return true;
			if (boletoCedenteSelecionado != null)
			{
				if (!boletoCedenteSelecionado.segunda_mensagem_padrao.Equals(boletoEditado.segunda_mensagem)) return true;
				if (!boletoCedenteSelecionado.mensagem_1_padrao.Equals(boletoEditado.mensagem_1)) return true;
				if (!boletoCedenteSelecionado.mensagem_2_padrao.Equals(boletoEditado.mensagem_2)) return true;
				if (!boletoCedenteSelecionado.mensagem_3_padrao.Equals(boletoEditado.mensagem_3)) return true;
				if (!boletoCedenteSelecionado.mensagem_4_padrao.Equals(boletoEditado.mensagem_4)) return true;
				if (boletoCedenteSelecionado.juros_mora != boletoEditado.juros_mora) return true;
				if (boletoCedenteSelecionado.perc_multa != boletoEditado.perc_multa) return true;
				if (boletoCedenteSelecionado.qtde_dias_protestar_apos_padrao != boletoEditado.qtde_dias_protesto) return true;
			}
			if (boletoOriginal.qtde_parcelas_boleto != boletoEditado.listaBoletoItem.Count) return true;
			for (int i = 0; i < boletoOriginal.listaItem.Count; i++)
			{
				if (boletoOriginal.listaItem[i].forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					if (boletoOriginal.listaItem[i].dt_vencto != boletoEditado.listaBoletoItem[indiceBoletoEditado].dt_vencto) return true;
					if (boletoOriginal.listaItem[i].valor != boletoEditado.listaBoletoItem[indiceBoletoEditado].valor) return true;
					if (boletoOriginal.listaItem[i].listaRateio.Count != boletoEditado.listaBoletoItem[indiceBoletoEditado].listaBoletoItemRateio.Count) return true;
					for (int j = 0; j < boletoOriginal.listaItem[i].listaRateio.Count; j++)
					{
						if (!boletoOriginal.listaItem[i].listaRateio[j].pedido.Equals(boletoEditado.listaBoletoItem[indiceBoletoEditado].listaBoletoItemRateio[j].pedido)) return true;
						if (boletoOriginal.listaItem[i].listaRateio[j].valor != boletoEditado.listaBoletoItem[indiceBoletoEditado].listaBoletoItemRateio[j].valor) return true;
					}
					indiceBoletoEditado++;
				}
			}

			return false;
		}
		#endregion

		#region [ consisteCampos ]
		/// <summary>
		/// Realiza a consistência dos campos na tela
		/// </summary>
		/// <returns>
		/// true: os campos estão devidamente preenchidos
		/// false: há campos não preenchidos corretamente
		/// </returns>
		private bool consisteCampos()
		{
			#region [ Declarações ]
			decimal vlTotalParcelas = 0;
			String strMsgErro;
			String strRelacaoEmailInvalido = "";
			String[] v;
			DateTime dt1, dt2;
			int intQtdeDiasProtestarApos;
			bool blnOk;
			BoletoPlanoContasDestino boletoPlanoContasDestino;
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
			#endregion

			#region [ Plano de contas ]
			try
			{
				boletoPlanoContasDestino = BoletoPreCadastradoDAO.obtemBoletoPlanoContasDestino(boletoPreCadastradoSelecionado.id);
			}
			catch (Exception ex)
			{
				strMsgErro = ex.Message;
				avisoErro(strMsgErro);
				return false;
			}
			#endregion

			#region [ Conta do cedente ]
			if (cbBoletoCedente.SelectedIndex == -1)
			{
				avisoErro("É necessário informar a conta do cedente!!");
				cbBoletoCedente.Focus();
				return false;
			}
			#endregion

			#region [ Nome do sacado ]
			if (txtClienteNome.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o nome do cliente!!");
				txtClienteNome.Focus();
				return false;
			}

			if (txtClienteNome.Text.Length > Global.Cte.Etc.MAX_TAM_BOLETO_CAMPO_NOME_SACADO)
			{
				avisoErro("É necessário editar o nome do cliente, pois está excedendo o tamanho máximo!!");
				txtClienteNome.Focus();
				if (txtClienteNome.Text.Length > 0)
				{
					txtClienteNome.SelectionStart = txtClienteNome.Text.Length;
					txtClienteNome.SelectionLength = 0;
				}
				return false;
			}
			#endregion

			#region [ Endereço ]
			if (txtEndereco.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o endereço do cliente!!");
				txtEndereco.Focus();
				return false;
			}

			if (txtEndereco.Text.Length > Global.Cte.Etc.MAX_TAM_BOLETO_CAMPO_ENDERECO)
			{
				avisoErro("É necessário editar o endereço, pois está excedendo o tamanho máximo!!");
				txtEndereco.Focus();
				if (txtEndereco.Text.Length > 0)
				{
					txtEndereco.SelectionStart = txtEndereco.Text.Length;
					txtEndereco.SelectionLength = 0;
				}
				return false;
			}
			#endregion

			#region [ CNPJ / CPF ]
			if (txtClienteCnpjCpf.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o CNPJ/CPF do cliente!!");
				txtClienteCnpjCpf.Focus();
				return false;
			}

			if (!Global.isCnpjCpfOk(txtClienteCnpjCpf.Text))
			{
				avisoErro("CNPJ/CPF do cliente é inválido!!");
				txtClienteCnpjCpf.Focus();
				return false;
			}
			#endregion

			#region [ CEP ]
			if (txtCep.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o CEP do cliente!!");
				txtCep.Focus();
				if (txtCep.Text.Length > 0)
				{
					txtCep.SelectionStart = txtCep.Text.Length;
					txtCep.SelectionLength = 0;
				}
				return false;
			}

			if (!Global.isCepOk(txtCep.Text))
			{
				avisoErro("CEP do cliente é inválido!!");
				txtCep.Focus();
				return false;
			}
			#endregion

			#region [ UF ]
			if (txtUF.Text.Length > 0)
			{
				if (!Global.isUfOk(txtUF.Text))
				{
					avisoErro("UF inválida!!");
					txtUF.Focus();
					return false;
				}
			}
			#endregion

			#region [ E-mail ]
			if (txtEmail.Text.Length > 0)
			{
				if (!Global.isEmailOk(txtEmail.Text, ref strRelacaoEmailInvalido))
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
					txtEmail.Focus();
					return false;
				}
			}
			#endregion

			#region [ Protestar após ]
			intQtdeDiasProtestarApos = (int)Global.converteInteiro(txtProtestarApos.Text);
			if ((intQtdeDiasProtestarApos > 0) && (intQtdeDiasProtestarApos < 5))
			{
				avisoErro("Para a instrução de protesto, é exigido o mínimo de 5 dias!!");
				txtProtestarApos.Focus();
				return false;
			}
			#endregion

			#region [ Parcelas ]

			#region [ Há parcelas? ]
			if (grdParcelas.Rows.Count == 0)
			{
				avisoErro("Não há parcelas de boleto a gerar!!");
				btnDummy.Focus();
				return false;
			}
			#endregion

			#region [ Consistência de vencimento / valor ]
			blnOk = true;
			strMsgErro = "";
			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				dt1 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString());
				if (dt1 < DateTime.Today)
				{
					blnOk = false;
					if (strMsgErro.Length > 0) strMsgErro += "\n";
					strMsgErro += "A parcela " + grdParcelas.Rows[i].Cells["grdParcelas_num_parcela"].Value.ToString() + " tem como data de vencimento uma data passada (" + grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString() + ")";
				}
				if (Global.converteNumeroDecimal(grdParcelas.Rows[i].Cells["grdParcelas_valor"].Value.ToString()) < 0)
				{
					blnOk = false;
					if (strMsgErro.Length > 0) strMsgErro += "\n";
					strMsgErro += "A parcela " + grdParcelas.Rows[i].Cells["grdParcelas_num_parcela"].Value.ToString() + " está com valor inválido (" + grdParcelas.Rows[i].Cells["grdParcelas_valor"].Value.ToString() + ")";
				}
			}
			if (!blnOk)
			{
				avisoErro(strMsgErro);
				return false;
			}
			#endregion

			#region [ Há parcelas com data de vencimento repetida? ]
			blnOk = true;
			strMsgErro = "";
			for (int i = 0; i < (grdParcelas.Rows.Count - 1); i++)
			{
				dt1 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString());
				dt2 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i + 1].Cells["grdParcelas_dt_vencto"].Value.ToString());
				if (dt1 == dt2)
				{
					blnOk = false;
					if (strMsgErro.Length > 0) strMsgErro += "\n";
					strMsgErro += "As parcelas " + grdParcelas.Rows[i].Cells["grdParcelas_num_parcela"].Value.ToString() + " e " + grdParcelas.Rows[i + 1].Cells["grdParcelas_num_parcela"].Value.ToString() + " possuem a mesma data de vencimento (" + grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString() + ")";
				}
			}
			if (!blnOk)
			{
				avisoErro(strMsgErro);
				return false;
			}
			#endregion

			#endregion

			#region [ Valor total confere? ]
			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				vlTotalParcelas += Global.converteNumeroDecimal(grdParcelas.Rows[i].Cells["grdParcelas_valor"].Value.ToString());
			}

			if (_vlTotalBoletoPedidos != vlTotalParcelas)
			{
				strMsgErro = "Divergência de valores:" +
							"\nTotal do(s) pedido(s) em boleto: " + Global.formataMoeda(_vlTotalBoletoPedidos) +
							"\nTotal do(s) boleto(s) em cadastramento: " + Global.formataMoeda(vlTotalParcelas) +
							"\n" +
							"Digite a senha para confirmar a gravação mesmo havendo esta divergência!!";
				fAutorizacao = new FAutorizacao(strMsgErro);
				drAutorizacao = fAutorizacao.ShowDialog();
				if (drAutorizacao != DialogResult.OK)
				{
					avisoErro("Operação cancelada!!");
					return false;
				}
				if (fAutorizacao.senha.ToUpper() != Global.Usuario.senhaDescriptografada.ToUpper())
				{
					avisoErro("Senha inválida!!\nA operação não foi realizada!!");
					return false;
				}
			}
			#endregion

			return true;
		}
		#endregion

		#region [ consisteBoleto ]
		private bool consisteBoleto()
		{
			#region [ Declarações ]
			int intNumeroAviso = 0;
			String strMsgConfirmacao = "";
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Verifica Status de Pagamento ]
			for (int i = 0; i < _listaPedidos.Count; i++)
			{
				if (_listaPedidos[i].st_pagto.Equals(Global.Cte.StPagtoPedido.ST_PAGTO_PAGO))
				{
					intNumeroAviso++;
					if (strMsgConfirmacao.Length > 0) strMsgConfirmacao += "\n";
					strMsgConfirmacao += intNumeroAviso.ToString() + ") O pedido " + _listaPedidos[i].pedido + " está com status de pagamento: " + Global.stPagtoPedidoDescricao(_listaPedidos[i].st_pagto).ToUpper() + "!!";
				}
			}
			#endregion

			#region [ Verifica se já possui boleto emitido ]
			for (int i = 0; i < lbPedido.Items.Count; i++)
			{
				strSql = "SELECT DISTINCT" +
							" tFBI.status" +
						" FROM t_FIN_BOLETO_ITEM tFBI" +
							" INNER JOIN t_FIN_BOLETO_ITEM_RATEIO tFBIR" +
								" ON (tFBI.id=tFBIR.id_boleto_item)" +
						" WHERE" +
							" (tFBIR.pedido = '" + Global.normalizaNumeroPedido(lbPedido.Items[i].ToString()) + "')" +
							" AND " +
								"(" +
									"tFBI.status NOT IN " +
									"(" +
										Global.Cte.FIN.CodBoletoItemStatus.CANCELADO_MANUAL.ToString() + "," +
										Global.Cte.FIN.CodBoletoItemStatus.BOLETO_BAIXADO.ToString() + "," +
										Global.Cte.FIN.CodBoletoItemStatus.VALA_COMUM.ToString() +
									")" +
								 ")" +
						" ORDER BY" +
							" tFBI.status";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count > 0)
				{
					intNumeroAviso++;
					if (strMsgConfirmacao.Length > 0) strMsgConfirmacao += "\n";
					strMsgConfirmacao += intNumeroAviso.ToString() + ") O pedido " + Global.normalizaNumeroPedido(lbPedido.Items[i].ToString()) + " já possui boletos cadastrados!!";
				}
			}
			#endregion

			if (strMsgConfirmacao.Length > 0)
			{
				strMsgConfirmacao += "\n\nConfirma o cadastramento assim mesmo?";
				fAutorizacao = new FAutorizacao(strMsgConfirmacao);
				while (true)
				{
					drAutorizacao = fAutorizacao.ShowDialog();
					if (drAutorizacao != DialogResult.OK)
					{
						avisoErro("Operação de cadastramento cancelada!!");
						return false;
					}
					else
					{
						if (fAutorizacao.senha.ToUpper() == Global.Usuario.senhaDescriptografada.ToUpper())
						{
							// Prossegue com o cadastramento dos boletos!
							return true;
						}
						else
						{
							avisoErro("Senha inválida!!");
						}
					}
				}
			}

			return true;
		}
		#endregion

		#region [ consisteBoletoCedente ]
		private bool consisteBoletoCedente(Boleto boletoEditado)
		{
			#region [ Declarações ]
			bool blnBoletoCedenteDivergente = false;
			String strMsgConfirmacao;
			NFeEmitente nfeEmitente;
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
			#endregion

			if (boletoEditado.id_boleto_cedente == 0)
			{
				avisoErro("Não foi selecionado um cedente válido!!");
				return false;
			}

			#region [ Verifica se o cedente do boleto está coerente com a empresa definida no pedido para ser emitente da NFe ]
			for (int i = 0; i < _listaPedidos.Count; i++)
			{
				if (_listaPedidos[i].id_nfe_emitente > 0)
				{
					nfeEmitente = NFeEmitenteDAO.getNFeEmitenteById(_listaPedidos[i].id_nfe_emitente);
					if (nfeEmitente != null)
					{
						if (nfeEmitente.id_boleto_cedente != boletoEditado.id_boleto_cedente)
						{
							blnBoletoCedenteDivergente = true;
							break;
						}
					}
				}
			}
			#endregion

			if (blnBoletoCedenteDivergente)
			{
				strMsgConfirmacao = "O cedente selecionado diverge da opção pré-definida no sistema!!\n\nConfirma o cadastramento assim mesmo?";
				fAutorizacao = new FAutorizacao(strMsgConfirmacao);
				while (true)
				{
					drAutorizacao = fAutorizacao.ShowDialog();
					if (drAutorizacao != DialogResult.OK)
					{
						avisoErro("Operação de cadastramento cancelada!!");
						return false;
					}
					else
					{
						if (fAutorizacao.senha.ToUpper() == Global.Usuario.senhaDescriptografada.ToUpper())
						{
							// Prossegue com o cadastramento dos boletos!
							return true;
						}
						else
						{
							avisoErro("Senha inválida!!");
						}
					}
				}
			}

			return true;
		}
		#endregion

		#region [ trataBotaoAnular ]
		private void trataBotaoAnular()
		{
			#region [ Declarações ]
			String strAux;
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strDescricaoLog = "";
			bool blnResultado;
			FinLog finLog = new FinLog();
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
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

			#region[ Confirmação ]
			strAux = "Os dados gerados durante a emissão da nota fiscal serão anulados!!" +
					 "\nA operação é irreversível!!" +
					 "\nDigite a senha para confirmar a ANULAÇÃO!!";
			fAutorizacao = new FAutorizacao(strAux);
			drAutorizacao = fAutorizacao.ShowDialog();
			if (drAutorizacao != DialogResult.OK)
			{
				avisoErro("Operação não confirmada!!\nA anulação dos dados não foi realizada!!");
				return;
			}
			if (fAutorizacao.senha.ToUpper() != Global.Usuario.senhaDescriptografada.ToUpper())
			{
				avisoErro("Senha inválida!!\nA anulação dos dados não foi realizada!!");
				return;
			}
			#endregion

			#region [ Anula os dados no banco de dados ]
			blnResultado = BoletoPreCadastradoDAO.anula(	Global.Usuario.usuario,
															boletoPreCadastradoSelecionado.id,
															ref strDescricaoLog,
															ref strMsgErro
															);
			#endregion

			#region [ Processamento pós tentativa de exclusão do BD ]
			if (blnResultado)
			{
				#region [ Grava log no BD ]
				finLog.usuario = Global.Usuario.usuario;
				finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_PRE_CADASTRADO_ANULA;
				finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
				finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
				finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_NF_PARCELA_PAGTO;
				finLog.id_registro_origem = boletoPreCadastradoSelecionado.id;
				finLog.id_cliente = boletoPreCadastradoSelecionado.id_cliente;
				finLog.cnpj_cpf = clienteSelecionado.cnpj_cpf;
				finLog.descricao = strDescricaoLog;
				FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
				#endregion

				_blnRegistroFoiAnulado = true;
				aviso("Os dados para cadastramento deste boleto foram anulados!!");
				// Fecha o painel!!
				Close();
			}
			else
			{
				avisoErro("Falha ao gravar o registro!!\n\n" + strMsgErro);
			}
			#endregion
		}
		#endregion

		#region [ trataBotaoCadastrar ]
		void trataBotaoCadastrar()
		{
			#region [ Declarações ]
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strDescricaoLog = "";
			String strDescricaoLogBoletoPreCadastrado = "";
			bool blnResultado = false;
			Boleto boletoEditado;
			FinLog finLog = new FinLog();
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
			if (!consisteCampos()) return;
			#endregion

			#region [ Obtém valores ]
			boletoEditado = obtemDadosBoletoCamposTela();
			if (boletoEditado == null) return;
			#endregion

			#region [ Confirmação ]
			if (!confirma("Confirma a gravação do boleto?")) return;
			#endregion

			#region [ Consistência com relação ao cadastro de boletos ]
			if (!consisteBoleto()) return;
			#endregion

			#region [ Consistência com relação ao cedente do boleto ]
			if (!consisteBoletoCedente(boletoEditado)) return;
			#endregion

			try
			{
				try
				{
					BD.iniciaTransacao();

					#region [ Grava o boleto no banco de dados ]
					blnResultado = BoletoDAO.boletoInsere(Global.Usuario.usuario,
														  boletoEditado,
														  ref strDescricaoLog,
														  ref strMsgErro
														  );
					if (!blnResultado) strMsgErro = "Falha ao gravar o boleto no banco de dados!!\n\n" + strMsgErro;
					#endregion

					#region [ Marca como já tratado o registro de t_FIN_NF_PARCELA_PAGTO ]
					if (blnResultado)
					{
						blnResultado = BoletoPreCadastradoDAO.marcaComoTratado(Global.Usuario.usuario,
																	 boletoPreCadastradoSelecionado.id,
																	 ref strDescricaoLogBoletoPreCadastrado,
																	 ref strMsgErro);
						if (!blnResultado) strMsgErro = "Falha ao marcar como já tratado o registro com os dados gerados quando a NF foi impressa!!\n\n" + strMsgErro;
					}
					#endregion
				}
				finally
				{
					if (blnResultado)
					{
						try
						{
							BD.commitTransacao();
						}
						catch (Exception ex)
						{
							strMsgErro = ex.Message;
							blnResultado = false;
						}
					}
					else
					{
						try
						{
							BD.rollbackTransacao();
						}
						catch (Exception ex)
						{
							Global.gravaLogAtividade(ex.ToString());
							avisoErro(ex.ToString());
						}
					}
				}

				#region [ Processamento pós tentativa de gravação no BD ]
				if (blnResultado)
				{
					#region [ Grava log no BD ]
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_CADASTRA;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO;
					finLog.id_registro_origem = boletoEditado.id;
					finLog.id_cliente = boletoEditado.id_cliente;
					finLog.cnpj_cpf = boletoEditado.num_inscricao_sacado;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
					#endregion

					_blnRegistroFoiGravado = true;
					SystemSounds.Asterisk.Play();
					Close();
				}
				else
				{
					if (strMsgErro.Length == 0) strMsgErro = "Falha desconhecida durante a operação!!";
					Global.gravaLogAtividade(strMsgErro);
					avisoErro(strMsgErro);
				}
				#endregion
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(ex.ToString());
				avisoErro(ex.ToString());
			}
		}
		#endregion

		#region [ trataBotaoPesquisaCep ]
		void trataBotaoPesquisaCep()
		{
			#region [ Declarações ]
			String strEspaco = "";
			DialogResult drCep;
			String s1, s2, strEndereco;
			char c1, c2;
			#endregion

			#region [ Exibe painel de consulta de CEP ]
			fCepPesquisa = new FCepPesquisa();
			fCepPesquisa.StartPosition = FormStartPosition.Manual;
			fCepPesquisa.Location = this.Location;
			fCepPesquisa.cepDefault = txtCep.Text;
			drCep = fCepPesquisa.ShowDialog();
			this.Location = fCepPesquisa.Location;
			if (drCep != DialogResult.OK) return;
			#endregion

			#region [ Atualiza com o resultado da pesquisa ]
			if (fCepPesquisa.cepSelecionado.Trim().Length > 0) txtCep.Text = fCepPesquisa.cepSelecionado;
			if (fCepPesquisa.bairroSelecionado.Trim().Length > 0) txtBairro.Text = fCepPesquisa.bairroSelecionado;
			if (fCepPesquisa.cidadeSelecionada.Trim().Length > 0) txtCidade.Text = fCepPesquisa.cidadeSelecionada;
			if (fCepPesquisa.ufSelecionado.Trim().Length > 0) txtUF.Text = fCepPesquisa.ufSelecionado;

			s1 = fCepPesquisa.logradouroSelecionado;
			s2 = fCepPesquisa.numeroOuComplementoSelecionado;
			if ((s1.Length > 0) && (s2.Length > 0))
			{
				c1 = s1.Substring(s1.Length - 1, 1)[0];
				c2 = s2.Substring(0, 1)[0];
				if (Global.isAlfaNumerico(c1) && Global.isAlfaNumerico(c2)) strEspaco = " ";
			}
			strEndereco = s1 + strEspaco + s2;
			if (strEndereco.Trim().Length > 0) txtEndereco.Text = strEndereco;
			#endregion
		}
		#endregion

		#region [ trataBotaoRestauraEnderecoOriginal ]
		void trataBotaoRestauraEnderecoOriginal()
		{
			#region [ Declarações ]
			String strEndereco;
			#endregion

			if (!confirma("Restaura o endereço original?")) return;

			strEndereco = clienteSelecionado.endereco;
			if (clienteSelecionado.endereco_numero.Length > 0) strEndereco += ", " + clienteSelecionado.endereco_numero;
			if (clienteSelecionado.endereco_complemento.Length > 0) strEndereco += " " + clienteSelecionado.endereco_complemento;

			txtEndereco.Text = strEndereco.ToUpper();
			txtCep.Text = Global.formataCep(clienteSelecionado.cep);
			txtBairro.Text = clienteSelecionado.bairro.ToUpper();
			txtCidade.Text = clienteSelecionado.cidade.ToUpper();
			txtUF.Text = clienteSelecionado.uf.ToUpper();
		}
		#endregion

		#region [ trataBotaoPedidoConsulta ]
		void trataBotaoPedidoConsulta()
		{
			#region [ Consistências ]
			if (lbPedido.Items.Count == 0)
			{
				avisoErro("Não há pedidos para consultar!!");
				return;
			}
			#endregion

			#region [ Exibe painel de detalhes do pedido ]
			fPedido = new FPedido(this);
			fPedido.StartPosition = FormStartPosition.Manual;
			fPedido.Location = this.Location;
			if (lbPedido.SelectedIndex == -1)
				fPedido.numeroPedidoDefault = lbPedido.Items[0].ToString();
			else
				fPedido.numeroPedidoDefault = lbPedido.Items[lbPedido.SelectedIndex].ToString();
			fPedido.listaPedidos = _listaPedidos;
			fPedido.cliente = clienteSelecionado;
			fPedido.Show();
			this.Visible = false;
			#endregion
		}
		#endregion

		#region [ trataBotaoExcluirParcela ]
		private void trataBotaoExcluirParcela()
		{
			#region [ Declarações ]
			int intLinhaGridSelecionada = -1;
			String strMsg;
			#endregion

			#region [ Consistências ]
			if (grdParcelas.Rows.Count == 0)
			{
				avisoErro("Não há parcelas!!");
				return;
			}

			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				if (grdParcelas.Rows[i].Selected)
				{
					intLinhaGridSelecionada = i;
					break;
				}
			}

			if (intLinhaGridSelecionada < 0)
			{
				avisoErro("Nenhuma parcela foi selecionada para exclusão!!");
				return;
			}
			#endregion

			#region [ Confirmação ]
			strMsg = "Parcela: " + grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_dt_vencto"].Value + " de " + grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_valor"].Value +
					 "\n\n" +
					 "Confirma a exclusão desta parcela?";
			if (!confirma(strMsg)) return;
			#endregion

			#region [ Exclui a linha do grid ]
			grdParcelas.Rows.RemoveAt(intLinhaGridSelecionada);
			#endregion

			#region [ Atualiza campos ]
			atualizaNumeracaoParcelas();
			recalculaValorTotalParcelas();
			#endregion
		}
		#endregion

		#region [ trataBotaoEditarParcela ]
		private void trataBotaoEditarParcela()
		{
			#region [ Declarações ]
			int intLinhaGridSelecionada = -1;
			String strMsg;
			DateTime dt1, dt2;
			DialogResult drBoletoParcelaEdita;
			String[] vTemp;
			bool blnOrdenado;
			#endregion

			#region [ Consistências ]
			if (grdParcelas.Rows.Count == 0)
			{
				avisoErro("Não há parcelas!!");
				return;
			}

			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				if (grdParcelas.Rows[i].Selected)
				{
					intLinhaGridSelecionada = i;
					break;
				}
			}

			if (intLinhaGridSelecionada < 0)
			{
				avisoErro("Nenhuma parcela foi selecionada para edição!!");
				return;
			}
			#endregion

			#region [ Exibe painel para editar parcela ]
			fBoletoParcelaEdita = new FBoletoParcelaEdita(FBoletoParcelaEdita.eBoletoParcelaOperacao.EDITAR);
			fBoletoParcelaEdita.StartPosition = FormStartPosition.Manual;
			fBoletoParcelaEdita.Left = this.Left + (this.Width - fBoletoParcelaEdita.Width) / 2;
			fBoletoParcelaEdita.Top = this.Top + (this.Height - fBoletoParcelaEdita.Height) / 2;
			fBoletoParcelaEdita.parcelaSelecionadaDadosRateio = grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_dados_rateio"].Value.ToString();
			fBoletoParcelaEdita.parcelaSelecionadaDtVencto = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_dt_vencto"].Value.ToString());
			fBoletoParcelaEdita.parcelaSelecionadaValor = Global.converteNumeroDecimal(grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_valor"].Value.ToString());
			drBoletoParcelaEdita = fBoletoParcelaEdita.ShowDialog();
			if (drBoletoParcelaEdita != DialogResult.OK) return;
			#endregion

			#region [ Consistência da data de vencimento ]
			if (fBoletoParcelaEdita.parcelaSelecionadaDtVencto < DateTime.Today)
			{
				strMsg = "A data de vencimento da parcela editada é uma data passada (" + Global.formataDataDdMmYyyyComSeparador(fBoletoParcelaEdita.parcelaSelecionadaDtVencto) + ")" +
						 "\n\n" +
						 "As alterações serão canceladas!!";
				avisoErro(strMsg);
				return;
			}

			if (fBoletoParcelaEdita.parcelaSelecionadaDtVencto < DateTime.Today.AddDays(20))
			{
				strMsg = "A data de vencimento da parcela editada está muito próxima de hoje (" + Global.formataDataDdMmYyyyComSeparador(fBoletoParcelaEdita.parcelaSelecionadaDtVencto) + ")" +
						 "\n\n" +
						 "Continua mesmo assim?";
				if (!confirma(strMsg))
				{
					aviso("As alterações foram canceladas!!");
					return;
				}
			}
			#endregion

			#region [ Verifica se editou a data de vencimento p/ uma data já usada por outra parcela ]
			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				if (i != intLinhaGridSelecionada)
				{
					dt1 = fBoletoParcelaEdita.parcelaSelecionadaDtVencto;
					dt2 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString());
					if (dt1 == dt2)
					{
						strMsg = "A data de vencimento da parcela editada (" + Global.formataDataDdMmYyyyComSeparador(dt1) + ") é a mesma data já usada pela parcela " + grdParcelas.Rows[i].Cells["grdParcelas_num_parcela"].Value.ToString() +
								 "\n\n" +
								 "As alterações serão canceladas!!";
						avisoErro(strMsg);
						return;
					}
				}
			}
			#endregion

			#region [ Atualiza os dados editados ]
			grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_dt_vencto"].Value = Global.formataDataDdMmYyyyComSeparador(fBoletoParcelaEdita.parcelaSelecionadaDtVencto);
			grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_valor"].Value = Global.formataMoeda(fBoletoParcelaEdita.parcelaSelecionadaValor);
			grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_dados_rateio"].Value = fBoletoParcelaEdita.parcelaSelecionadaDadosRateio;
			#endregion

			#region [ No caso de ter alterado a data de vencimento, reordena as parcelas ]
			vTemp = new String[grdParcelas.Columns.Count];
			do
			{
				blnOrdenado = true;
				for (int i = 0; i < (grdParcelas.Rows.Count - 1); i++)
				{
					dt1 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString());
					dt2 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i + 1].Cells["grdParcelas_dt_vencto"].Value.ToString());
					if (dt1 > dt2)
					{
						blnOrdenado = false;
						for (int j = 0; j < grdParcelas.Columns.Count; j++)
						{
							vTemp[j] = grdParcelas.Rows[i + 1].Cells[j].Value.ToString();
						}
						for (int j = 0; j < grdParcelas.Columns.Count; j++)
						{
							grdParcelas.Rows[i + 1].Cells[j].Value = grdParcelas.Rows[i].Cells[j].Value;
						}
						for (int j = 0; j < grdParcelas.Columns.Count; j++)
						{
							grdParcelas.Rows[i].Cells[j].Value = vTemp[j];
						}
					}
				}
			} while (!blnOrdenado);

			atualizaNumeracaoParcelas();
			#endregion

			#region [ Recalcula o valor total ]
			recalculaValorTotalParcelas();
			#endregion

			#region [ Exibe o grid com a parcela editada selecionada ]
			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				dt1 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString());
				if (dt1 == fBoletoParcelaEdita.parcelaSelecionadaDtVencto)
				{
					grdParcelas.Rows[i].Selected = true;
					grdParcelas.FirstDisplayedScrollingRowIndex = i;
					break;
				}
			}
			#endregion
		}
		#endregion

		#region [ trataBotaoAdicionarParcela ]
		private void trataBotaoAdicionarParcela()
		{
			#region [ Declarações ]
			String strMsg;
			String strDadosRateio = "";
			int intLinhaGridSelecionada;
			DateTime dt1, dt2;
			DialogResult drBoletoParcelaEdita;
			String[] vTemp;
			bool blnOrdenado;
			#endregion

			#region [ Monta os dados de rateio ]
			for (int i = 0; i < lbPedido.Items.Count; i++)
			{
				if (strDadosRateio.Length > 0) strDadosRateio += "|";
				strDadosRateio += lbPedido.Items[i].ToString() + "=" + Global.formataMoeda(0m);
			}
			#endregion

			#region [ Exibe painel para preencher os dados de nova parcela ]
			fBoletoParcelaEdita = new FBoletoParcelaEdita(FBoletoParcelaEdita.eBoletoParcelaOperacao.INCLUIR);
			fBoletoParcelaEdita.StartPosition = FormStartPosition.Manual;
			fBoletoParcelaEdita.Left = this.Left + (this.Width - fBoletoParcelaEdita.Width) / 2;
			fBoletoParcelaEdita.Top = this.Top + (this.Height - fBoletoParcelaEdita.Height) / 2;
			fBoletoParcelaEdita.parcelaSelecionadaDadosRateio = strDadosRateio;
			fBoletoParcelaEdita.parcelaSelecionadaDtVencto = DateTime.MinValue;
			fBoletoParcelaEdita.parcelaSelecionadaValor = 0;
			drBoletoParcelaEdita = fBoletoParcelaEdita.ShowDialog();
			if (drBoletoParcelaEdita != DialogResult.OK) return;
			#endregion

			#region [ Consistência da data de vencimento ]
			if (fBoletoParcelaEdita.parcelaSelecionadaDtVencto < DateTime.Today)
			{
				strMsg = "A data de vencimento da nova parcela é uma data passada (" + Global.formataDataDdMmYyyyComSeparador(fBoletoParcelaEdita.parcelaSelecionadaDtVencto) + ")" +
						 "\n\n" +
						 "A operação será cancelada!!";
				avisoErro(strMsg);
				return;
			}

			if (fBoletoParcelaEdita.parcelaSelecionadaDtVencto < DateTime.Today.AddDays(20))
			{
				strMsg = "A data de vencimento da nova parcela está muito próxima de hoje (" + Global.formataDataDdMmYyyyComSeparador(fBoletoParcelaEdita.parcelaSelecionadaDtVencto) + ")" +
						 "\n\n" +
						 "Continua mesmo assim?";
				if (!confirma(strMsg))
				{
					aviso("A operação foi cancelada!!");
					return;
				}
			}
			#endregion

			#region [ Verifica se editou a data de vencimento p/ uma data já usada por outra parcela ]
			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				dt1 = fBoletoParcelaEdita.parcelaSelecionadaDtVencto;
				dt2 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString());
				if (dt1 == dt2)
				{
					strMsg = "A data de vencimento da nova parcela (" + Global.formataDataDdMmYyyyComSeparador(dt1) + ") é a mesma data já usada pela parcela " + grdParcelas.Rows[i].Cells["grdParcelas_num_parcela"].Value.ToString() +
							 "\n\n" +
							 "A operação será cancelada!!";
					avisoErro(strMsg);
					return;
				}
			}
			#endregion

			#region [ Inclui a nova parcela no grid ]
			grdParcelas.Rows.Add();
			intLinhaGridSelecionada = grdParcelas.Rows.Count - 1;
			grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_num_parcela"].Value = "";
			grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_dt_vencto"].Value = Global.formataDataDdMmYyyyComSeparador(fBoletoParcelaEdita.parcelaSelecionadaDtVencto);
			grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_valor"].Value = Global.formataMoeda(fBoletoParcelaEdita.parcelaSelecionadaValor);
			grdParcelas.Rows[intLinhaGridSelecionada].Cells["grdParcelas_dados_rateio"].Value = fBoletoParcelaEdita.parcelaSelecionadaDadosRateio;
			#endregion

			#region [ No caso de ter alterado a data de vencimento, reordena as parcelas ]
			vTemp = new String[grdParcelas.Columns.Count];
			do
			{
				blnOrdenado = true;
				for (int i = 0; i < (grdParcelas.Rows.Count - 1); i++)
				{
					dt1 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString());
					dt2 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i + 1].Cells["grdParcelas_dt_vencto"].Value.ToString());
					if (dt1 > dt2)
					{
						blnOrdenado = false;
						for (int j = 0; j < grdParcelas.Columns.Count; j++)
						{
							vTemp[j] = grdParcelas.Rows[i + 1].Cells[j].Value.ToString();
						}
						for (int j = 0; j < grdParcelas.Columns.Count; j++)
						{
							grdParcelas.Rows[i + 1].Cells[j].Value = grdParcelas.Rows[i].Cells[j].Value;
						}
						for (int j = 0; j < grdParcelas.Columns.Count; j++)
						{
							grdParcelas.Rows[i].Cells[j].Value = vTemp[j];
						}
					}
				}
			} while (!blnOrdenado);

			atualizaNumeracaoParcelas();
			#endregion

			#region [ Recalcula o valor total ]
			recalculaValorTotalParcelas();
			#endregion

			#region [ Exibe o grid com a nova parcela selecionada ]
			for (int i = 0; i < grdParcelas.Rows.Count; i++)
			{
				dt1 = Global.converteDdMmYyyyParaDateTime(grdParcelas.Rows[i].Cells["grdParcelas_dt_vencto"].Value.ToString());
				if (dt1 == fBoletoParcelaEdita.parcelaSelecionadaDtVencto)
				{
					grdParcelas.Rows[i].Selected = true;
					grdParcelas.FirstDisplayedScrollingRowIndex = i;
					break;
				}
			}
			#endregion
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ Form: FBoletoCadastraDetalhe ]

		#region [ FBoletoCadastraDetalhe_Load ]
		private void FBoletoCadastraDetalhe_Load(object sender, EventArgs e)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			#endregion

			try
			{
				limpaCampos();
				blnSucesso = true;
			}
			catch (Exception ex)
			{
				_OcorreuExceptionNaInicializacao = true;
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

		#region [ FBoletoCadastraDetalhe_Shown ]
		private void FBoletoCadastraDetalhe_Shown(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String strPedido;
			String strDadosRateio;
			String strDadosRateioParcela;
			String strEndereco;
			bool blnAchou;
			Pedido pedido;
			int intIndiceLinhaGrid;
			int intQtdeParcelasBoleto = 0;
			decimal vlTotalParcelasNF = 0;
			decimal vlTotalParcelasBoleto = 0;
			#endregion

			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					btnDummy.Focus();
					info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

					#region [ Obtém os dados para cadastramento do boleto ]
					try
					{
						boletoPreCadastradoSelecionado = BoletoPreCadastradoDAO.getBoletoPreCadastrado(_idBoletoPreCadastradoSelecionado);
						clienteSelecionado = ClienteDAO.getCliente(boletoPreCadastradoSelecionado.id_cliente);
						_blnRegistroFoiGravado = false;
						_blnRegistroFoiAnulado = false;
					}
					catch (FinanceiroException ex)
					{
						avisoErro("Falha ao obter os dados para cadastramento do boleto!!\n\n" + ex.Message);
						Close();
						return;
					}
					#endregion

					#region [ Lista de pedidos que fazem parte do rateio ]
					lbPedido.Items.Clear();
					for (int i = 0; i < boletoPreCadastradoSelecionado.listaItem.Count; i++)
					{
						for (int j = 0; j < boletoPreCadastradoSelecionado.listaItem[i].listaRateio.Count; j++)
						{
							blnAchou = false;
							strPedido = boletoPreCadastradoSelecionado.listaItem[i].listaRateio[j].pedido;
							for (int k = 0; k < lbPedido.Items.Count; k++)
							{
								if (lbPedido.Items[k].ToString().Equals(strPedido))
								{
									blnAchou = true;
									break;
								}
							}
							if (!blnAchou) lbPedido.Items.Add(strPedido);
						}
					}
					#endregion

					#region [ Obtém dados dos pedidos no BD ]
					for (int i = 0; i < lbPedido.Items.Count; i++)
					{
						pedido = PedidoDAO.getPedido(lbPedido.Items[i].ToString());
						_listaPedidos.Add(pedido);

						_vlTotalBoletoPedidos += pedido.vlTotalBoletoDestePedido;
						_vlTotalFormaPagtoPedidos += pedido.vlTotalFormaPagtoDestePedido;
						_vlTotalPrecoNfPedidos += pedido.vlTotalPrecoNfDestePedido;
						_vlTotalFamiliaPagoPedidos += pedido.vlTotalFamiliaPago;
						_vlTotalDevolucoesPedidos += pedido.vlTotalFamiliaDevolucaoPrecoNF;
					}
					#endregion

					#region [ Exibe valor total do pedido ]
					atualizaExibicaoValorTotalPedidoSelecionado();
					lblTotalTodosPedidos.Text = Global.formataMoeda(_vlTotalPrecoNfPedidos);
					lblTotalValorPagoTodosPedidos.Text = Global.formataMoeda(_vlTotalFamiliaPagoPedidos);
					lblTotalDevolucoesTodosPedidos.Text = Global.formataMoeda(_vlTotalDevolucoesPedidos);
					lblTotalDevolucoesTodosPedidos.ForeColor = (_vlTotalDevolucoesPedidos == 0 ? Color.Black : Color.Red);
					#endregion

					#region [ Preenchimento dos campos ]

					#region [ Combo Cedente ]
					cbBoletoCedente.ValueMember = "id";
					cbBoletoCedente.DisplayMember = "descricao_formatada";
					cbBoletoCedente.DataSource = ComboDAO.criaDtbBoletoCedenteCombo(ComboDAO.eFiltraStAtivo.SOMENTE_ATIVOS);
					if (!comboBoletoCedentePosicionaDefault()) cbBoletoCedente.SelectedIndex = -1;
					// Se houver apenas 1 opção, então seleciona
					if ((cbBoletoCedente.Items.Count == 1) && (cbBoletoCedente.SelectedIndex == -1)) cbBoletoCedente.SelectedIndex = 0;
					#endregion

					#region [ Nº NF ]
					if (boletoPreCadastradoSelecionado.numero_NF == 0)
						txtNumeroNF.Text = "";
					else
						txtNumeroNF.Text = Global.formataInteiro(boletoPreCadastradoSelecionado.numero_NF);
					#endregion

					#region [ Dados do cliente ]
					strEndereco = clienteSelecionado.endereco;
					if (clienteSelecionado.endereco_numero.Length > 0) strEndereco += ", " + clienteSelecionado.endereco_numero;
					if (clienteSelecionado.endereco_complemento.Length > 0) strEndereco += " " + clienteSelecionado.endereco_complemento;

					txtClienteNome.Text = clienteSelecionado.nome.ToUpper();
					txtClienteCnpjCpf.Text = Global.formataCnpjCpf(clienteSelecionado.cnpj_cpf);
					txtEndereco.Text = strEndereco.ToUpper();
					txtCep.Text = Global.formataCep(clienteSelecionado.cep);
					txtBairro.Text = clienteSelecionado.bairro.ToUpper();
					txtCidade.Text = clienteSelecionado.cidade.ToUpper();
					txtUF.Text = clienteSelecionado.uf.ToUpper();
					txtEmail.Text = clienteSelecionado.email.ToLower();
					#endregion

					#region [ Dados das parcelas impressas na NF ]
					if (boletoPreCadastradoSelecionado.listaItem.Count > 0) grdParcelasNF.Rows.Add(boletoPreCadastradoSelecionado.listaItem.Count);
					intIndiceLinhaGrid = 0;
					for (int i = 0; i < boletoPreCadastradoSelecionado.listaItem.Count; i++)
					{
						strDadosRateio = "";
						for (int j = 0; j < boletoPreCadastradoSelecionado.listaItem[i].listaRateio.Count; j++)
						{
							strDadosRateioParcela = boletoPreCadastradoSelecionado.listaItem[i].listaRateio[j].pedido + "=" + Global.formataMoeda(boletoPreCadastradoSelecionado.listaItem[i].listaRateio[j].valor);
							if (strDadosRateio.Length > 0) strDadosRateio += "|";
							strDadosRateio += strDadosRateioParcela;
						}
						grdParcelasNF.Rows[intIndiceLinhaGrid].Cells["grdParcelasNF_num_parcela"].Value = (intIndiceLinhaGrid + 1).ToString() + " / " + boletoPreCadastradoSelecionado.listaItem.Count.ToString();
						grdParcelasNF.Rows[intIndiceLinhaGrid].Cells["grdParcelasNF_forma_pagto"].Value = Global.formaPagtoPedidoDescricao(boletoPreCadastradoSelecionado.listaItem[i].forma_pagto);
						grdParcelasNF.Rows[intIndiceLinhaGrid].Cells["grdParcelasNF_dt_vencto"].Value = Global.formataDataDdMmYyyyComSeparador(boletoPreCadastradoSelecionado.listaItem[i].dt_vencto);
						grdParcelasNF.Rows[intIndiceLinhaGrid].Cells["grdParcelasNF_valor"].Value = Global.formataMoeda(boletoPreCadastradoSelecionado.listaItem[i].valor);
						grdParcelasNF.Rows[intIndiceLinhaGrid].Cells["grdParcelasNF_dados_rateio"].Value = strDadosRateio;
						vlTotalParcelasNF += boletoPreCadastradoSelecionado.listaItem[i].valor;
						intIndiceLinhaGrid++;
					}

					ajustaPosicaoLblTotalGridParcelasNF();
					lblTotalGridParcelasNF.Text = Global.formataMoeda(vlTotalParcelasNF);

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < grdParcelasNF.Rows.Count; i++)
					{
						if (grdParcelasNF.Rows[i].Selected) grdParcelasNF.Rows[i].Selected = false;
					}
					#endregion

					#endregion

					#region [ Dados das parcelas dos boletos a gerar ]
					intIndiceLinhaGrid = 0;
					for (int i = 0; i < boletoPreCadastradoSelecionado.listaItem.Count; i++)
					{
						if (boletoPreCadastradoSelecionado.listaItem[i].forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
						{
							intQtdeParcelasBoleto++;
						}
					}

					for (int i = 0; i < boletoPreCadastradoSelecionado.listaItem.Count; i++)
					{
						if (boletoPreCadastradoSelecionado.listaItem[i].forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
						{
							grdParcelas.Rows.Add();
							strDadosRateio = "";
							for (int j = 0; j < boletoPreCadastradoSelecionado.listaItem[i].listaRateio.Count; j++)
							{
								strDadosRateioParcela = boletoPreCadastradoSelecionado.listaItem[i].listaRateio[j].pedido + "=" + Global.formataMoeda(boletoPreCadastradoSelecionado.listaItem[i].listaRateio[j].valor);
								if (strDadosRateio.Length > 0) strDadosRateio += "|";
								strDadosRateio += strDadosRateioParcela;
							}
							grdParcelas.Rows[intIndiceLinhaGrid].Cells["grdParcelas_num_parcela"].Value = (intIndiceLinhaGrid + 1).ToString() + " / " + intQtdeParcelasBoleto;
							grdParcelas.Rows[intIndiceLinhaGrid].Cells["grdParcelas_dt_vencto"].Value = Global.formataDataDdMmYyyyComSeparador(boletoPreCadastradoSelecionado.listaItem[i].dt_vencto);
							grdParcelas.Rows[intIndiceLinhaGrid].Cells["grdParcelas_valor"].Value = Global.formataMoeda(boletoPreCadastradoSelecionado.listaItem[i].valor);
							grdParcelas.Rows[intIndiceLinhaGrid].Cells["grdParcelas_dados_rateio"].Value = strDadosRateio;
							vlTotalParcelasBoleto += boletoPreCadastradoSelecionado.listaItem[i].valor;
							intIndiceLinhaGrid++;
						}
					}

					ajustaPosicaoLblTotalGridParcelas();
					lblTotalGridParcelas.Text = Global.formataMoeda(vlTotalParcelasBoleto);

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < grdParcelas.Rows.Count; i++)
					{
						if (grdParcelas.Rows[i].Selected) grdParcelas.Rows[i].Selected = false;
					}
					#endregion

					#endregion

					#endregion

					_InicializacaoOk = true;
				}
				#endregion
			}
			catch (Exception ex)
			{
				_OcorreuExceptionNaInicializacao = true;
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

		#region [ FBoletoCadastraDetalhe_FormClosing ]
		private void FBoletoCadastraDetalhe_FormClosing(object sender, FormClosingEventArgs e)
		{
			#region [ Declarações ]
			Boleto boletoEditado;
			#endregion

			try
			{
				#region [ Trata situação em que os dados do boleto foram anulados ]
				if (_blnRegistroFoiAnulado)
				{
					// Aciona evento para refazer a pesquisa e atualizar os dados do grid
					if (evtRegistroAnulado != null) evtRegistroAnulado();
					return;
				}
				#endregion

				#region [ Trata situação em que o boleto foi cadastrado ]
				if (_blnRegistroFoiGravado)
				{
					// Aciona evento para refazer a pesquisa e atualizar os dados do grid
					if (evtRegistroGravado != null) evtRegistroGravado();
					return;
				}
				#endregion

				#region [ Verifica se houve alterações ]
				boletoEditado = obtemDadosBoletoCamposTela();
				if (boletoEditado != null)
				{
					if (isBoletoEditado(boletoPreCadastradoSelecionado, boletoEditado))
					{
						if (!confirma("As alterações serão perdidas!!\nContinua assim mesmo?"))
						{
							e.Cancel = true;
							return;
						}
					}
				}
				#endregion
			}
			finally
			{
				#region [ Torna visível o form chamador? ]
				if (!e.Cancel)
				{
					if (_formChamador != null)
					{
						_formChamador.Location = this.Location;
						_formChamador.Visible = true;
						this.Visible = false;
					}
				}
				#endregion
			}
		}
		#endregion

		#endregion

		#region [ cbBoletoCedente ]

		#region [ cbBoletoCedente_KeyDown ]
		private void cbBoletoCedente_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataComboBoxKeyDown(sender, e, txtClienteNome);
		}
		#endregion

		#region [ cbBoletoCedente_SelectedValueChanged ]
		private void cbBoletoCedente_SelectedValueChanged(object sender, EventArgs e)
		{
			trataSelecaoBoletoCedente();
		}
		#endregion

		#endregion

		#region [ txtNumeroNF ]

		#region [ txtNumeroNF_Enter ]
		private void txtNumeroNF_Enter(object sender, EventArgs e)
		{
			btnDummy.Focus();
		}
		#endregion

		#endregion

		#region [ txtClienteNome ]

		#region [ txtClienteNome_Enter ]
		private void txtClienteNome_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtClienteNome_Leave ]
		private void txtClienteNome_Leave(object sender, EventArgs e)
		{
			txtClienteNome.Text = txtClienteNome.Text.Trim();
		}
		#endregion

		#region [ txtClienteNome_KeyDown ]
		private void txtClienteNome_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtEndereco);
		}
		#endregion

		#region [ txtClienteNome_KeyPress ]
		private void txtClienteNome_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#region [ txtClienteNome_TextChanged ]
		private void txtClienteNome_TextChanged(object sender, EventArgs e)
		{
			int intTamanho;
			intTamanho = Global.Cte.Etc.MAX_TAM_BOLETO_CAMPO_NOME_SACADO - txtClienteNome.Text.Length;
			lblClienteNomeTamanhoRestante.Text = "(" + intTamanho.ToString() + ")";
			if (intTamanho > 0)
				lblClienteNomeTamanhoRestante.ForeColor = Color.DarkGreen;
			else if (intTamanho < 0)
				lblClienteNomeTamanhoRestante.ForeColor = Color.DarkRed;
			else
				lblClienteNomeTamanhoRestante.ForeColor = Color.DimGray;
		}
		#endregion

		#endregion

		#region [ txtClienteCnpjCpf ]

		#region [ txtClienteCnpjCpf_Enter ]
		private void txtClienteCnpjCpf_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtClienteCnpjCpf_Leave ]
		private void txtClienteCnpjCpf_Leave(object sender, EventArgs e)
		{
			if (txtClienteCnpjCpf.Text.Length == 0) return;
			txtClienteCnpjCpf.Text = Global.formataCnpjCpf(txtClienteCnpjCpf.Text);
			if (!Global.isCnpjCpfOk(txtClienteCnpjCpf.Text))
			{
				avisoErro("CNPJ/CPF inválido!!");
				txtClienteCnpjCpf.Focus();
				return;
			}
		}
		#endregion

		#region [ txtClienteCnpjCpf_KeyDown ]
		private void txtClienteCnpjCpf_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtBairro);
		}
		#endregion

		#region [ txtClienteCnpjCpf_KeyPress ]
		private void txtClienteCnpjCpf_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoCnpjCpf(e.KeyChar);
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
			Global.trataTextBoxKeyDown(sender, e, txtClienteCnpjCpf);
		}
		#endregion

		#region [ txtEndereco_KeyPress ]
		private void txtEndereco_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#region [ txtEndereco_TextChanged ]
		private void txtEndereco_TextChanged(object sender, EventArgs e)
		{
			int intTamanho;
			intTamanho = Global.Cte.Etc.MAX_TAM_BOLETO_CAMPO_ENDERECO - txtEndereco.Text.Length;
			lblEnderecoTamanhoRestante.Text = "(" + intTamanho.ToString() + ")";
			if (intTamanho > 0)
				lblEnderecoTamanhoRestante.ForeColor = Color.DarkGreen;
			else if (intTamanho < 0)
				lblEnderecoTamanhoRestante.ForeColor = Color.DarkRed;
			else
				lblEnderecoTamanhoRestante.ForeColor = Color.DimGray;
		}
		#endregion

		#endregion

		#region [ txtBairro ]

		#region [ txtBairro_Enter ]
		private void txtBairro_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtBairro_Leave ]
		private void txtBairro_Leave(object sender, EventArgs e)
		{
			txtBairro.Text = txtBairro.Text.Trim();
		}
		#endregion

		#region [ txtBairro_KeyDown ]
		private void txtBairro_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtCep);
		}
		#endregion

		#region [ txtBairro_KeyPress ]
		private void txtBairro_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
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
			txtCep.Text = Global.formataCep(txtCep.Text);
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
			Global.trataTextBoxKeyDown(sender, e, txtCidade);
		}
		#endregion

		#region [ txtCep_KeyPress ]
		private void txtCep_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoCep(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtCidade ]

		#region [ txtCidade_Enter ]
		private void txtCidade_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtCidade_Leave ]
		private void txtCidade_Leave(object sender, EventArgs e)
		{
			txtCidade.Text = txtCidade.Text.Trim();
		}
		#endregion

		#region [ txtCidade_KeyDown ]
		private void txtCidade_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtUF);
		}
		#endregion

		#region [ txtCidade_KeyPress ]
		private void txtCidade_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtUF ]

		#region [ txtUF_Enter ]
		private void txtUF_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtUF_Leave ]
		private void txtUF_Leave(object sender, EventArgs e)
		{
			if (txtUF.Text.Length == 0) return;

			txtUF.Text = txtUF.Text.Trim().ToUpper();

			if (!Global.isUfOk(txtUF.Text))
			{
				avisoErro("UF inválida!!");
				txtUF.Focus();
				return;
			}
		}
		#endregion

		#region [ txtUF_KeyDown ]
		private void txtUF_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtEmail);
		}
		#endregion

		#region [ txtUF_KeyPress ]
		private void txtUF_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoSomenteLetras(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtEmail ]

		#region [ txtEmail_Enter ]
		private void txtEmail_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtEmail_Leave ]
		private void txtEmail_Leave(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String strMsgErro = "";
			String strRelacaoEmailInvalido = "";
			String[] v;
			#endregion

			txtEmail.Text = txtEmail.Text.Trim();
			if (txtEmail.Text.Length == 0) return;

			if (!Global.isEmailOk(txtEmail.Text, ref strRelacaoEmailInvalido))
			{
				v = strRelacaoEmailInvalido.Split(' ');
				for (int i = 0; i < v.Length; i++)
				{
					if (strMsgErro.Length > 0) strMsgErro += "\n";
					strMsgErro += v[i];
				}
				if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
				strMsgErro = "E-mail inválido:" + strMsgErro;
				avisoErro(strMsgErro);
				txtEmail.Focus();
				return;
			}
		}
		#endregion

		#region [ txtEmail_KeyDown ]
		private void txtEmail_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtSegundaMensagem);
		}
		#endregion

		#region [ txtEmail_KeyPress ]
		private void txtEmail_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoEmail(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtSegundaMensagem ]

		#region [ txtSegundaMensagem_Enter ]
		private void txtSegundaMensagem_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtSegundaMensagem_Leave ]
		private void txtSegundaMensagem_Leave(object sender, EventArgs e)
		{
			txtSegundaMensagem.Text = txtSegundaMensagem.Text.Trim();
		}
		#endregion

		#region [ txtSegundaMensagem_KeyDown ]
		private void txtSegundaMensagem_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtMensagem1);
		}
		#endregion

		#region [ txtSegundaMensagem_KeyPress ]
		private void txtSegundaMensagem_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtMensagem1 ]

		#region [ txtMensagem1_Enter ]
		private void txtMensagem1_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtMensagem1_Leave ]
		private void txtMensagem1_Leave(object sender, EventArgs e)
		{
			txtMensagem1.Text = txtMensagem1.Text.Trim();
		}
		#endregion

		#region [ txtMensagem1_KeyDown ]
		private void txtMensagem1_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtMensagem2);
		}
		#endregion

		#region [ txtMensagem1_KeyPress ]
		private void txtMensagem1_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtMensagem2 ]

		#region [ txtMensagem2_Enter ]
		private void txtMensagem2_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtMensagem2_Leave ]
		private void txtMensagem2_Leave(object sender, EventArgs e)
		{
			txtMensagem2.Text = txtMensagem2.Text.Trim();
		}
		#endregion

		#region [ txtMensagem2_KeyDown ]
		private void txtMensagem2_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtMensagem3);
		}
		#endregion

		#region [ txtMensagem2_KeyPress ]
		private void txtMensagem2_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtMensagem3 ]

		#region [ txtMensagem3_Enter ]
		private void txtMensagem3_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtMensagem3_Leave ]
		private void txtMensagem3_Leave(object sender, EventArgs e)
		{
			txtMensagem3.Text = txtMensagem3.Text.Trim();
		}
		#endregion

		#region [ txtMensagem3_KeyDown ]
		private void txtMensagem3_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtMensagem4);
		}
		#endregion

		#region [ txtMensagem3_KeyPress ]
		private void txtMensagem3_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtMensagem4 ]

		#region [ txtMensagem4_Enter ]
		private void txtMensagem4_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtMensagem4_Leave ]
		private void txtMensagem4_Leave(object sender, EventArgs e)
		{
			txtMensagem4.Text = txtMensagem4.Text.Trim();
		}
		#endregion

		#region [ txtMensagem4_KeyDown ]
		private void txtMensagem4_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnDummy);
		}
		#endregion

		#region [ txtMensagem4_KeyPress ]
		private void txtMensagem4_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoTexto(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtJurosMora ]

		#region [ txtJurosMora_Enter ]
		private void txtJurosMora_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtJurosMora_Leave ]
		private void txtJurosMora_Leave(object sender, EventArgs e)
		{
			if (Global.converteNumeroDecimal(txtJurosMora.Text) < 0)
			{
				avisoErro("Percentual de juros de mora é inválido!!");
				txtJurosMora.Focus();
				return;
			}
		}
		#endregion

		#region [ txtJurosMora_KeyDown ]
		private void txtJurosMora_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtPercMulta);
		}
		#endregion

		#region [ txtJurosMora_KeyPress ]
		private void txtJurosMora_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoPercentual(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtPercMulta ]

		#region [ txtPercMulta_Enter ]
		private void txtPercMulta_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtPercMulta_Leave ]
		private void txtPercMulta_Leave(object sender, EventArgs e)
		{
			if ((Global.converteNumeroDecimal(txtPercMulta.Text) < 0) || (Global.converteNumeroDecimal(txtPercMulta.Text) >= 100))
			{
				avisoErro("Percentual de multa é inválido!!");
				txtPercMulta.Focus();
				return;
			}
		}
		#endregion

		#region [ txtPercMulta_KeyDown ]
		private void txtPercMulta_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtProtestarApos);
		}
		#endregion

		#region [ txtPercMulta_KeyPress ]
		private void txtPercMulta_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoPercentual(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ txtProtestarApos ]

		#region [ txtProtestarApos_Enter ]
		private void txtProtestarApos_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtProtestarApos_Leave ]
		private void txtProtestarApos_Leave(object sender, EventArgs e)
		{
			txtProtestarApos.Text = txtProtestarApos.Text.Trim();
		}
		#endregion

		#region [ txtProtestarApos_KeyDown ]
		private void txtProtestarApos_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, btnDummy);
		}
		#endregion

		#region [ txtProtestarApos_KeyPress ]
		private void txtProtestarApos_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
		}
		#endregion

		#endregion

		#region [ lbPedido ]

		#region [ lbPedido_DoubleClick ]
		private void lbPedido_DoubleClick(object sender, EventArgs e)
		{
			trataBotaoPedidoConsulta();
		}
		#endregion

		#region [ lbPedido_SelectedIndexChanged ]
		private void lbPedido_SelectedIndexChanged(object sender, EventArgs e)
		{
			atualizaExibicaoValorTotalPedidoSelecionado();
		}
		#endregion

		#endregion

		#region [ grdParcelas ]

		#region [ grdParcelas_RowsAdded ]
		private void grdParcelas_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
		{
			ajustaPosicaoLblTotalGridParcelas();
		}
		#endregion

		#region [ grdParcelas_RowsRemoved ]
		private void grdParcelas_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
		{
			ajustaPosicaoLblTotalGridParcelas();
		}
		#endregion

		#region [ grdParcelas_DoubleClick ]
		private void grdParcelas_DoubleClick(object sender, EventArgs e)
		{
			trataBotaoEditarParcela();
		}
		#endregion

		#region [ grdParcelas_KeyDown ]
		private void grdParcelas_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				e.SuppressKeyPress = true;
				trataBotaoEditarParcela();
				return;
			}
		}
		#endregion

		#endregion

		#region [ Botões/Menu ]

		#region [ Cadastrar ]

		#region [ btnCadastrar_Click ]
		private void btnCadastrar_Click(object sender, EventArgs e)
		{
			trataBotaoCadastrar();
		}
		#endregion

		#region [ menuBoletoCadastrar_Click ]
		private void menuBoletoCadastrar_Click(object sender, EventArgs e)
		{
			trataBotaoCadastrar();
		}
		#endregion

		#endregion

		#region [ Anular ]

		#region [ btnAnular_Click ]
		private void btnAnular_Click(object sender, EventArgs e)
		{
			trataBotaoAnular();
		}
		#endregion

		#region [ menuBoletoAnular_Click ]
		private void menuBoletoAnular_Click(object sender, EventArgs e)
		{
			trataBotaoAnular();
		}
		#endregion

		#endregion

		#region [ Pesquisa CEP ]

		private void btnCepPesquisa_Click(object sender, EventArgs e)
		{
			trataBotaoPesquisaCep();
		}

		#endregion

		#region [ btnEnderecoOriginal ]

		#region [ btnEnderecoOriginal_Click ]
		private void btnEnderecoOriginal_Click(object sender, EventArgs e)
		{
			trataBotaoRestauraEnderecoOriginal();
		}
		#endregion

		#endregion

		#region [ btnPedidoConsulta ]

		#region [ btnPedidoConsulta_Click ]
		private void btnPedidoConsulta_Click(object sender, EventArgs e)
		{
			trataBotaoPedidoConsulta();
		}
		#endregion

		#endregion

		#region [ btnEditarParcela ]

		#region [ btnEditarParcela_Click ]
		private void btnEditarParcela_Click(object sender, EventArgs e)
		{
			trataBotaoEditarParcela();
		}
		#endregion

		#endregion

		#region [ btnExcluirParcela ]

		#region [ btnExcluirParcela_Click ]
		private void btnExcluirParcela_Click(object sender, EventArgs e)
		{
			trataBotaoExcluirParcela();
		}
		#endregion

		#endregion

		#region [ btnAdicionarParcela ]

		#region [ btnAdicionarParcela_Click ]
		private void btnAdicionarParcela_Click(object sender, EventArgs e)
		{
			trataBotaoAdicionarParcela();
		}
		#endregion

		#endregion

		#endregion

		#endregion
	}
}
