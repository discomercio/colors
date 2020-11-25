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
	public partial class FPedido : Financeiro.FModelo
	{
		#region [ Atributos ]
		private Form _formChamador = null;
		
		bool _InicializacaoOk;
		bool _carregandoComboBoxPedido;
		String _ultimoPedidoExibido = "";

		private Cliente _cliente;
		internal Cliente cliente
		{
			get { return _cliente; }
			set { _cliente = value; }
		}
		
		private List<Pedido> _listaPedidos;
		internal List<Pedido> listaPedidos
		{
			get { return _listaPedidos; }
			set { _listaPedidos = value; }
		}

		private String _numeroPedidoDefault;
		public String numeroPedidoDefault
		{
			get { return _numeroPedidoDefault; }
			set { _numeroPedidoDefault = value; }
		}
		#endregion

		#region [ Construtor ]
		public FPedido(Form formChamador)
		{
			InitializeComponent();

			_formChamador = formChamador;
		}
		#endregion

		#region [ Métodos ]

		#region [ limpaCamposPedido ]
		void limpaCamposPedido()
		{
			lblLoja.Text = "";
			lblVendedor.Text = "";
			lblDataPedido.Text = "";
			lblStEntrega.Text = "";
			lblPedidoRecebidoStatus.Text = "";
			lblIndicador.Text = "";

			lblNomeCliente.Text = "";
			lblCnpjCpf.Text = "";
			lblIeRg.Text = "";
			lblEndereco.Text = "";
			lblTelRes.Text = "";
			lblTelCom.Text = "";
			lblContato.Text = "";
			lblEmail.Text = "";

			grdItem.Rows.Clear();
			lblStatusPagto.Text = "";
			lblTotalPagoFamilia.Text = "";
			lblTotalDevolucoes.Text = "";
			lblRaBruto.Text = "";
			lblTotalVenda.Text = "";
			lblTotalNf.Text = "";
			lblValorTotal.Text = "";

			txtObs1.Text = "";
			lblObs2.Text = "";
			txtDescricaoFormaPagto.Text = "";

			grdDevolucao.Rows.Clear();

			txtFormaPagto.Text = "";
		}
		#endregion

		#region [ trataSelecaoPedido ]
		void trataSelecaoPedido()
		{
			if (_carregandoComboBoxPedido) return;
			if (cbPedido.SelectedIndex > -1) exibeDadosPedido(cbPedido.Items[cbPedido.SelectedIndex].ToString());
		}
		#endregion

		#region [ exibeDadosPedido ]
		void exibeDadosPedido(String strNumeroPedido)
		{
			#region [ Declarações ]
			Pedido pedido = null;
			bool blnAchou = false;
			int intIndiceGrid;
			decimal vlTotalVenda;
			decimal vlTotalNF;
			int intLarguraVScrollBar;
			#endregion

			#region [ Consistência ]
			if (strNumeroPedido == null) {
				limpaCamposPedido();
				return;
			}

			if (strNumeroPedido.Trim().Length == 0)
			{
				limpaCamposPedido();
				return;
			}
			#endregion

			#region [ Os dados exibidos já são deste pedido? ]
			if (_ultimoPedidoExibido.Equals(strNumeroPedido)) return;
			#endregion

			#region [ Limpa os campos ]
			limpaCamposPedido();
			#endregion

			#region [ Localiza o pedido na lista ]
			for (int i = 0; i < _listaPedidos.Count; i++)
			{
				pedido = _listaPedidos[i];
				if (pedido.pedido.Equals(strNumeroPedido))
				{
					blnAchou = true;
					break;
				}
			}

			if (!blnAchou)
			{
				avisoErro("Pedido " + strNumeroPedido + " não encontrado nos dados lidos do BD!!");
				return;
			}
			#endregion

			#region [ Preenche as informações do pedido ]

			#region [ Loja / Status do pedido ]
			lblLoja.Text = pedido.loja + " - " + (pedido.loja_razao_social.Trim().Length > 0 ? pedido.loja_razao_social.Trim() : pedido.loja_nome);
			lblVendedor.Text = pedido.vendedor_nome + " (" + pedido.vendedor + ")";
			lblDataPedido.Text = Global.formataDataDdMmYyyyComSeparador(pedido.data);
			lblStEntrega.ForeColor = Global.stEntregaPedidoCor(pedido.st_entrega);
			lblStEntrega.Text = Global.stEntregaPedidoDescricao(pedido.st_entrega).ToUpper();
			if (pedido.st_entrega.Equals(Global.Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE))
			{
				lblStEntrega.Text += " (" + Global.formataDataDdMmYyyyComSeparador(pedido.entregue_data) + ")";
				if (pedido.pedidoRecebidoStatus == Global.Cte.StPedidoRecebido.COD_ST_PEDIDO_RECEBIDO_SIM)
				{
					lblPedidoRecebidoStatus.Text = "Recebido (".ToUpper() + Global.formataDataDdMmYyyyComSeparador(pedido.pedidoRecebidoData) + ")";
				}
			}
			else if (pedido.st_entrega.Equals(Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO))
			{
				lblStEntrega.Text += " (" + Global.formataDataDdMmYyyyComSeparador(pedido.cancelado_data) + ")";
			}
			
			lblIndicador.Text = pedido.indicador;
			if (pedido.indicador.Trim().Length > 0)
			{
				if (pedido.indicador_desempenho_nota.Trim().Length > 0) lblIndicador.Text += " (" + pedido.indicador_desempenho_nota + ")";
			}
			#endregion

			#endregion

			#region [ Dados do cliente ]
			if (pedido.st_memorizacao_completa_enderecos != 0)
			{
				#region [ Prioridade em usar dados do endereço memorizado no pedido, se houver ]
				lblNomeCliente.Text = pedido.endereco_nome;
				if (pedido.endereco_tipo_pessoa.Equals(Global.Cte.Etc.ID_PF))
				{
					lblTitCnpjCpf.Text = "CPF";
					lblTitIeRg.Text = "RG";
					lblIeRg.Text = pedido.endereco_rg;
				}
				else
				{
					lblTitCnpjCpf.Text = "CNPJ";
					lblTitIeRg.Text = "I.E.";
					if ((pedido.endereco_contribuinte_icms_status == Global.Cte.StClienteContribuinteIcmsStatus.CONTRIBUINTE_ICMS_SIM)
						|| (pedido.endereco_contribuinte_icms_status == Global.Cte.StClienteContribuinteIcmsStatus.CONTRIBUINTE_ICMS_NAO))
					{
						lblIeRg.Text = pedido.endereco_ie;
					}
				}
				lblCnpjCpf.Text = Global.formataCnpjCpf(pedido.endereco_cnpj_cpf);
				lblEndereco.Text = Global.formataEndereco(pedido.endereco_logradouro, pedido.endereco_numero, pedido.endereco_complemento, pedido.endereco_bairro, pedido.endereco_cidade, pedido.endereco_uf, pedido.endereco_cep);
				lblTelRes.Text = Global.formataTelefone(pedido.endereco_ddd_res, pedido.endereco_tel_res);
				lblTelCom.Text = Global.formataTelefone(pedido.endereco_ddd_com, pedido.endereco_tel_com, pedido.endereco_ramal_com);
				lblContato.Text = pedido.endereco_contato;
				lblEmail.Text = pedido.endereco_email;
				#endregion
			}
			else
			{
				#region [ Se não houver dados do endereço memorizado no pedido, usa dados do cadastro do cliente ]
				lblNomeCliente.Text = cliente.nome;
				if (cliente.tipo.Equals(Global.Cte.Etc.ID_PF))
				{
					lblTitCnpjCpf.Text = "CPF";
					lblTitIeRg.Text = "RG";
					lblIeRg.Text = cliente.rg;
				}
				else
				{
					lblTitCnpjCpf.Text = "CNPJ";
					lblTitIeRg.Text = "I.E.";
					lblIeRg.Text = cliente.ie;
				}
				lblCnpjCpf.Text = Global.formataCnpjCpf(cliente.cnpj_cpf);
				lblEndereco.Text = Global.formataEndereco(cliente.endereco, cliente.endereco_numero, cliente.endereco_complemento, cliente.bairro, cliente.cidade, cliente.uf, cliente.cep);
				lblTelRes.Text = Global.formataTelefone(cliente.ddd_res, cliente.tel_res);
				lblTelCom.Text = Global.formataTelefone(cliente.ddd_com, cliente.tel_com, cliente.ramal_com);
				lblContato.Text = cliente.contato;
				lblEmail.Text = cliente.email;
				#endregion
			}
			#endregion

			#region [ Itens do pedido ]

			#region [ Preenche linhas do grid ]
			vlTotalNF = 0;
			vlTotalVenda = 0;
			intIndiceGrid = 0;
			if (pedido.listaPedidoItem.Count > 0) grdItem.Rows.Add(pedido.listaPedidoItem.Count);
			foreach (PedidoItem item in pedido.listaPedidoItem)
			{
				grdItem.Rows[intIndiceGrid].Cells["fabricante"].Value = item.fabricante;
				grdItem.Rows[intIndiceGrid].Cells["produto"].Value = item.produto;
				grdItem.Rows[intIndiceGrid].Cells["descricao"].Value = item.descricao;
				grdItem.Rows[intIndiceGrid].Cells["qtde"].Value = item.qtde;
				grdItem.Rows[intIndiceGrid].Cells["preco_venda"].Value = Global.formataMoeda(item.preco_venda);
				grdItem.Rows[intIndiceGrid].Cells["preco_NF"].Value = Global.formataMoeda(item.preco_NF);
				grdItem.Rows[intIndiceGrid].Cells["vl_total"].Value = Global.formataMoeda(item.qtde * item.preco_NF);
				vlTotalVenda += item.qtde * item.preco_venda;
				vlTotalNF += item.qtde * item.preco_NF;
				intIndiceGrid++;
			}
			#endregion

			#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
			for (int i = 0; i < grdItem.Rows.Count; i++)
			{
				if (grdItem.Rows[i].Selected) grdItem.Rows[i].Selected = false;
			}
			#endregion

			lblStatusPagto.Text = Global.stPagtoPedidoDescricao(pedido.st_pagto).ToUpper();
			lblStatusPagto.ForeColor = Global.stPagtoPedidoCor(pedido.st_pagto);
			lblTotalPagoFamilia.Text = Global.formataMoeda(pedido.vlTotalFamiliaPago);
			lblTotalDevolucoes.Text = Global.formataMoeda(pedido.vlTotalFamiliaDevolucaoPrecoNF);
			if (pedido.vlTotalFamiliaDevolucaoPrecoNF == 0)
				lblTotalDevolucoes.ForeColor = Color.Black;
			else
				lblTotalDevolucoes.ForeColor = Color.Red;
			lblRaBruto.Text = Global.formataMoeda(pedido.vlTotalFamiliaPrecoNF - pedido.vlTotalFamiliaPrecoVenda);
			lblTotalVenda.Text = Global.formataMoeda(vlTotalVenda);
			lblTotalNf.Text = Global.formataMoeda(vlTotalNF);
			lblValorTotal.Text = Global.formataMoeda(vlTotalNF);

			if (Global.isVScrollBarVisible(grdItem))
			{
				intLarguraVScrollBar = Global.getVScrollBarWidth(grdItem);
				lblTitTotalVenda.Left -= intLarguraVScrollBar;
				lblTotalVenda.Left -= intLarguraVScrollBar;
				lblTitTotalNf.Left -= intLarguraVScrollBar;
				lblTotalNf.Left -= intLarguraVScrollBar;
				lblTitValorTotal.Left -= intLarguraVScrollBar;
				lblValorTotal.Left -= intLarguraVScrollBar;
			}
			#endregion

			#region [ Obs I / Obs II / Descrição da forma de pagamento ]
			txtObs1.Text = pedido.obs_1;
			lblObs2.Text = pedido.obs_2;
			txtDescricaoFormaPagto.Text = pedido.forma_pagto;
			#endregion

			#region [ Devoluções ]

			#region [ Preenche linhas do grid ]
			intIndiceGrid = 0;
			if (pedido.listaPedidoItemDevolvido.Count > 0) grdDevolucao.Rows.Add(pedido.listaPedidoItemDevolvido.Count);
			foreach (PedidoItemDevolvido itemDevolvido in pedido.listaPedidoItemDevolvido)
			{
				grdDevolucao.Rows[intIndiceGrid].Cells["produto_devolucao"].Value = itemDevolvido.produto;
				grdDevolucao.Rows[intIndiceGrid].Cells["data_devolucao"].Value = Global.formataDataDdMmYyyyComSeparador(itemDevolvido.devolucao_data);
				grdDevolucao.Rows[intIndiceGrid].Cells["motivo"].Value = itemDevolvido.motivo;
				grdDevolucao.Rows[intIndiceGrid].Cells["qtde_devolucao"].Value = itemDevolvido.qtde;
				grdDevolucao.Rows[intIndiceGrid].Cells["preco_NF_devolucao"].Value = Global.formataMoeda(itemDevolvido.preco_NF);
				intIndiceGrid++;
			}
			#endregion

			#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
			for (int i = 0; i < grdDevolucao.Rows.Count; i++)
			{
				if (grdDevolucao.Rows[i].Selected) grdDevolucao.Rows[i].Selected = false;
			}
			#endregion

			#endregion

			#region [ Descrição do tipo de parcelamento do pedido ]
			txtFormaPagto.Text = Global.retornaDescricaoTipoParcelamentoPedido(pedido);
			#endregion

			_ultimoPedidoExibido = strNumeroPedido;
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FPedido ]

		#region [ FPedido_Load ]
		private void FPedido_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				limpaCamposPedido();

				#region [ Preenche dados ]

				#region [ Lista dos pedidos ]
				_carregandoComboBoxPedido = true;
				try
				{
					cbPedido.Items.Clear();
					if (listaPedidos != null)
					{
						for (int i = 0; i < listaPedidos.Count; i++)
						{
							cbPedido.Items.Add(listaPedidos[i].pedido.ToString());
						}
					}
					// Posiciona o item default
					if (numeroPedidoDefault != null)
					{
						for (int i = 0; i < cbPedido.Items.Count; i++)
						{
							if (cbPedido.Items[i].ToString().Equals(numeroPedidoDefault))
							{
								cbPedido.SelectedIndex = i;
								break;
							}
						}
					}
					if (cbPedido.SelectedIndex > -1) exibeDadosPedido(cbPedido.Items[cbPedido.SelectedIndex].ToString());
				}
				finally
				{
					_carregandoComboBoxPedido = false;
				}
				#endregion

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

		#region [ FPedido_Shown ]
		private void FPedido_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					btnDummy.Focus();

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

		#region [ FPedido_FormClosing ]
		private void FPedido_FormClosing(object sender, FormClosingEventArgs e)
		{
			try
			{
				// NOP
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

		#region [ cbPedido ]

		#region [ cbPedido_SelectedIndexChanged ]
		private void cbPedido_SelectedIndexChanged(object sender, EventArgs e)
		{
			trataSelecaoPedido();
		}
		#endregion

		#endregion

		#endregion
	}
}
