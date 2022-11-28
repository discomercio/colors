#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Media;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	public partial class FBoletoAvulsoComPedidoCadDetalhe : Financeiro.FModelo
	{
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

		private List<String> _listaNumeroPedidoSelecionado = new List<String>();
		public List<String> listaNumeroPedidoSelecionado
		{
			get { return _listaNumeroPedidoSelecionado; }
		}

		private List<Pedido> _listaPedidos = new List<Pedido>();

		private int _numeroNF = 0;
		private int _numeroDocumentoBoletoAvulso = 0;

		private decimal _vlTotalFamiliaPagoPedidos = 0;
		private decimal _vlTotalDevolucoesPedidos = 0;
		private decimal _vlTotalPrecoNfPedidos = 0;
		private decimal _vlTotalBoletoPedidos = 0;
		private decimal _vlTotalFormaPagtoPedidos = 0;

		BoletoAvulsoComPedido boletoAvulsoComPedidoSelecionado;
		Cliente clienteSelecionado;
		BoletoCliente boletoCliente;
		BoletoCedente boletoCedenteSelecionado;

		FCepPesquisa fCepPesquisa;
		FPedido fPedido;
		FBoletoParcelaEdita fBoletoParcelaEdita;
		#endregion

		#region [ Construtor ]
		public FBoletoAvulsoComPedidoCadDetalhe(Form formChamador, List<String> listaNumeroPedidoSelecionado)
		{
			InitializeComponent();

			_formChamador = formChamador;
			_listaNumeroPedidoSelecionado = listaNumeroPedidoSelecionado;
		}
		#endregion

		#region [ Métodos ]

		#region [ calculaDataPrimeiroBoleto ]
		private DateTime calculaDataPrimeiroBoleto(int intPrazoEmissaoPrimeiroBoleto)
		{
			DateTime dtResposta;

			if (intPrazoEmissaoPrimeiroBoleto <= 29)
			{
				dtResposta = DateTime.Today.AddDays(30);
			}
			else
			{
				dtResposta = DateTime.Today.AddDays(intPrazoEmissaoPrimeiroBoleto + 7);
			}

			return dtResposta;
		}
		#endregion

		#region [ geraDadosParcelasPagto ]
		/// <summary>
		/// Dada uma lista com os números de pedidos, analisa a forma de pagamento definida e monta uma
		/// lista com todas as parcelas de pagamento, independentemente se a parcela é por boleto ou não.
		/// Os pedidos devem ser todos de um mesmo cliente e precisam ter a mesma forma de pagamento
		/// definida (à vista, parcela única, parcelado com entrada, parcelado sem entrada) e os mesmos
		/// prazos e quantidades de parcelas, quando isso se aplicar.
		/// </summary>
		/// <param name="listaNumeroPedido">
		/// Lista com os números de pedido que serão cobrados juntos nesta série de boletos
		/// </param>
		/// <returns>
		/// Retorna uma lista com todas as parcelas de pagamento definidas no(s) pedido(s), 
		/// independentemente se a parcela é por boleto ou não.
		/// </returns>
		private List<TipoLinhaDadosParcelaPagto> geraDadosParcelasPagto(List<String> listaNumeroPedido, out string msgSolicitacaoConfirmacao)
		{
			#region [ Declarações ]
			String strMsgErro = "";
			String strPedido;
			String strSql;
			String strWhere;
			String strListaPedidosPagtoBoleto = "";
			String strListaPedidosPagtoNaoBoleto = "";
			bool blnPagtoPorBoleto;
			int intQtdeTotalPedidos = 0;
			int intQtdePedidosPagtoBoleto = 0;
			int intQtdePlanoContas = 0;
			int intQtdeTotalParcelas = 0;
			short tipoParcelamento;
			decimal vlTotalPedido = 0;
			decimal vlTotalFormaPagto = 0;
			decimal vlDiferencaArredondamento;
			decimal vlDiferencaArredondamentoRestante;
			decimal vlRateio;
			DateTime dtUltimoPagtoCalculado;
			List<TipoPedidoCalculoParcelasBoleto> vPedidoCalculoParcelas = new List<TipoPedidoCalculoParcelasBoleto>();
			TipoPedidoCalculoParcelasBoleto itemPedidoCalculoParcelas;
			List<TipoLinhaDadosParcelaPagto> vParcelaPagto = new List<TipoLinhaDadosParcelaPagto>();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataTable dtbAux = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Inicialização ]
			msgSolicitacaoConfirmacao = "";
			vParcelaPagto.Add(new TipoLinhaDadosParcelaPagto());
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Para cada pedido, obtém os dados no BD ]
			for (int i = 0; i < listaNumeroPedido.Count; i++)
			{
				strPedido = Global.normalizaNumeroPedido(listaNumeroPedido[i].ToString().Trim());
				if (strPedido.Length > 0)
				{
					#region [ Executa a consulta no BD ]
					strSql =
						"SELECT" +
							" t_PEDIDO__BASE.tipo_parcelamento," +
							" t_PEDIDO__BASE.av_forma_pagto," +
							" t_PEDIDO__BASE.pc_qtde_parcelas," +
							" t_PEDIDO__BASE.pc_valor_parcela," +
							" t_PEDIDO__BASE.pc_maquineta_qtde_parcelas," +
							" t_PEDIDO__BASE.pc_maquineta_valor_parcela," +
							" t_PEDIDO__BASE.pce_forma_pagto_entrada," +
							" t_PEDIDO__BASE.pce_forma_pagto_prestacao," +
							" t_PEDIDO__BASE.pce_entrada_valor," +
							" t_PEDIDO__BASE.pce_prestacao_qtde," +
							" t_PEDIDO__BASE.pce_prestacao_valor," +
							" t_PEDIDO__BASE.pce_prestacao_periodo," +
							" t_PEDIDO__BASE.pse_forma_pagto_prim_prest," +
							" t_PEDIDO__BASE.pse_forma_pagto_demais_prest," +
							" t_PEDIDO__BASE.pse_prim_prest_valor," +
							" t_PEDIDO__BASE.pse_prim_prest_apos," +
							" t_PEDIDO__BASE.pse_demais_prest_qtde," +
							" t_PEDIDO__BASE.pse_demais_prest_valor," +
							" t_PEDIDO__BASE.pse_demais_prest_periodo," +
							" t_PEDIDO__BASE.pu_forma_pagto," +
							" t_PEDIDO__BASE.pu_valor," +
							" t_PEDIDO__BASE.pu_vencto_apos" +
						" FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" +
							" ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" +
						" WHERE" +
							" (t_PEDIDO.pedido = '" + strPedido + "')";
					cmCommand.CommandText = strSql;
					dtbResultado.Reset();
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						if (strMsgErro.Length > 0) strMsgErro += "\n";
						strMsgErro += "Pedido " + strPedido + " não está cadastrado!!";
					}
					else
					{
						intQtdeTotalPedidos++;

						rowResultado = dtbResultado.Rows[0];
						tipoParcelamento = BD.readToShort(rowResultado["tipo_parcelamento"]);

						#region [ Analisa se a forma de pagamento envolve boletos ]
						blnPagtoPorBoleto = false;
						if (tipoParcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
						{
							if ((BD.readToShort(rowResultado["av_forma_pagto"]) == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
								|| (BD.readToShort(rowResultado["av_forma_pagto"]) == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO_AV)) blnPagtoPorBoleto = true;
						}
						else if (tipoParcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
						{
							if (BD.readToShort(rowResultado["pce_forma_pagto_entrada"]) == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO) blnPagtoPorBoleto = true;
							if (BD.readToShort(rowResultado["pce_forma_pagto_prestacao"]) == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO) blnPagtoPorBoleto = true;
						}
						else if (tipoParcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
						{
							if (BD.readToShort(rowResultado["pse_forma_pagto_prim_prest"]) == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO) blnPagtoPorBoleto = true;
							if (BD.readToShort(rowResultado["pse_forma_pagto_demais_prest"]) == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO) blnPagtoPorBoleto = true;
						}
						else if (tipoParcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
						{
							if (BD.readToShort(rowResultado["pu_forma_pagto"]) == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO) blnPagtoPorBoleto = true;
						}
						#endregion

						#region [ Armazena os dados da forma de pagamento ]
						itemPedidoCalculoParcelas = new TipoPedidoCalculoParcelasBoleto();
						itemPedidoCalculoParcelas.pedido = strPedido;
						itemPedidoCalculoParcelas.tipo_parcelamento = BD.readToShort(rowResultado["tipo_parcelamento"]);
						itemPedidoCalculoParcelas.av_forma_pagto = BD.readToShort(rowResultado["av_forma_pagto"]);
						itemPedidoCalculoParcelas.pu_forma_pagto = BD.readToShort(rowResultado["pu_forma_pagto"]);
						itemPedidoCalculoParcelas.pu_valor = BD.readToDecimal(rowResultado["pu_valor"]);
						itemPedidoCalculoParcelas.pu_vencto_apos = BD.readToShort(rowResultado["pu_vencto_apos"]);
						itemPedidoCalculoParcelas.pc_qtde_parcelas = BD.readToShort(rowResultado["pc_qtde_parcelas"]);
						itemPedidoCalculoParcelas.pc_valor_parcela = BD.readToDecimal(rowResultado["pc_valor_parcela"]);
						itemPedidoCalculoParcelas.pc_maquineta_qtde_parcelas = BD.readToShort(rowResultado["pc_maquineta_qtde_parcelas"]);
						itemPedidoCalculoParcelas.pc_maquineta_valor_parcela = BD.readToDecimal(rowResultado["pc_maquineta_valor_parcela"]);
						itemPedidoCalculoParcelas.pce_forma_pagto_entrada = BD.readToShort(rowResultado["pce_forma_pagto_entrada"]);
						itemPedidoCalculoParcelas.pce_forma_pagto_prestacao = BD.readToShort(rowResultado["pce_forma_pagto_prestacao"]);
						itemPedidoCalculoParcelas.pce_entrada_valor = BD.readToDecimal(rowResultado["pce_entrada_valor"]);
						itemPedidoCalculoParcelas.pce_prestacao_qtde = BD.readToShort(rowResultado["pce_prestacao_qtde"]);
						itemPedidoCalculoParcelas.pce_prestacao_valor = BD.readToDecimal(rowResultado["pce_prestacao_valor"]);
						itemPedidoCalculoParcelas.pce_prestacao_periodo = BD.readToShort(rowResultado["pce_prestacao_periodo"]);
						itemPedidoCalculoParcelas.pse_forma_pagto_prim_prest = BD.readToShort(rowResultado["pse_forma_pagto_prim_prest"]);
						itemPedidoCalculoParcelas.pse_forma_pagto_demais_prest = BD.readToShort(rowResultado["pse_forma_pagto_demais_prest"]);
						itemPedidoCalculoParcelas.pse_prim_prest_valor = BD.readToDecimal(rowResultado["pse_prim_prest_valor"]);
						itemPedidoCalculoParcelas.pse_prim_prest_apos = BD.readToShort(rowResultado["pse_prim_prest_apos"]);
						itemPedidoCalculoParcelas.pse_demais_prest_qtde = BD.readToShort(rowResultado["pse_demais_prest_qtde"]);
						itemPedidoCalculoParcelas.pse_demais_prest_valor = BD.readToDecimal(rowResultado["pse_demais_prest_valor"]);
						itemPedidoCalculoParcelas.pse_demais_prest_periodo = BD.readToShort(rowResultado["pse_demais_prest_periodo"]);
						#endregion

						#region [ Calcula o valor total deste pedido ]

						#region [ Monta o Select interno ]
						strSql =
							"SELECT" +
								" p.pedido," +
								" Coalesce(Sum(qtde*preco_NF),0) AS vl_total" +
							" FROM t_PEDIDO p" +
								" INNER JOIN t_PEDIDO_ITEM i ON (p.pedido=i.pedido)" +
							" WHERE" +
								" (p.pedido = '" + strPedido + "')" +
							" GROUP BY" +
								" p.pedido" +
							" UNION " +
							" SELECT" +
								" p.pedido," +
								" -1*Coalesce(Sum(qtde*preco_NF),0) AS vl_total" +
							" FROM t_PEDIDO p" +
								" INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO id ON (p.pedido=id.pedido)" +
							" WHERE" +
								" (p.pedido = '" + strPedido + "')" +
							" GROUP BY" +
								" p.pedido";
						#endregion

						#region [ Monta o Select externo ]
						strSql =
							"SELECT" +
								" pedido," +
								" Sum(vl_total) AS vl_total" +
							" FROM" +
								"(" +
									strSql +
								") t" +
							" GROUP BY" +
								" pedido";
						#endregion

						cmCommand.CommandText = strSql;
						dtbAux.Reset();
						daDataAdapter.Fill(dtbAux);
						if (dtbAux.Rows.Count == 0)
						{
							itemPedidoCalculoParcelas.vlTotalDestePedido = 0;
						}
						else
						{
							itemPedidoCalculoParcelas.vlTotalDestePedido = BD.readToDecimal(dtbAux.Rows[0]["vl_total"]);
						}
						#endregion

						#region [ Calcula o valor total da família de pedidos ]

						#region [ Monta o Select interno ]
						strSql =
							"SELECT" +
								" Coalesce(Sum(qtde*preco_NF),0) AS vl_total" +
							" FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM i ON (p.pedido=i.pedido)" +
							" WHERE" +
								" (p.pedido LIKE '" + Global.retornaNumeroPedidoBase(strPedido) + BD.CARACTER_CURINGA_TODOS + "')" +
								" AND (st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
							" UNION " +
							" SELECT" +
								" -1*Coalesce(Sum(qtde*preco_NF),0) AS vl_total" +
							" FROM t_PEDIDO p INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO id ON (p.pedido=id.pedido)" +
							" WHERE" +
								" (p.pedido LIKE '" + Global.retornaNumeroPedidoBase(strPedido) + BD.CARACTER_CURINGA_TODOS + "')";
						#endregion

						#region [ Monta o Select externo ]
						strSql = "SELECT" +
								" Sum(vl_total) AS vl_total" +
							" FROM" +
								"(" +
									strSql +
								") t";
						#endregion

						cmCommand.CommandText = strSql;
						dtbAux.Reset();
						daDataAdapter.Fill(dtbAux);
						if (dtbAux.Rows.Count == 0)
						{
							itemPedidoCalculoParcelas.vlTotalFamiliaPedidos = 0;
						}
						else
						{
							itemPedidoCalculoParcelas.vlTotalFamiliaPedidos = BD.readToDecimal(dtbAux.Rows[0]["vl_total"]);
						}
						#endregion

						#region [ Calcula a razão entre os valores deste pedido e a família de pedidos ]
						if (itemPedidoCalculoParcelas.vlTotalFamiliaPedidos == 0)
						{
							itemPedidoCalculoParcelas.razaoValorPedidoFilhote = 0;
						}
						else
						{
							itemPedidoCalculoParcelas.razaoValorPedidoFilhote = itemPedidoCalculoParcelas.vlTotalDestePedido / itemPedidoCalculoParcelas.vlTotalFamiliaPedidos;
						}
						#endregion

						#region [ Analisa se há parcela com pagamento por boleto ]
						if (blnPagtoPorBoleto)
						{
							intQtdePedidosPagtoBoleto++;
							if (strListaPedidosPagtoBoleto.Length > 0) strListaPedidosPagtoBoleto += ", ";
							strListaPedidosPagtoBoleto += strPedido;
						}
						else
						{
							if (strListaPedidosPagtoNaoBoleto.Length > 0) strListaPedidosPagtoNaoBoleto += ", ";
							strListaPedidosPagtoNaoBoleto += strPedido;
						}
						#endregion

						vPedidoCalculoParcelas.Add(itemPedidoCalculoParcelas);
					}
				}
			}
			#endregion

			#region [ Houve algum erro? ]
			if (strMsgErro.Length > 0) throw new FinanceiroException(strMsgErro);
			#endregion

			// Quando há 2 pedidos ou mais, a forma de pagamento deve ser idêntica p/ que se possa somar
			// os valores de cada parcela, caso contrário será lançada uma exceção.

			#region [ Não há pedido(s) que defina qualquer parcela de pagamento por boleto! ]
			if (intQtdePedidosPagtoBoleto == 0)
			{
				if (listaNumeroPedido.Count == 1)
					strMsgErro = "No pedido informado, não há nenhuma parcela que especifique o pagamento por boleto!";
				else
					strMsgErro = "Nos pedidos informados, não há nenhuma parcela que especifique o pagamento por boleto!";

				throw new FinanceiroException(strMsgErro);
			}
			#endregion

			#region [ Há pedidos que são por boleto e outros que não ]
			if (intQtdePedidosPagtoBoleto != intQtdeTotalPedidos)
			{
				strMsgErro = "Há pedido(s) que especifica(m) pagamento via boleto bancário e há pedido(s) que especifica(m) outro(s) meio(s) de pagamento:\n" +
							 "Pagamento via boleto bancário: " + strListaPedidosPagtoBoleto + "\n" +
							 "Pagamento via outros meios: " + strListaPedidosPagtoNaoBoleto + "\n" +
							 "\n" +
							 "Não é possível gerar os dados das parcelas dos boletos!!";
				throw new FinanceiroException(strMsgErro);
			}
			#endregion

			#region [ Há mais do que 1 pedido? ]
			if (listaNumeroPedido.Count > 1)
			{
				#region [ Há pedidos de clientes diferentes? ]
				strWhere = "";
				for (int i = 0; i < listaNumeroPedido.Count; i++)
				{
					if (listaNumeroPedido[i].Trim().Length > 0)
					{
						if (strWhere.Length > 0) strWhere += " OR";
						strWhere += " (pedido = '" + Global.normalizaNumeroPedido(listaNumeroPedido[i].Trim()) + "')";
					}
				}
				
				strSql =
					"SELECT DISTINCT" +
						" id_cliente" +
					" FROM t_PEDIDO" +
					" WHERE" +
						strWhere;

				cmCommand.CommandText = strSql;
				dtbAux.Reset();
				daDataAdapter.Fill(dtbAux);
				if (dtbAux.Rows.Count > 1)
				{
					strMsgErro = "Os pedidos são de clientes diferentes!!" +
								"\n\n" +
								"Não é possível gerar os dados das parcelas dos boletos!!";
					throw new FinanceiroException(strMsgErro);
				}
				#endregion

				#region [ Há pedidos que especificam diferentes formas de pagamento? ]
				for (int i = 0; i < (vPedidoCalculoParcelas.Count - 1); i++)
				{
					if (vPedidoCalculoParcelas[i].tipo_parcelamento != vPedidoCalculoParcelas[i + 1].tipo_parcelamento)
					{
						if (strMsgErro.Length > 0) strMsgErro += "\n";
						strMsgErro += "Pedido " + vPedidoCalculoParcelas[i].pedido + "=" + Global.tipoParcelamentoPedidoDescricao(vPedidoCalculoParcelas[i].tipo_parcelamento) +
									  " e pedido " + vPedidoCalculoParcelas[i + 1].pedido + "=" + Global.tipoParcelamentoPedidoDescricao(vPedidoCalculoParcelas[i + 1].tipo_parcelamento);
					}
				}

				if (strMsgErro.Length > 0)
				{
					strMsgErro = "Os pedidos especificam diferentes formas de pagamento!!" +
								"\n" +
								strMsgErro +
								"\n\n" +
								"Não é possível gerar os dados das parcelas dos boletos!!";
					throw new FinanceiroException(strMsgErro);
				}
				#endregion

				#region [ Há pedidos que p/ uma forma de pagamento definem diferentes prazos de pagamento? ]
				for (int i = 0; i < (vPedidoCalculoParcelas.Count - 1); i++)
				{
					#region [ Parcelado com entrada ]
					if (vPedidoCalculoParcelas[i].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
					{
						if (vPedidoCalculoParcelas[i].pce_forma_pagto_entrada != vPedidoCalculoParcelas[i + 1].pce_forma_pagto_entrada)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência na forma de pagamento da entrada: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + Global.formaPagtoPedidoDescricao(vPedidoCalculoParcelas[i].pce_forma_pagto_entrada) + ") e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + Global.formaPagtoPedidoDescricao(vPedidoCalculoParcelas[i + 1].pce_forma_pagto_entrada) + ")";
						}
						if (vPedidoCalculoParcelas[i].pce_forma_pagto_prestacao != vPedidoCalculoParcelas[i + 1].pce_forma_pagto_prestacao)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência na forma de pagamento das prestações: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + Global.formaPagtoPedidoDescricao(vPedidoCalculoParcelas[i].pce_forma_pagto_prestacao) + ") e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + Global.formaPagtoPedidoDescricao(vPedidoCalculoParcelas[i + 1].pce_forma_pagto_prestacao) + ")";
						}
						if (vPedidoCalculoParcelas[i].pce_prestacao_qtde != vPedidoCalculoParcelas[i + 1].pce_prestacao_qtde)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência na quantidade de prestações: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + vPedidoCalculoParcelas[i].pce_prestacao_qtde.ToString() + " " + (vPedidoCalculoParcelas[i].pce_prestacao_qtde > 1 ? "prestações" : "prestação") + ") e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + vPedidoCalculoParcelas[i + 1].pce_prestacao_qtde.ToString() + " " + (vPedidoCalculoParcelas[i + 1].pce_prestacao_qtde > 1 ? "prestações" : "prestação") + ")";
						}
						if (vPedidoCalculoParcelas[i].pce_prestacao_periodo != vPedidoCalculoParcelas[i + 1].pce_prestacao_periodo)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência no período de vencimento das prestações: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + vPedidoCalculoParcelas[i].pce_prestacao_periodo.ToString() + " dias) e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + vPedidoCalculoParcelas[i + 1].pce_prestacao_periodo.ToString() + " dias)";
						}
					}
					#endregion

					#region [ Parcelado sem entrada ]
					else if (vPedidoCalculoParcelas[i].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
					{
						if (vPedidoCalculoParcelas[i].pse_forma_pagto_prim_prest != vPedidoCalculoParcelas[i + 1].pse_forma_pagto_prim_prest)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência na forma de pagamento da 1ª prestação: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + Global.formaPagtoPedidoDescricao(vPedidoCalculoParcelas[i].pse_forma_pagto_prim_prest) + ") e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + Global.formaPagtoPedidoDescricao(vPedidoCalculoParcelas[i + 1].pse_forma_pagto_prim_prest) + ")";
						}
						if (vPedidoCalculoParcelas[i].pse_forma_pagto_demais_prest != vPedidoCalculoParcelas[i + 1].pse_forma_pagto_demais_prest)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência na forma de pagamento das demais prestações: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + Global.formaPagtoPedidoDescricao(vPedidoCalculoParcelas[i].pse_forma_pagto_demais_prest) + ") e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + Global.formaPagtoPedidoDescricao(vPedidoCalculoParcelas[i + 1].pse_forma_pagto_demais_prest) + ")";
						}
						if (vPedidoCalculoParcelas[i].pse_prim_prest_apos != vPedidoCalculoParcelas[i + 1].pse_prim_prest_apos)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência no prazo de pagamento da 1ª prestação: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + vPedidoCalculoParcelas[i].pse_prim_prest_apos.ToString() + ") e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + vPedidoCalculoParcelas[i + 1].pse_prim_prest_apos.ToString() + ")";
						}
						if (vPedidoCalculoParcelas[i].pse_demais_prest_qtde != vPedidoCalculoParcelas[i + 1].pse_demais_prest_qtde)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência na quantidade de prestações: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + vPedidoCalculoParcelas[i].pse_demais_prest_qtde.ToString() + ") e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + vPedidoCalculoParcelas[i + 1].pse_demais_prest_qtde.ToString() + ")";
						}
						if (vPedidoCalculoParcelas[i].pse_demais_prest_periodo != vPedidoCalculoParcelas[i + 1].pse_demais_prest_periodo)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência no período de vencimento das prestações: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + vPedidoCalculoParcelas[i].pse_demais_prest_periodo.ToString() + " dias) e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + vPedidoCalculoParcelas[i + 1].pse_demais_prest_periodo.ToString() + " dias)";
						}
					}
					#endregion

					#region [ Parcela única ]
					else if (vPedidoCalculoParcelas[i].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
					{
						if (vPedidoCalculoParcelas[i].pu_vencto_apos != vPedidoCalculoParcelas[i + 1].pu_vencto_apos)
						{
							if (strMsgErro.Length > 0) strMsgErro += "\n";
							strMsgErro += "Divergência no prazo de vencimento da parcela única: " +
										  vPedidoCalculoParcelas[i].pedido + " (" + vPedidoCalculoParcelas[i].pu_vencto_apos.ToString() + " dia(s)) e " +
										  vPedidoCalculoParcelas[i + 1].pedido + " (" + vPedidoCalculoParcelas[i + 1].pu_vencto_apos.ToString() + " dia(s))";
						}
					}
					#endregion
				}

				if (strMsgErro.Length > 0)
				{
					strMsgErro = "Os pedidos especificam diferentes prazos e/ou condições de pagamento para a mesma forma de pagamento: " + Global.tipoParcelamentoPedidoDescricao(vPedidoCalculoParcelas[vPedidoCalculoParcelas.Count - 1].tipo_parcelamento) + "!!" +
								 "\n\n" +
								 strMsgErro +
								 "\n\n" +
								 "Não é possível gerar os dados das parcelas dos boletos!!";
					throw new FinanceiroException(strMsgErro);
				}
				#endregion

				#region [ Os pedidos são de lojas que especificam diferentes planos de conta? ]
				strWhere = "";
				for (int i = 0; i < listaNumeroPedido.Count; i++)
				{
					if (listaNumeroPedido[i].Trim().Length > 0)
					{
						if (strWhere.Length > 0) strWhere += " OR";
						strWhere += " (pedido = '" + Global.normalizaNumeroPedido(listaNumeroPedido[i].Trim()) + "')";
					}
				}

				strSql =
					"SELECT DISTINCT" +
						" id_plano_contas_empresa," +
						" id_plano_contas_grupo," +
						" id_plano_contas_conta," +
						" natureza" +
					" FROM t_PEDIDO tP" +
						" INNER JOIN t_LOJA tL ON (tP.loja=tL.loja)" +
					" WHERE" +
						strWhere;

				cmCommand.CommandText = strSql;
				dtbAux.Reset();
				daDataAdapter.Fill(dtbAux);
				intQtdePlanoContas = dtbAux.Rows.Count;
				if (intQtdePlanoContas > 1)
				{
					strMsgErro = "Os pedidos são de lojas que especificam diferentes planos de conta!!" +
								"\n\n" +
								"Não é possível gerar os dados das parcelas dos boletos!!";
					throw new FinanceiroException(strMsgErro);
				}
				#endregion
			}
			#endregion

			#region [ Houve algum erro? ]
			if (strMsgErro.Length > 0)
			{
				throw new FinanceiroException(strMsgErro);
			}
			#endregion

			#region [ Obtém o valor total ]
			for (int i = 0; i < vPedidoCalculoParcelas.Count; i++)
			{
				if (vPedidoCalculoParcelas[i].pedido.Trim().Length > 0)
				{
					vlTotalPedido += vPedidoCalculoParcelas[i].vlTotalDestePedido;
					
					#region [ Dados do rateio no caso de pagamento à vista ]
					if (vPedidoCalculoParcelas[i].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
					{
						if (vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio.Length > 0) vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += "|";
						vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += vPedidoCalculoParcelas[i].pedido + "=" + Global.formataMoeda(vPedidoCalculoParcelas[i].vlTotalDestePedido);
					}
					#endregion
				}
			}
			#endregion

			#region [ Consiste valor total c/ a soma dos valores definidos na forma de pagto ]
			for (int i = 0; i < vPedidoCalculoParcelas.Count; i++)
			{
				if (vPedidoCalculoParcelas[i].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
				{
					vlTotalFormaPagto += Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pu_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
				}
				else if (vPedidoCalculoParcelas[i].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
				{
					vlTotalFormaPagto += Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pce_entrada_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
					vlTotalFormaPagto += vPedidoCalculoParcelas[i].pce_prestacao_qtde * Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pce_prestacao_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
				}
				else if (vPedidoCalculoParcelas[i].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
				{
					vlTotalFormaPagto += Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pse_prim_prest_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
					vlTotalFormaPagto += vPedidoCalculoParcelas[i].pse_demais_prest_qtde * Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pse_demais_prest_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
				}
			}

			if (vPedidoCalculoParcelas[0].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
			{
				vlTotalFormaPagto = vlTotalPedido;
			}

			vlDiferencaArredondamento = vlTotalPedido - vlTotalFormaPagto;
			vlDiferencaArredondamentoRestante = vlDiferencaArredondamento;

			if (Math.Abs(vlDiferencaArredondamento) > 1)
			{
				strMsgErro = "A soma dos valores definidos na forma de pagamento (" + Global.formataMoeda(vlTotalFormaPagto) + ") não coincide com o valor total do(s) pedido(s) (" + Global.formataMoeda(vlTotalPedido) + ")!";
				if (Global.Parametro.BoletoAvulso_PermitirDivergenciaValoresFormaPagtoVsPedido == 0)
				{
					strMsgErro += "\n" +
								"Não é possível gerar os dados das parcelas dos boletos!";
					throw new FinanceiroException(strMsgErro);
				}
				else
				{
					msgSolicitacaoConfirmacao = strMsgErro;
				}
			}
			#endregion

			#region [ Calcula os dados das parcelas dos boletos ]
			// LEMBRANDO QUE:
			//      SE O PRAZO DEFINIDO PARA O 1º BOLETO FOR ATÉ 29 DIAS ENTÃO:
			//          VENCIMENTO = DATA EM QUE A NF ESTÁ SENDO EMITIDA + 30 DIAS
			//      SENÃO
			//          VENCIMENTO = DATA EM QUE A NF ESTÁ SENDO EMITIDA + PRAZO DEFINIDO PELO CLIENTE + 7 DIAS

			#region [ À Vista ]
			if (vPedidoCalculoParcelas[0].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
			{
				vParcelaPagto[vParcelaPagto.Count - 1].intNumDestaParcela = 1;
				vParcelaPagto[vParcelaPagto.Count - 1].intNumTotalParcelas = 1;
				vParcelaPagto[vParcelaPagto.Count - 1].id_forma_pagto = vPedidoCalculoParcelas[0].av_forma_pagto;
				vParcelaPagto[vParcelaPagto.Count - 1].vlValor = vlTotalPedido;
				vParcelaPagto[vParcelaPagto.Count - 1].dtVencto = DateTime.Today.AddDays(30);
			}
			#endregion

			#region [ Parcela única ]
			if (vPedidoCalculoParcelas[0].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
			{
				vParcelaPagto[vParcelaPagto.Count - 1].intNumDestaParcela = 1;
				vParcelaPagto[vParcelaPagto.Count - 1].intNumTotalParcelas = 1;
				vParcelaPagto[vParcelaPagto.Count - 1].id_forma_pagto = vPedidoCalculoParcelas[0].pu_forma_pagto;
				vParcelaPagto[vParcelaPagto.Count - 1].dtVencto = calculaDataPrimeiroBoleto(vPedidoCalculoParcelas[0].pu_vencto_apos);
				for (int i = 0; i < vPedidoCalculoParcelas.Count; i++)
				{
					vParcelaPagto[vParcelaPagto.Count - 1].vlValor += Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pu_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
					if (vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio.Length > 0) vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += "|";
					vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += vPedidoCalculoParcelas[i].pedido + "=" + Global.formataMoeda(Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pu_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote));
				}
			}
			#endregion

			#region [ Parcelado com entrada ]
			if (vPedidoCalculoParcelas[0].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
			{
				#region [ Entrada ]
				vParcelaPagto[vParcelaPagto.Count - 1].intNumDestaParcela = 1;
				intQtdeTotalParcelas = 1;
				vParcelaPagto[vParcelaPagto.Count - 1].id_forma_pagto = vPedidoCalculoParcelas[0].pce_forma_pagto_entrada;
				// Entrada é por boleto?
				if (vPedidoCalculoParcelas[0].pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					dtUltimoPagtoCalculado = DateTime.Today.AddDays(30);
				else
					dtUltimoPagtoCalculado = DateTime.Today;

				vParcelaPagto[vParcelaPagto.Count - 1].dtVencto = dtUltimoPagtoCalculado;
				for (int i = 0; i < vPedidoCalculoParcelas.Count; i++)
				{
					vParcelaPagto[vParcelaPagto.Count - 1].vlValor += Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pce_entrada_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
					vlRateio = Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pce_entrada_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
					if (vPedidoCalculoParcelas[0].pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					{
						if (vlDiferencaArredondamentoRestante != 0)
						{
							vParcelaPagto[vParcelaPagto.Count - 1].vlValor += vlDiferencaArredondamentoRestante;
							vlRateio += vlDiferencaArredondamentoRestante;
							vlDiferencaArredondamentoRestante = 0;
						}
					}
					if (vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio.Length > 0) vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += "|";
					vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += vPedidoCalculoParcelas[i].pedido + "=" + Global.formataMoeda(vlRateio);
				}
				#endregion

				#region [ Prestações ]
				for (int i = 1; i <= vPedidoCalculoParcelas[0].pce_prestacao_qtde; i++)
				{
					intQtdeTotalParcelas++;
					if (vParcelaPagto[vParcelaPagto.Count - 1].intNumDestaParcela != 0)
					{
						vParcelaPagto.Add(new TipoLinhaDadosParcelaPagto());
					}

					vParcelaPagto[vParcelaPagto.Count - 1].intNumDestaParcela = intQtdeTotalParcelas;
					vParcelaPagto[vParcelaPagto.Count - 1].id_forma_pagto = vPedidoCalculoParcelas[0].pce_forma_pagto_prestacao;

					#region [ Prestações são por boleto? ]
					if (vPedidoCalculoParcelas[0].pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					{
						#region [ A entrada não foi paga por boleto! ]
						if (intQtdeTotalParcelas == 1)
						{
							#region [ Esta prestação será o 1º boleto da série ]
							if (vPedidoCalculoParcelas[0].pce_prestacao_periodo == 30)
							{
								dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddMonths(1);
							}
							else if (vPedidoCalculoParcelas[0].pce_prestacao_periodo <= 29)
							{
								dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddDays(30);
							}
							else
							{
								dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddDays(vPedidoCalculoParcelas[0].pce_prestacao_periodo);
							}
							#endregion
						}
						#endregion

						#region [ A entrada foi paga por boleto! ]
						else
						{
							#region [ Calcula a data dos demais boletos ]
							if (vPedidoCalculoParcelas[0].pce_prestacao_periodo == 30)
							{
								dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddMonths(1);
							}
							else
							{
								dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddDays(vPedidoCalculoParcelas[0].pce_prestacao_periodo);
							}
							#endregion
						}
						#endregion
					}
					#endregion

					#region [ Prestações não são por boleto! ]
					else
					{
						#region [ Cálculo p/ prestações que não são por boleto ]
						if (vPedidoCalculoParcelas[0].pce_prestacao_periodo == 30)
						{
							dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddMonths(1);
						}
						else
						{
							dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddDays(vPedidoCalculoParcelas[0].pce_prestacao_periodo);
						}
						#endregion
					}
					#endregion

					vParcelaPagto[vParcelaPagto.Count - 1].dtVencto = dtUltimoPagtoCalculado;

					for (int j = 0; j < vPedidoCalculoParcelas.Count; j++)
					{
						vParcelaPagto[vParcelaPagto.Count - 1].vlValor += Global.arredondaParaMonetario(vPedidoCalculoParcelas[j].pce_prestacao_valor * vPedidoCalculoParcelas[j].razaoValorPedidoFilhote);
						vlRateio = Global.arredondaParaMonetario(vPedidoCalculoParcelas[j].pce_prestacao_valor * vPedidoCalculoParcelas[j].razaoValorPedidoFilhote);
						if (vPedidoCalculoParcelas[0].pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
						{
							if (vlDiferencaArredondamentoRestante != 0)
							{
								vParcelaPagto[vParcelaPagto.Count - 1].vlValor += vlDiferencaArredondamentoRestante;
								vlRateio += vlDiferencaArredondamentoRestante;
								vlDiferencaArredondamentoRestante = 0;
							}
						}

						if (vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio.Length > 0) vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += "|";
						vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += vPedidoCalculoParcelas[j].pedido + "=" + Global.formataMoeda(vlRateio);
					}
				}
				#endregion

				for (int i = 0; i < vParcelaPagto.Count; i++)
				{
					vParcelaPagto[i].intNumTotalParcelas = intQtdeTotalParcelas;
				}
			}
			#endregion

			#region [ Parcelado sem entrada ]
			if (vPedidoCalculoParcelas[0].tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
			{
				#region [ 1ª prestação ]
				vParcelaPagto[vParcelaPagto.Count - 1].intNumDestaParcela = 1;
				intQtdeTotalParcelas = 1;
				vParcelaPagto[vParcelaPagto.Count - 1].id_forma_pagto = vPedidoCalculoParcelas[0].pse_forma_pagto_prim_prest;
				// 1ª prestação é por boleto?
				if (vPedidoCalculoParcelas[0].pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					dtUltimoPagtoCalculado = calculaDataPrimeiroBoleto(vPedidoCalculoParcelas[0].pse_prim_prest_apos);
				}
				else
				{
					dtUltimoPagtoCalculado = DateTime.Today.AddDays(vPedidoCalculoParcelas[0].pse_prim_prest_apos);
				}

				vParcelaPagto[vParcelaPagto.Count - 1].dtVencto = dtUltimoPagtoCalculado;
				for (int i = 0; i < vPedidoCalculoParcelas.Count; i++)
				{
					vParcelaPagto[vParcelaPagto.Count - 1].vlValor += Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pse_prim_prest_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
					vlRateio = Global.arredondaParaMonetario(vPedidoCalculoParcelas[i].pse_prim_prest_valor * vPedidoCalculoParcelas[i].razaoValorPedidoFilhote);
					if (vPedidoCalculoParcelas[0].pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					{
						if (vlDiferencaArredondamentoRestante != 0)
						{
							vParcelaPagto[vParcelaPagto.Count - 1].vlValor += vlDiferencaArredondamentoRestante;
							vlRateio += vlDiferencaArredondamentoRestante;
							vlDiferencaArredondamentoRestante = 0;
						}
					}

					if (vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio.Length > 0) vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += "|";
					vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += vPedidoCalculoParcelas[i].pedido + "=" + Global.formataMoeda(vlRateio);
				}
				#endregion

				#region [ Demais prestações ]
				for (int i = 1; i <= vPedidoCalculoParcelas[0].pse_demais_prest_qtde; i++)
				{
					intQtdeTotalParcelas++;
					if (vParcelaPagto[vParcelaPagto.Count - 1].intNumDestaParcela != 0)
					{
						vParcelaPagto.Add(new TipoLinhaDadosParcelaPagto());
					}

					vParcelaPagto[vParcelaPagto.Count - 1].intNumDestaParcela = intQtdeTotalParcelas;
					vParcelaPagto[vParcelaPagto.Count - 1].id_forma_pagto = vPedidoCalculoParcelas[0].pse_forma_pagto_demais_prest;

					#region [ Demais prestações são por boleto? ]
					if (vPedidoCalculoParcelas[0].pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					{
						#region [ A 1ª prestação não foi paga por boleto! ]
						if (intQtdeTotalParcelas == 1)
						{
							#region [ Esta prestação será o 1º boleto da série ]
							if ((vPedidoCalculoParcelas[0].pse_prim_prest_apos +
								 vPedidoCalculoParcelas[0].pse_demais_prest_periodo) >= 30)
							{
								if (vPedidoCalculoParcelas[0].pse_demais_prest_periodo == 30)
								{
									dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddMonths(1);
								}
								else
								{
									dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddDays(vPedidoCalculoParcelas[0].pse_demais_prest_periodo);
								}
							}
							else
							{
								dtUltimoPagtoCalculado = DateTime.Today.AddDays(30);
							}
							#endregion
						}
						else
						{
							#region [ Calcula a data dos demais boletos ]
							if (vPedidoCalculoParcelas[0].pse_demais_prest_periodo == 30)
							{
								dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddMonths(1);
							}
							else
							{
								dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddDays(vPedidoCalculoParcelas[0].pse_demais_prest_periodo);
							}
							#endregion
						}
						#endregion
					}
					#endregion

					#region [ Cálculo p/ prestações que não são por boleto ]
					else
					{
						if (vPedidoCalculoParcelas[0].pse_demais_prest_periodo == 30)
						{
							dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddMonths(1);
						}
						else
						{
							dtUltimoPagtoCalculado = dtUltimoPagtoCalculado.AddDays(vPedidoCalculoParcelas[0].pse_demais_prest_periodo);
						}
					}
					#endregion

					vParcelaPagto[vParcelaPagto.Count - 1].dtVencto = dtUltimoPagtoCalculado;
					for (int j = 0; j < vPedidoCalculoParcelas.Count; j++)
					{
						vParcelaPagto[vParcelaPagto.Count - 1].vlValor += Global.arredondaParaMonetario(vPedidoCalculoParcelas[j].pse_demais_prest_valor * vPedidoCalculoParcelas[j].razaoValorPedidoFilhote);
						vlRateio = Global.arredondaParaMonetario(vPedidoCalculoParcelas[j].pse_demais_prest_valor * vPedidoCalculoParcelas[j].razaoValorPedidoFilhote);
						if (vPedidoCalculoParcelas[0].pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
						{
							if (vlDiferencaArredondamentoRestante != 0)
							{
								vParcelaPagto[vParcelaPagto.Count - 1].vlValor += vlDiferencaArredondamentoRestante;
								vlRateio += vlDiferencaArredondamentoRestante;
								vlDiferencaArredondamentoRestante = 0;
							}
						}
						if (vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio.Length > 0) vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += "|";
						vParcelaPagto[vParcelaPagto.Count - 1].strDadosRateio += vPedidoCalculoParcelas[j].pedido + "=" + Global.formataMoeda(vlRateio);
					}
				}
				#endregion

				for (int i = 0; i < vParcelaPagto.Count; i++)
				{
					vParcelaPagto[i].intNumTotalParcelas = intQtdeTotalParcelas;
				}
			}
			#endregion

			#endregion

			return vParcelaPagto;
		}
		#endregion

		#region [ geraDadosBoletoAvulsoComPedido ]
		/// <summary>
		/// A partir da lista com as parcelas de pagamento, independentemente se a parcela é por boleto ou não,
		/// monta os dados para editar/cadastrar o boleto avulso, nos mesmos moldes que é feito através dos
		/// dados gerados na emissão da NF.
		/// </summary>
		/// <param name="listaNumeroPedido">
		/// Lista com os números de pedido que serão cobrados juntos nesta série de boletos
		/// </param>
		/// <param name="vParcelaPagto">
		/// Lista com as parcelas de pagamento, independentemente se a parcela é por boleto ou não
		/// </param>
		/// <returns>
		/// Retorna um objeto BoletoAvulsoComPedido com os dados para editar/cadastrar o boleto.
		/// </returns>
		private BoletoAvulsoComPedido geraDadosBoletoAvulsoComPedido(List<String> listaNumeroPedido, List<TipoLinhaDadosParcelaPagto> vParcelaPagto)
		{
			#region [ Declarações ]
			int intQtdeParcelas = 0;
			int intQtdeParcelasBoleto = 0;
			int idBoleto;
			int idNfParcelaPagto;
			int numeroNF;
			bool blnSucesso;
			String strAux;
			String strSql;
			String strWhere;
			String strMsgErro = "";
			String strDescricaoLog = "";
			String strMsgAutorizacao;
			String strIdCliente;
			String[] vRateio;
			String[] vRateioDetalhe;
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
			BoletoAvulsoComPedido boletoAvulsoComPedido = new BoletoAvulsoComPedido();
			BoletoAvulsoComPedidoItem boletoAvulsoComPedidoItem;
			BoletoAvulsoComPedidoItemRateio boletoAvulsoComPedidoItemRateio;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataTable dtbAux = new DataTable();
			#endregion

			#region [ Verifica se a conexão c/ o BD está ok ]
			if (!BD.isConexaoOk())
			{
				if (!FMain.reiniciaBancoDados())
				{
					throw new FinanceiroException("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
				}
			}
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Consistência ]
			for (int i = 0; i < vParcelaPagto.Count; i++)
			{
				if (vParcelaPagto[i].intNumDestaParcela > 0) intQtdeParcelas++;
				if ((vParcelaPagto[i].id_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO) 
					|| (vParcelaPagto[i].id_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO_AV)) intQtdeParcelasBoleto++;
			}

			if (intQtdeParcelas == 0) throw new FinanceiroException("Não há nenhuma parcela de pagamento definida!!");
			if (intQtdeParcelasBoleto == 0) throw new FinanceiroException("Não há nenhuma parcela de pagamento por boleto!!");
			#endregion

			#region [ Obtém identificação do cliente ]
			strWhere = "";
			for (int i = 0; i < listaNumeroPedido.Count; i++)
			{
				if (listaNumeroPedido[i].Trim().Length > 0)
				{
					if (strWhere.Length > 0) strWhere += " OR";
					strWhere += " (pedido = '" + Global.normalizaNumeroPedido(listaNumeroPedido[i].Trim()) + "')";
				}
			}

			strSql =
				"SELECT DISTINCT" +
					" id_cliente" +
				" FROM t_PEDIDO" +
				" WHERE" +
					strWhere;

			cmCommand.CommandText = strSql;
			dtbAux.Reset();
			daDataAdapter.Fill(dtbAux);
			if (dtbAux.Rows.Count == 0)
			{
				strMsgErro = "Falha ao tentar obter a identificação do cliente!!" +
							 "\n\n" +
							 "Não é possível gerar os dados das parcelas dos boletos!!";
				throw new FinanceiroException(strMsgErro);
			}
			else if (dtbAux.Rows.Count > 1)
			{
				strMsgErro = "Os pedidos são de clientes diferentes!!" +
							 "\n\n" +
							 "Não é possível gerar os dados das parcelas dos boletos!!";
				throw new FinanceiroException(strMsgErro);
			}
			else
			{
				strIdCliente = BD.readToString(dtbAux.Rows[0]["id_cliente"]);
			}
			#endregion

			#region [ Há registro em t_FIN_NF_PARCELA_PAGTO (dados já foram gerados pela emissão de NF)? ]
			strWhere = "";
			for (int i = 0; i < listaNumeroPedido.Count; i++)
			{
				if (listaNumeroPedido[i].Trim().Length > 0)
				{
					if (strWhere.Length > 0) strWhere += " OR";
					strWhere += " (pedido = '" + Global.normalizaNumeroPedido(listaNumeroPedido[i].Trim()) + "')";
				}
			}

			strSql =
				"SELECT DISTINCT" +
					" tpp.id" +
				" FROM t_FIN_NF_PARCELA_PAGTO tpp" +
					" INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM tppi" +
						" ON (tpp.id=tppi.id_nf_parcela_pagto)" +
					" INNER JOIN t_FIN_NF_PARCELA_PAGTO_ITEM_RATEIO tppir" +
						" ON (tppi.id=tppir.id_nf_parcela_pagto_item)" +
				" WHERE" +
					" (tpp.status = " + Global.Cte.FIN.ST_T_FIN_NF_PARCELA_PAGTO.INICIAL.ToString() + ")" +
					" AND (" + strWhere + ")";

			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);

			for (int intCounter = 0; intCounter < dtbResultado.Rows.Count; intCounter++)
			{
				idNfParcelaPagto = BD.readToInt(dtbResultado.Rows[intCounter]["id"]);

				#region [ Obtém o nº da NF ]
				strSql =
					"SELECT" +
						" numero_NF" +
					" FROM t_FIN_NF_PARCELA_PAGTO" +
					" WHERE" +
						" (id = " + idNfParcelaPagto.ToString() + ")";
				cmCommand.CommandText = strSql;
				dtbAux.Reset();
				daDataAdapter.Fill(dtbAux);
				if (dtbAux.Rows.Count > 0)
					numeroNF = BD.readToInt(dtbAux.Rows[0]["numero_NF"]);
				else
					numeroNF = 0;
				#endregion

				#region [ Obtém os números dos pedidos e solicita confirmação p/ cancelar registro ]
				strMsgAutorizacao = "";
				foreach (var item in BoletoPreCadastradoDAO.obtemListaNumeroPedidoRateio(idNfParcelaPagto))
				{
					if (strMsgAutorizacao.Length > 0) strMsgAutorizacao += ", ";
					strMsgAutorizacao += item;
				}
				strMsgAutorizacao = "A emissão da NF: " + numeroNF.ToString() + " gerou dados para cadastramento de boletos para o(s) pedido(s): " + strMsgAutorizacao + "!!" +
							 "\n\n" +
							 "Cancela esses dados gerados durante a emissão da NF?";

				fAutorizacao = new FAutorizacao(strMsgAutorizacao);
				while (true)
				{
					drAutorizacao = fAutorizacao.ShowDialog();
					if (drAutorizacao != DialogResult.OK)
					{
						strMsgErro = "Os dados gerados devido à emissão da NF " + numeroNF.ToString() + " NÃO foram cancelados!!" +
									 "\n\n" +
									 "Atenção para não cadastrar boletos em duplicidade!!";
						aviso(strMsgErro);
						break;
					}
					else
					{
						if (fAutorizacao.senha.ToUpper() == Global.Usuario.senhaDescriptografada.ToUpper())
						{
							#region [ Cancela registro! ]
							if (!BoletoPreCadastradoDAO.anula(Global.Usuario.usuario, idNfParcelaPagto, ref strDescricaoLog, ref strMsgErro))
							{
								throw new FinanceiroException("Falha ao tentar cancelar os dados gerados durante a emissão da NF " + numeroNF.ToString() + "!!\n" + strMsgErro);
							}
							break;
							#endregion
						}
						else
						{
							avisoErro("Senha inválida!!");
						}
					}
				}
				#endregion
			}
			#endregion

			#region [ Há boleto cadastrado ainda não enviado no arquivo de remessa? ]
			for (int intCounter = 0; intCounter < listaNumeroPedido.Count; intCounter++)
			{
				strSql =
					"SELECT DISTINCT" +
						" tB.id" +
					" FROM t_FIN_BOLETO tB" +
						" INNER JOIN t_FIN_BOLETO_ITEM tBI ON (tB.id=tBI.id_boleto)" +
						" INNER JOIN t_FIN_BOLETO_ITEM_RATEIO tBIR ON (tBI.id=tBIR.id_boleto_item)" +
					" WHERE" +
						" (tB.status = " + Global.Cte.FIN.CodBoletoStatus.INICIAL.ToString() + ")" +
						" AND (tBI.status = " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ")" +
						" AND (pedido = '" + Global.normalizaNumeroPedido(listaNumeroPedido[intCounter].Trim()) + "')" +
					" ORDER BY" +
						" tB.id";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					idBoleto = BD.readToInt(dtbResultado.Rows[i]["id"]);
					strMsgAutorizacao = "";
					foreach (var item in BoletoDAO.obtemListaNumeroPedidoRateio(idBoleto))
					{
						if (strMsgAutorizacao.Length > 0) strMsgAutorizacao += ", ";
						strMsgAutorizacao += item;
					}
					strMsgAutorizacao = "Já existe boleto cadastrado (ainda não enviado no arquivo de remessa) para o(s) pedido(s): " + strMsgAutorizacao + "!!" +
										"\nCancela automaticamente esse boleto?";

					#region [ Detalhes das parcelas ]
					strSql =
						"SELECT DISTINCT" +
							" num_parcela," +
							" dt_vencto," +
							" valor" +
						" FROM t_FIN_BOLETO_ITEM" +
						" WHERE" +
							" (id_boleto = " + idBoleto.ToString() + ")" +
							" AND (status = " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ")" +
						" ORDER BY" +
							" dt_vencto";
					cmCommand.CommandText = strSql;
					dtbAux.Reset();
					daDataAdapter.Fill(dtbAux);
					strAux = "";
					for (int j = 0; j < dtbAux.Rows.Count; j++)
					{
						if (strAux.Length > 0) strAux += "\n";
						strAux += "Parc. " + BD.readToByte(dtbAux.Rows[j]["num_parcela"]).ToString() + ": " + Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(dtbAux.Rows[j]["dt_vencto"])) + " = " + Global.formataMoeda(BD.readToDecimal(dtbAux.Rows[j]["valor"]));
					}

					if (strAux.Length > 0) strMsgAutorizacao += "\n\nDetalhes dos boletos:\n" + strAux;
					#endregion

					fAutorizacao = new FAutorizacao(strMsgAutorizacao);
					while (true)
					{
						drAutorizacao = fAutorizacao.ShowDialog();
						if (drAutorizacao != DialogResult.OK)
						{
							strMsgErro = "Os boletos NÃO foram cancelados!!" +
										 "\n\n" +
										 "Atenção para não cadastrar boletos em duplicidade!!";
							aviso(strMsgErro);
							break;
						}
						else
						{
							if (fAutorizacao.senha.ToUpper() == Global.Usuario.senhaDescriptografada.ToUpper())
							{
								#region [ Cancela os boletos! ]
								blnSucesso = false;
								try
								{
									BD.iniciaTransacao();

									if (!BoletoDAO.marcaBoletoCanceladoManual(Global.Usuario.usuario, idBoleto, ref strMsgErro))
									{
										throw new FinanceiroException("Falha ao tentar cancelar os boletos anteriores!!\n" + strMsgErro);
									}

									if (!BoletoDAO.marcaBoletoItemCanceladoManualByIdBoleto(Global.Usuario.usuario, idBoleto, ref strMsgErro))
									{
										throw new FinanceiroException("Falha ao tentar cancelar as parcelas dos boletos anteriores!!\n" + strMsgErro);
									}

									blnSucesso = true;
								}
								finally
								{
									if (blnSucesso)
										BD.commitTransacao();
									else
										BD.rollbackTransacao();
								}

								break;
								#endregion
							}
							else
							{
								avisoErro("Senha inválida!!");
							}
						}
					}
				}
			}
			#endregion

			#region [ Há boleto cadastrado e já enviado no arquivo de remessa? ]
			for (int intCounter = 0; intCounter < listaNumeroPedido.Count; intCounter++)
			{
				strSql =
					"SELECT DISTINCT" +
						" tB.id" +
					" FROM t_FIN_BOLETO tB" +
						" INNER JOIN t_FIN_BOLETO_ITEM tBI ON (tB.id=tBI.id_boleto)" +
						" INNER JOIN t_FIN_BOLETO_ITEM_RATEIO tBIR ON (tBI.id=tBIR.id_boleto_item)" +
					" WHERE" +
						" (tBI.status = " + Global.Cte.FIN.CodBoletoItemStatus.ENVIADO_REMESSA_BANCO.ToString() + ")" +
						" AND (pedido = '" + Global.normalizaNumeroPedido(listaNumeroPedido[intCounter].Trim()) + "')" +
					" ORDER BY" +
						" tB.id";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					idBoleto = BD.readToInt(dtbResultado.Rows[i]["id"]);
					strMsgAutorizacao = "";
					foreach (var item in BoletoDAO.obtemListaNumeroPedidoRateio(idBoleto))
					{
						if (strMsgAutorizacao.Length > 0) strMsgAutorizacao += ", ";
						strMsgAutorizacao += item;
					}
					strMsgAutorizacao = "Já existe boleto cadastrado (já enviado no arquivo de remessa, mas cuja entrada ainda não foi confirmada) para o(s) pedido(s): " + strMsgAutorizacao + "!!" +
										"\nAtenção para não cadastrar boletos em duplicidade!!";
					aviso(strMsgAutorizacao);
				}
			}
			#endregion

			#region [ Há boleto cadastrado e com entrada confirmada? ]
			for (int intCounter = 0; intCounter < listaNumeroPedido.Count; intCounter++)
			{
				strSql =
					"SELECT DISTINCT" +
						" tB.id" +
					" FROM t_FIN_BOLETO tB" +
						" INNER JOIN t_FIN_BOLETO_ITEM tBI ON (tB.id=tBI.id_boleto)" +
						" INNER JOIN t_FIN_BOLETO_ITEM_RATEIO tBIR ON (tBI.id=tBIR.id_boleto_item)" +
					" WHERE" +
						" (tBI.status = " + Global.Cte.FIN.CodBoletoItemStatus.ENTRADA_CONFIRMADA.ToString() + ")" +
						" AND (pedido = '" + Global.normalizaNumeroPedido(listaNumeroPedido[intCounter].Trim()) + "')" +
					" ORDER BY" +
						" tB.id";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					idBoleto = BD.readToInt(dtbResultado.Rows[i]["id"]);
					strMsgAutorizacao = "";
					foreach (var item in BoletoDAO.obtemListaNumeroPedidoRateio(idBoleto))
					{
						if (strMsgAutorizacao.Length > 0) strMsgAutorizacao += ", ";
						strMsgAutorizacao += item;
					}
					strMsgAutorizacao = "Já existe boleto cadastrado (entrada confirmada) para o(s) pedido(s): " + strMsgAutorizacao + "!!" +
										"\nAtenção para não cadastrar boletos em duplicidade!!";
					aviso(strMsgAutorizacao);
				}
			}
			#endregion

			boletoAvulsoComPedido.id_cliente = strIdCliente;
			boletoAvulsoComPedido.qtde_parcelas = (byte)intQtdeParcelas;
			boletoAvulsoComPedido.qtde_parcelas_boleto = (byte)intQtdeParcelasBoleto;
			boletoAvulsoComPedido.listaItem = new List<BoletoAvulsoComPedidoItem>();

			for (int intCounter = 0; intCounter < vParcelaPagto.Count; intCounter++)
			{
				boletoAvulsoComPedidoItem = new BoletoAvulsoComPedidoItem();
				boletoAvulsoComPedidoItem.dt_vencto = vParcelaPagto[intCounter].dtVencto;
				boletoAvulsoComPedidoItem.valor = vParcelaPagto[intCounter].vlValor;
				boletoAvulsoComPedidoItem.forma_pagto = vParcelaPagto[intCounter].id_forma_pagto;
				boletoAvulsoComPedidoItem.num_parcela = (byte)vParcelaPagto[intCounter].intNumDestaParcela;

				boletoAvulsoComPedidoItem.listaRateio = new List<BoletoAvulsoComPedidoItemRateio>();

				vRateio = vParcelaPagto[intCounter].strDadosRateio.Split('|');
				for (int i = 0; i < vRateio.Length; i++)
				{
					boletoAvulsoComPedidoItemRateio = new BoletoAvulsoComPedidoItemRateio();
					vRateioDetalhe = vRateio[i].ToString().Split('=');
					boletoAvulsoComPedidoItemRateio.pedido = Global.normalizaNumeroPedido(vRateioDetalhe[0]);
					boletoAvulsoComPedidoItemRateio.valor = Global.converteNumeroDecimal(vRateioDetalhe[1].ToString());
					boletoAvulsoComPedidoItem.listaRateio.Add(boletoAvulsoComPedidoItemRateio);
				}
				boletoAvulsoComPedido.listaItem.Add(boletoAvulsoComPedidoItem);
			}
			
			return boletoAvulsoComPedido;
		}
		#endregion

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

		#region [ ajustaPosicaoLblTotalGridParcelasBase ]
		private void ajustaPosicaoLblTotalGridParcelasBase()
		{
			lblTotalGridParcelasBase.Left = grdParcelasBase.Left + grdParcelasBase.Width - lblTotalGridParcelasBase.Width - 3;
			if (Global.isVScrollBarVisible(grdParcelasBase)) lblTotalGridParcelasBase.Left -= Global.getVScrollBarWidth(grdParcelasBase);
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
			lblTotalGridParcelasBase.Text = "";
			lblTotalGridParcelas.Text = "";
			lbPedido.Items.Clear();
			txtNumeroDocumento.Text = "";
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
			DsDataSource.DtbBoletoCedenteComboRow rowBoletoCedente;
			#endregion

			for (int i = 0; i < _listaPedidos.Count; i++)
			{
				intIdBoletoCedenteAux = BoletoDAO.obtemBoletoCedenteDefinidoParaLoja((int)Global.converteInteiro(_listaPedidos[i].loja));
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

			}

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
			int intNumeroDocumento;
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

			if (rbNumeroNF.Checked)
			{
				boletoEditado.numero_NF = (int)Global.converteInteiro(txtNumeroDocumento.Text);
				boletoEditado.num_documento_boleto_avulso = 0;
			}
			else
			{
				boletoEditado.num_documento_boleto_avulso = (int)Global.converteInteiro(txtNumeroDocumento.Text);
				boletoEditado.numero_NF = 0;
			}

			boletoEditado.id_cliente = clienteSelecionado.id;
			boletoEditado.tipo_vinculo = Global.Cte.FIN.CodBoletoTipoVinculo.BOLETO_AVULSO_COM_PEDIDO;
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
			intNumeroDocumento = (int)Global.converteInteiro(txtNumeroDocumento.Text);
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

				if (rbNumeroNF.Checked)
				{
					boletoItem.numero_documento = intNumeroDocumento.ToString() +
												  "/" +
												  boletoItem.num_parcela.ToString().PadLeft(2, '0');
				}
				else
				{
					boletoItem.numero_documento = Global.Cte.FIN.PREFIXO_NUMERO_DOCUMENTO_BOLETO_AVULSO +
												  intNumeroDocumento.ToString() +
												  "/" +
												  boletoItem.num_parcela.ToString().PadLeft(2, '0');
				}

				#region [ Instrução de protesto ]
				// Devido ao custo dos cartório, apenas algumas parcelas serão geradas c/ instrução de protesto
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
		private bool isBoletoEditado(BoletoAvulsoComPedido boletoOriginal, Boleto boletoEditado)
		{
			#region [ Declarações ]
			int indiceBoletoEditado = 0;
			String strEndereco;
			#endregion

			if (boletoCliente != null)
			{
				strEndereco = boletoCliente.endereco_logradouro;
				if (boletoCliente.endereco_numero.Length > 0) strEndereco += ", " + boletoCliente.endereco_numero;
				if (boletoCliente.endereco_complemento.Length > 0) strEndereco += " " + boletoCliente.endereco_complemento;

				if (!boletoCliente.nome.ToUpper().Equals(boletoEditado.nome_sacado)) return true;
				if (!strEndereco.ToUpper().Equals(boletoEditado.endereco_sacado)) return true;
				if (!boletoCliente.cnpj_cpf.Equals(boletoEditado.num_inscricao_sacado)) return true;
				if (!boletoCliente.endereco_cep.Equals(boletoEditado.cep_sacado)) return true;
				if (!boletoCliente.email.Equals(boletoEditado.email_sacado)) return true;
			}

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
				if ((boletoOriginal.listaItem[i].forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO) 
					|| (boletoOriginal.listaItem[i].forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO_AV))
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

			if (boletoEditado.numero_NF > 0) return true;
			if (boletoEditado.num_documento_boleto_avulso > 0) return true;

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
			String strPedido;
			List<String> listaPedidos = new List<String>();
			bool blnJaExiste;
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
				for (int i = 0; i < boletoAvulsoComPedidoSelecionado.listaItem.Count; i++)
				{
					for (int j = 0; j < boletoAvulsoComPedidoSelecionado.listaItem[i].listaRateio.Count; j++)
					{
						strPedido = boletoAvulsoComPedidoSelecionado.listaItem[i].listaRateio[j].pedido.ToString().Trim();
						if (strPedido.Length > 0)
						{
							blnJaExiste = false;
							for (int k = 0; k < listaPedidos.Count; k++)
							{
								if (listaPedidos[k].Equals(strPedido))
								{
									blnJaExiste = true;
									break;
								}
							}
							if (!blnJaExiste) listaPedidos.Add(strPedido);
						}
					}
				}
				boletoPlanoContasDestino = BoletoAvulsoComPedidoDAO.obtemBoletoPlanoContasDestino(listaPedidos);
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

			#region [ Número do documento ]
			if ((!rbNumeroBoletoAvulso.Checked) && (!rbNumeroNF.Checked))
			{
				strMsgErro = "Selecione uma das opções para gerar o \"Nº Documento\":" +
							 "\n        " + rbNumeroBoletoAvulso.Text +
							 "\n        " + rbNumeroNF.Text;
				avisoErro(strMsgErro);
				return false;
			}

			if (Global.converteInteiro(txtNumeroDocumento.Text) <= 0)
			{
				avisoErro("Número do documento informado é inválido!!");
				txtNumeroDocumento.Focus();
				return false;
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
			for (int i = 0; i < _listaNumeroPedidoSelecionado.Count; i++)
			{
				strSql = "SELECT DISTINCT" +
							" tFBI.status" +
						" FROM t_FIN_BOLETO_ITEM tFBI" +
							" INNER JOIN t_FIN_BOLETO_ITEM_RATEIO tFBIR" +
								" ON (tFBI.id=tFBIR.id_boleto_item)" +
						" WHERE" +
							" (tFBIR.pedido = '" + Global.normalizaNumeroPedido(_listaNumeroPedidoSelecionado[i]) + "')" +
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
					strMsgConfirmacao += intNumeroAviso.ToString() + ") O pedido " + Global.normalizaNumeroPedido(_listaNumeroPedidoSelecionado[i]) + " já possui boletos cadastrados!!";
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
			List<int> listaNFeEmitente;
			StringBuilder sbPedidoInconsistente = new StringBuilder("");
			bool blnAchouNFeEmitente;
			bool blnHaPedidoCedenteInconsistente = false;
			String strMsgConfirmacao;
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
			#endregion

			if (boletoEditado.id_boleto_cedente == 0)
			{
				avisoErro("Não foi selecionado um cedente válido!!");
				return false;
			}

			listaNFeEmitente = BoletoCedenteDAO.getNFeEmitentesBoletoCedente(boletoEditado.id_boleto_cedente);

			for (int i = 0; i < _listaPedidos.Count; i++)
			{
				blnAchouNFeEmitente = false;
				for (int j = 0; j < listaNFeEmitente.Count; j++)
				{
					if (_listaPedidos[i].id_nfe_emitente == listaNFeEmitente[j])
					{
						blnAchouNFeEmitente = true;
						break;
					}
				}

				if (!blnAchouNFeEmitente)
				{
					blnHaPedidoCedenteInconsistente = true;
					if (sbPedidoInconsistente.Length > 0) sbPedidoInconsistente.Append(", ");
					sbPedidoInconsistente.Append(_listaPedidos[i].pedido);
				}
			}

			if (blnHaPedidoCedenteInconsistente)
			{
				strMsgConfirmacao = "O cedente selecionado diverge da opção definida no sistema para o(s) pedido(s): " + sbPedidoInconsistente.ToString()+"!!\n\nConfirma o cadastramento assim mesmo?";
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

		#region [ trataBotaoCadastrar ]
		void trataBotaoCadastrar()
		{
			#region [ Declarações ]
			String strMsgErro = "";
			String strMsgErroLog = "";
			String strDescricaoLog = "";
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

			#region [ Consistência dos campos ]
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
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_AVULSO_CADASTRA;
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

			strEndereco = boletoCliente.endereco_logradouro;
			if (boletoCliente.endereco_numero.Length > 0) strEndereco += ", " + boletoCliente.endereco_numero;
			if (boletoCliente.endereco_complemento.Length > 0) strEndereco += " " + boletoCliente.endereco_complemento;

			txtEndereco.Text = strEndereco.ToUpper();
			txtCep.Text = Global.formataCep(boletoCliente.endereco_cep);
			txtBairro.Text = boletoCliente.endereco_bairro.ToUpper();
			txtCidade.Text = boletoCliente.endereco_cidade.ToUpper();
			txtUF.Text = boletoCliente.endereco_uf.ToUpper();
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

		#region [ Form: FBoletoAvulsoComPedidoCadDetalhe ]

		#region [ FBoletoAvulsoComPedidoCadDetalhe_Load ]
		private void FBoletoAvulsoComPedidoCadDetalhe_Load(object sender, EventArgs e)
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

		#region [ FBoletoAvulsoComPedidoCadDetalhe_Shown ]
		private void FBoletoAvulsoComPedidoCadDetalhe_Shown(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String strPedido;
			String strDadosRateio;
			String strDadosRateioParcela;
			String strEndereco;
			String strMsg;
			String msgSolicitacaoConfirmacao;
			bool blnAchou;
			Pedido pedido;
			int intIndiceLinhaGrid;
			int intQtdeParcelasBoleto = 0;
			decimal vlTotalParcelasBase = 0;
			decimal vlTotalParcelasBoleto = 0;
			List<TipoLinhaDadosParcelaPagto> vParcelaPagto;
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
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
						vParcelaPagto = geraDadosParcelasPagto(_listaNumeroPedidoSelecionado, out msgSolicitacaoConfirmacao);
						if (msgSolicitacaoConfirmacao.Length > 0)
						{
							strMsg = msgSolicitacaoConfirmacao +
									"\n" +
									"Digite a senha para confirmar que deseja prosseguir mesmo assim!";
							fAutorizacao = new FAutorizacao(strMsg);
							while (true)
							{
								drAutorizacao = fAutorizacao.ShowDialog();
								if (drAutorizacao != DialogResult.OK)
								{
									avisoErro("Operação cancelada!");
									return;
								}

								if (fAutorizacao.senha.ToUpper() != Global.Usuario.senhaDescriptografada.ToUpper())
								{
									avisoErro("Senha inválida!");
								}
								else
								{
									break;
								}
							}
						}

						boletoAvulsoComPedidoSelecionado = geraDadosBoletoAvulsoComPedido(_listaNumeroPedidoSelecionado, vParcelaPagto);
						clienteSelecionado = ClienteDAO.getCliente(boletoAvulsoComPedidoSelecionado.id_cliente);
						_blnRegistroFoiGravado = false;
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
					for (int i = 0; i < listaNumeroPedidoSelecionado.Count; i++)
					{
						blnAchou = false;
						strPedido = listaNumeroPedidoSelecionado[i];
						for (int j = 0; j < lbPedido.Items.Count; j++)
						{
							if (lbPedido.Items[j].ToString().Equals(strPedido))
							{
								blnAchou = true;
								break;
							}
						}
						if (!blnAchou) lbPedido.Items.Add(strPedido);
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
					txtNumeroDocumento.Text = "";
					#endregion

					#region [ Dados do cliente ]

					#region [ Prioridade em usar dados do endereço memorizado no pedido, se houver ]
					foreach (Pedido p in _listaPedidos)
					{
						if (p.st_memorizacao_completa_enderecos != 0)
						{
							boletoCliente = new BoletoCliente();
							boletoCliente.nome = p.endereco_nome.ToUpper();
							boletoCliente.cnpj_cpf = p.endereco_cnpj_cpf;
							boletoCliente.email = p.endereco_email.ToLower();
							boletoCliente.endereco_logradouro = p.endereco_logradouro;
							boletoCliente.endereco_numero = p.endereco_numero;
							boletoCliente.endereco_complemento = p.endereco_complemento;
							boletoCliente.endereco_cep = p.endereco_cep;
							boletoCliente.endereco_bairro = p.endereco_bairro.ToUpper();
							boletoCliente.endereco_cidade = p.endereco_cidade.ToUpper();
							boletoCliente.endereco_uf = p.endereco_uf.ToUpper();

							break;
						}
					}
					#endregion

					#region [ Se não houver dados do endereço memorizado no pedido, usa dados do cadastro do cliente ]
					if (boletoCliente == null)
					{
						boletoCliente = new BoletoCliente();
						boletoCliente.nome = clienteSelecionado.nome.ToUpper();
						boletoCliente.cnpj_cpf = clienteSelecionado.cnpj_cpf;
						boletoCliente.email = clienteSelecionado.email.ToLower();
						boletoCliente.endereco_logradouro = clienteSelecionado.endereco;
						boletoCliente.endereco_numero = clienteSelecionado.endereco_numero;
						boletoCliente.endereco_complemento = clienteSelecionado.endereco_complemento;
						boletoCliente.endereco_cep = clienteSelecionado.cep;
						boletoCliente.endereco_bairro = clienteSelecionado.bairro.ToUpper();
						boletoCliente.endereco_cidade = clienteSelecionado.cidade.ToUpper();
						boletoCliente.endereco_uf = clienteSelecionado.uf.ToUpper();
					}
					#endregion

					#region [ Exibe os dados ]
					if (boletoCliente != null)
					{
						strEndereco = boletoCliente.endereco_logradouro;
						if (boletoCliente.endereco_numero.Length > 0) strEndereco += ", " + boletoCliente.endereco_numero;
						if (boletoCliente.endereco_complemento.Length > 0) strEndereco += " " + boletoCliente.endereco_complemento;

						txtClienteNome.Text = boletoCliente.nome;
						txtClienteCnpjCpf.Text = Global.formataCnpjCpf(boletoCliente.cnpj_cpf);
						txtEndereco.Text = strEndereco.ToUpper();
						txtCep.Text = Global.formataCep(boletoCliente.endereco_cep);
						txtBairro.Text = boletoCliente.endereco_bairro;
						txtCidade.Text = boletoCliente.endereco_cidade;
						txtUF.Text = boletoCliente.endereco_uf;
						txtEmail.Text = boletoCliente.email;
					}
					#endregion

					#endregion

					#region [ Dados das parcelas do(s) pedido(s) ]
					if (boletoAvulsoComPedidoSelecionado.listaItem.Count > 0) grdParcelasBase.Rows.Add(boletoAvulsoComPedidoSelecionado.listaItem.Count);
					intIndiceLinhaGrid = 0;
					for (int i = 0; i < boletoAvulsoComPedidoSelecionado.listaItem.Count; i++)
					{
						strDadosRateio = "";
						for (int j = 0; j < boletoAvulsoComPedidoSelecionado.listaItem[i].listaRateio.Count; j++)
						{
							strDadosRateioParcela = boletoAvulsoComPedidoSelecionado.listaItem[i].listaRateio[j].pedido + "=" + Global.formataMoeda(boletoAvulsoComPedidoSelecionado.listaItem[i].listaRateio[j].valor);
							if (strDadosRateio.Length > 0) strDadosRateio += "|";
							strDadosRateio += strDadosRateioParcela;
						}
						grdParcelasBase.Rows[intIndiceLinhaGrid].Cells["grdParcelasBase_num_parcela"].Value = (intIndiceLinhaGrid + 1).ToString() + " / " + boletoAvulsoComPedidoSelecionado.listaItem.Count.ToString();
						grdParcelasBase.Rows[intIndiceLinhaGrid].Cells["grdParcelasBase_forma_pagto"].Value = Global.formaPagtoPedidoDescricao(boletoAvulsoComPedidoSelecionado.listaItem[i].forma_pagto);
						grdParcelasBase.Rows[intIndiceLinhaGrid].Cells["grdParcelasBase_dt_vencto"].Value = Global.formataDataDdMmYyyyComSeparador(boletoAvulsoComPedidoSelecionado.listaItem[i].dt_vencto);
						grdParcelasBase.Rows[intIndiceLinhaGrid].Cells["grdParcelasBase_valor"].Value = Global.formataMoeda(boletoAvulsoComPedidoSelecionado.listaItem[i].valor);
						grdParcelasBase.Rows[intIndiceLinhaGrid].Cells["grdParcelasBase_dados_rateio"].Value = strDadosRateio;
						vlTotalParcelasBase += boletoAvulsoComPedidoSelecionado.listaItem[i].valor;
						intIndiceLinhaGrid++;
					}

					ajustaPosicaoLblTotalGridParcelasBase();
					lblTotalGridParcelasBase.Text = Global.formataMoeda(vlTotalParcelasBase);

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < grdParcelasBase.Rows.Count; i++)
					{
						if (grdParcelasBase.Rows[i].Selected) grdParcelasBase.Rows[i].Selected = false;
					}
					#endregion

					#endregion

					#region [ Dados das parcelas dos boletos a gerar ]
					intIndiceLinhaGrid = 0;
					for (int i = 0; i < boletoAvulsoComPedidoSelecionado.listaItem.Count; i++)
					{
						if ((boletoAvulsoComPedidoSelecionado.listaItem[i].forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO) 
							|| (boletoAvulsoComPedidoSelecionado.listaItem[i].forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO_AV))
						{
							intQtdeParcelasBoleto++;
						}
					}

					for (int i = 0; i < boletoAvulsoComPedidoSelecionado.listaItem.Count; i++)
					{
						if ((boletoAvulsoComPedidoSelecionado.listaItem[i].forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
							|| (boletoAvulsoComPedidoSelecionado.listaItem[i].forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO_AV))
						{
							grdParcelas.Rows.Add();
							strDadosRateio = "";
							for (int j = 0; j < boletoAvulsoComPedidoSelecionado.listaItem[i].listaRateio.Count; j++)
							{
								strDadosRateioParcela = boletoAvulsoComPedidoSelecionado.listaItem[i].listaRateio[j].pedido + "=" + Global.formataMoeda(boletoAvulsoComPedidoSelecionado.listaItem[i].listaRateio[j].valor);
								if (strDadosRateio.Length > 0) strDadosRateio += "|";
								strDadosRateio += strDadosRateioParcela;
							}
							grdParcelas.Rows[intIndiceLinhaGrid].Cells["grdParcelas_num_parcela"].Value = (intIndiceLinhaGrid + 1).ToString() + " / " + intQtdeParcelasBoleto;
							grdParcelas.Rows[intIndiceLinhaGrid].Cells["grdParcelas_dt_vencto"].Value = Global.formataDataDdMmYyyyComSeparador(boletoAvulsoComPedidoSelecionado.listaItem[i].dt_vencto);
							grdParcelas.Rows[intIndiceLinhaGrid].Cells["grdParcelas_valor"].Value = Global.formataMoeda(boletoAvulsoComPedidoSelecionado.listaItem[i].valor);
							grdParcelas.Rows[intIndiceLinhaGrid].Cells["grdParcelas_dados_rateio"].Value = strDadosRateio;
							vlTotalParcelasBoleto += boletoAvulsoComPedidoSelecionado.listaItem[i].valor;
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

		#region [ FBoletoAvulsoComPedidoCadDetalhe_FormClosing ]
		private void FBoletoAvulsoComPedidoCadDetalhe_FormClosing(object sender, FormClosingEventArgs e)
		{
			#region [ Declarações ]
			Boleto boletoEditado;
			#endregion

			try
			{
				#region [ Trata situação em que o boleto foi cadastrado ]
				if (_blnRegistroFoiGravado)
				{
					return;
				}
				#endregion

				#region [ Verifica se houve alterações ]
				if (_InicializacaoOk && (!_OcorreuExceptionNaInicializacao))
				{
					boletoEditado = obtemDadosBoletoCamposTela();
					if (boletoEditado != null)
					{
						if (isBoletoEditado(boletoAvulsoComPedidoSelecionado, boletoEditado))
						{
							if (!confirma("As alterações serão perdidas!!\nContinua assim mesmo?"))
							{
								e.Cancel = true;
								return;
							}
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

		#region [ rbNumeroBoletoAvulso ]

		#region [ rbNumeroBoletoAvulso_CheckedChanged ]
		private void rbNumeroBoletoAvulso_CheckedChanged(object sender, EventArgs e)
		{
			#region [ Declarações ]
			String strMsgErro = "";
			int intNsuNovo = 0;
			#endregion

			if (rbNumeroBoletoAvulso.Checked)
			{
				if (_numeroDocumentoBoletoAvulso == 0)
				{
					#region [ Gera o nº para o boleto avulso ]
					if (!BD.geraNsu(Global.Cte.FIN.NSU.NSU_BOLETO_AVULSO_NUMERO_DOCUMENTO, ref intNsuNovo, ref strMsgErro))
					{
						strMsgErro = "Falha ao tentar gerar automaticamente o número do documento para o boleto avulso!!" +
									 "\n" +
									 strMsgErro;
						avisoErro(strMsgErro);
						return;
					}
					_numeroDocumentoBoletoAvulso = intNsuNovo;
					#endregion
				}

				txtNumeroDocumento.Text = _numeroDocumentoBoletoAvulso.ToString();
				txtNumeroDocumento.ReadOnly = true;
			}
			else
			{
				#region [ Memoriza a informação ]
				_numeroDocumentoBoletoAvulso = (int)Global.converteInteiro(txtNumeroDocumento.Text);
				#endregion
			}
		}
		#endregion

		#endregion

		#region [ rbNumeroNF ]

		#region [ rbNumeroNF_CheckedChanged ]
		private void rbNumeroNF_CheckedChanged(object sender, EventArgs e)
		{
			if (rbNumeroNF.Checked)
			{
				if (_numeroNF > 0)
				{
					txtNumeroDocumento.Text = _numeroNF.ToString();
				}
				else
				{
					txtNumeroDocumento.Text = "";
				}

				txtNumeroDocumento.ReadOnly = false;
				txtNumeroDocumento.Focus();
			}
			else
			{
				#region [ Memoriza a informação ]
				_numeroNF = (int)Global.converteInteiro(txtNumeroDocumento.Text);
				#endregion
			}
		}
		#endregion

		#endregion

		#region [ txtNumeroDocumento ]

		#region [ txtNumeroDocumento_Enter ]
		private void txtNumeroDocumento_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtNumeroDocumento_Leave ]
		private void txtNumeroDocumento_Leave(object sender, EventArgs e)
		{
			txtNumeroDocumento.Text = txtNumeroDocumento.Text.Trim();
		}
		#endregion

		#region [ txtNumeroDocumento_KeyDown ]
		private void txtNumeroDocumento_KeyDown(object sender, KeyEventArgs e)
		{
			Global.trataTextBoxKeyDown(sender, e, txtJurosMora);
		}
		#endregion

		#region [ txtNumeroDocumento_KeyPress ]
		private void txtNumeroDocumento_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.KeyChar = Global.filtraDigitacaoNumeroInteiro(e.KeyChar);
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

		#endregion

		#region [ Botões/Menu ]

		#region [ Cadastrar ]

		#region [ btnCadastrar_Click ]
		private void btnCadastrar_Click(object sender, EventArgs e)
		{
			trataBotaoCadastrar();
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

	#region [ TipoLinhaDadosParcelaPagto ]
	public class TipoLinhaDadosParcelaPagto
	{
		#region [ Atributos / Getters-Setters ]

		private int _intNumDestaParcela;
		public int intNumDestaParcela
		{
			get { return _intNumDestaParcela; }
			set { _intNumDestaParcela = value; }
		}

		private int _intNumTotalParcelas;
		public int intNumTotalParcelas
		{
			get { return _intNumTotalParcelas; }
			set { _intNumTotalParcelas = value; }
		}

		private short _id_forma_pagto;
		public short id_forma_pagto
		{
			get { return _id_forma_pagto; }
			set { _id_forma_pagto = value; }
		}

		private DateTime _dtVencto;
		public DateTime dtVencto
		{
			get { return _dtVencto; }
			set { _dtVencto = value; }
		}

		private decimal _vlValor;
		public decimal vlValor
		{
			get { return _vlValor; }
			set { _vlValor = value; }
		}

		private String _strDadosRateio;
		public String strDadosRateio
		{
			get { return _strDadosRateio; }
			set { _strDadosRateio = value; }
		}

		#endregion

		#region [ Construtor ]
		public TipoLinhaDadosParcelaPagto()
		{
			_strDadosRateio = "";
			_id_forma_pagto = 0;
			_dtVencto = DateTime.MinValue;
		}
		#endregion
	}
	#endregion

	#region [ TipoPedidoCalculoParcelasBoleto ]
	public class TipoPedidoCalculoParcelasBoleto
	{
		#region [ Atributos / Getters-Setters ]

		private String _pedido;
		public String pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private decimal _vlTotalFamiliaPedidos;
		public decimal vlTotalFamiliaPedidos
		{
			get { return _vlTotalFamiliaPedidos; }
			set { _vlTotalFamiliaPedidos = value; }
		}

		private decimal _vlTotalDestePedido;
		public decimal vlTotalDestePedido
		{
			get { return _vlTotalDestePedido; }
			set { _vlTotalDestePedido = value; }
		}

		private decimal _razaoValorPedidoFilhote;
		public decimal razaoValorPedidoFilhote
		{
			get { return _razaoValorPedidoFilhote; }
			set { _razaoValorPedidoFilhote = value; }
		}

		private short _tipo_parcelamento;
		public short tipo_parcelamento
		{
			get { return _tipo_parcelamento; }
			set { _tipo_parcelamento = value; }
		}

		private short _av_forma_pagto;
		public short av_forma_pagto
		{
			get { return _av_forma_pagto; }
			set { _av_forma_pagto = value; }
		}

		private short _pc_qtde_parcelas;
		public short pc_qtde_parcelas
		{
			get { return _pc_qtde_parcelas; }
			set { _pc_qtde_parcelas = value; }
		}

		private decimal _pc_valor_parcela;
		public decimal pc_valor_parcela
		{
			get { return _pc_valor_parcela; }
			set { _pc_valor_parcela = value; }
		}

		private short _pc_maquineta_qtde_parcelas;
		public short pc_maquineta_qtde_parcelas
		{
			get { return _pc_maquineta_qtde_parcelas; }
			set { _pc_maquineta_qtde_parcelas = value; }
		}

		private decimal _pc_maquineta_valor_parcela;
		public decimal pc_maquineta_valor_parcela
		{
			get { return _pc_maquineta_valor_parcela; }
			set { _pc_maquineta_valor_parcela = value; }
		}

		private short _pce_forma_pagto_entrada;
		public short pce_forma_pagto_entrada
		{
			get { return _pce_forma_pagto_entrada; }
			set { _pce_forma_pagto_entrada = value; }
		}

		private short _pce_forma_pagto_prestacao;
		public short pce_forma_pagto_prestacao
		{
			get { return _pce_forma_pagto_prestacao; }
			set { _pce_forma_pagto_prestacao = value; }
		}

		private decimal _pce_entrada_valor;
		public decimal pce_entrada_valor
		{
			get { return _pce_entrada_valor; }
			set { _pce_entrada_valor = value; }
		}

		private short _pce_prestacao_qtde;
		public short pce_prestacao_qtde
		{
			get { return _pce_prestacao_qtde; }
			set { _pce_prestacao_qtde = value; }
		}

		private decimal _pce_prestacao_valor;
		public decimal pce_prestacao_valor
		{
			get { return _pce_prestacao_valor; }
			set { _pce_prestacao_valor = value; }
		}

		private short _pce_prestacao_periodo;
		public short pce_prestacao_periodo
		{
			get { return _pce_prestacao_periodo; }
			set { _pce_prestacao_periodo = value; }
		}

		private short _pse_forma_pagto_prim_prest;
		public short pse_forma_pagto_prim_prest
		{
			get { return _pse_forma_pagto_prim_prest; }
			set { _pse_forma_pagto_prim_prest = value; }
		}

		private short _pse_forma_pagto_demais_prest;
		public short pse_forma_pagto_demais_prest
		{
			get { return _pse_forma_pagto_demais_prest; }
			set { _pse_forma_pagto_demais_prest = value; }
		}

		private decimal _pse_prim_prest_valor;
		public decimal pse_prim_prest_valor
		{
			get { return _pse_prim_prest_valor; }
			set { _pse_prim_prest_valor = value; }
		}

		private short _pse_prim_prest_apos;
		public short pse_prim_prest_apos
		{
			get { return _pse_prim_prest_apos; }
			set { _pse_prim_prest_apos = value; }
		}

		private short _pse_demais_prest_qtde;
		public short pse_demais_prest_qtde
		{
			get { return _pse_demais_prest_qtde; }
			set { _pse_demais_prest_qtde = value; }
		}

		private decimal _pse_demais_prest_valor;
		public decimal pse_demais_prest_valor
		{
			get { return _pse_demais_prest_valor; }
			set { _pse_demais_prest_valor = value; }
		}

		private short _pse_demais_prest_periodo;
		public short pse_demais_prest_periodo
		{
			get { return _pse_demais_prest_periodo; }
			set { _pse_demais_prest_periodo = value; }
		}

		private short _pu_forma_pagto;
		public short pu_forma_pagto
		{
			get { return _pu_forma_pagto; }
			set { _pu_forma_pagto = value; }
		}

		private decimal _pu_valor;
		public decimal pu_valor
		{
			get { return _pu_valor; }
			set { _pu_valor = value; }
		}

		private short _pu_vencto_apos;
		public short pu_vencto_apos
		{
			get { return _pu_vencto_apos; }
			set { _pu_vencto_apos = value; }
		}

		#endregion
	}
	#endregion
}
