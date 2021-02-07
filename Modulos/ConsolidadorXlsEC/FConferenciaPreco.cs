using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConsolidadorXlsEC
{
	public partial class FConferenciaPreco : FModelo
	{
		#region [ Atributos ]
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

		private string _tituloBoxDisplayInformativo = "Mensagens Informativas";
		private int _qtdeMsgDisplayInformativo = 0;
		private string _tituloBoxDisplayErro = "Mensagens de Erro";
		private int _qtdeMsgDisplayErro = 0;
		string MARGEM_MSG_NIVEL_2 = new string(' ', 8);
		#endregion

		#region [ Constantes ]
		// Obs: a coluna 'ColVisibleOrdenacaoPadrao' é a coluna visível usada p/ poder ser clicada e fazer a ordenação conforme o padrão inicial, sendo que as células dessa coluna ficam vazias.
		// E a coluna 'ColHiddenValorOrdenacaoPadrao' é a coluna invisível que possui os dados usados p/ a ordenação padrão.
		const string GRID_COL_VISIBLE_ORDENACAO_PADRAO = "ColVisibleOrdenacaoPadrao";
		const string GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO = "ColHiddenValorOrdenacaoPadrao";
		const string GRID_COL_SKU = "SKU";
		const string GRID_COL_DESCRICAO = "Descricao";
		const string GRID_COL_PRECO_MAGENTO = "PrecoMagento";
		const string GRID_COL_PRECO_CSV = "PrecoCSV";
		const string GRID_COL_DIFERENCA_VALOR = "DiferencaValor";
		const string GRID_COL_DIFERENCA_PERC = "DiferencaPerc";
		#endregion

		#region [ Construtor ]
		public FConferenciaPreco()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ adicionaDisplay ]
		private void adicionaDisplay(String mensagem)
		{
			String strMensagem;
			_qtdeMsgDisplayInformativo++;
			strMensagem = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ":  " + mensagem;
			foreach (string linha in strMensagem.Split('\n'))
			{
				lbMensagem.Items.Add(linha);
			}
			lbMensagem.SelectedIndex = lbMensagem.Items.Count - 1;
			gboxMensagensInformativas.Text = _tituloBoxDisplayInformativo + "  (" + _qtdeMsgDisplayInformativo.ToString() + ")";
			Global.gravaLogAtividade(mensagem);
		}
		#endregion

		#region [ adicionaErro ]
		private void adicionaErro(String mensagem)
		{
			String strMensagem;
			_qtdeMsgDisplayErro++;
			strMensagem = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ":  " + mensagem;
			foreach (string linha in strMensagem.Split('\n'))
			{
				lbErro.Items.Add(linha);
			}
			lbErro.SelectedIndex = lbErro.Items.Count - 1;
			gboxMsgErro.Text = _tituloBoxDisplayErro + "  (" + _qtdeMsgDisplayErro.ToString() + ")";
			Global.gravaLogAtividade("ERRO: " + mensagem);
		}
		#endregion

		#region [ pathArquivoValorDefault ]
		private String pathArquivoValorDefault()
		{
			String strResp = "";

			try
			{
				strResp = Path.GetPathRoot(Application.StartupPath);
			}
			catch (Exception)
			{
				strResp = "";
			}

			if (strResp.Length == 0) strResp = @"\";
			if (Global.Usuario.Defaults.FConferenciaPreco.pathArquivo.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FConferenciaPreco.pathArquivo))
				{
					strResp = Global.Usuario.Defaults.FConferenciaPreco.pathArquivo;
				}
			}
			return strResp;
		}
		#endregion

		#region [ fileNameArquivoValorDefault ]
		private String fileNameArquivoValorDefault()
		{
			String strResp = "";

			if ((Global.Usuario.Defaults.FConferenciaPreco.fileNameArquivo ?? "").Length > 0)
			{
				if (File.Exists(Global.Usuario.Defaults.FConferenciaPreco.pathArquivo + "\\" + Global.Usuario.Defaults.FConferenciaPreco.fileNameArquivo))
				{
					strResp = Global.Usuario.Defaults.FConferenciaPreco.fileNameArquivo;
				}
			}
			return strResp;
		}
		#endregion

		#region [ limpaCamposMensagem ]
		private void limpaCamposMensagem()
		{
			lbMensagem.Items.Clear();
			_qtdeMsgDisplayInformativo = 0;
			gboxMensagensInformativas.Text = _tituloBoxDisplayInformativo;

			lbErro.Items.Clear();
			_qtdeMsgDisplayErro = 0;
			gboxMsgErro.Text = _tituloBoxDisplayErro;
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Arquivo ]
			if (txtArquivo.Text.Trim().Length == 0)
			{
				avisoErro("É necessário selecionar o arquivo CSV que será analisado!!");
				return false;
			}
			if (!File.Exists(txtArquivo.Text))
			{
				avisoErro("O arquivo CSV informado não existe!!");
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ montaCampoOrdenacao ]
		private string montaCampoOrdenacao(int idxPrioridade, decimal valor, string sku)
		{
			#region [ Declarações ]
			string campo;
			#endregion

			campo = Global.normalizaCodigo(idxPrioridade.ToString(), 2) +
					"|" +
					Global.normalizaCodigo(Global.digitos(Global.formataMoeda(Math.Abs(valor))), 18) +
					"|" +
					Global.normalizaCodigoProduto(sku);
			return campo;
		}
		#endregion

		#region [ trataBotaoAbreArquivo ]
		private void trataBotaoAbreArquivo()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "trataBotaoAbreArquivo()";
			string strNomeArquivo;
			#endregion

			strNomeArquivo = txtArquivo.Text.Trim();

			if (strNomeArquivo.Length == 0) return;

			if (!File.Exists(strNomeArquivo))
			{
				avisoErro("Arquivo '" + Path.GetFileName(strNomeArquivo) + "' não foi encontrado!!");
				return;
			}

			try
			{
				Process.Start(strNomeArquivo);
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": exception ao tentar abrir o arquivo '" + Path.GetFileName(strNomeArquivo) + "'\r\n" + ex.ToString());
			}
		}
		#endregion

		#region [ trataBotaoSelecionaArquivo ]
		private void trataBotaoSelecionaArquivo()
		{
			#region [ Declarações ]
			DialogResult dr;
			#endregion

			try
			{
				openFileDialogCtrl.InitialDirectory = pathArquivoValorDefault();
				openFileDialogCtrl.FileName = fileNameArquivoValorDefault();
				dr = openFileDialogCtrl.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				if ((openFileDialogCtrl.FileName.Length > 0) && (txtArquivo.Text.Length > 0))
				{
					if (openFileDialogCtrl.FileName.ToUpper().Equals(txtArquivo.Text.ToUpper())) return;
				}
				#endregion

				#region [ Limpa campos de mensagens ]
				limpaCamposMensagem();
				#endregion

				txtArquivo.Text = openFileDialogCtrl.FileName;
				Global.Usuario.Defaults.FConferenciaPreco.pathArquivo = Path.GetDirectoryName(openFileDialogCtrl.FileName);
				Global.Usuario.Defaults.FConferenciaPreco.fileNameArquivo = Path.GetFileName(openFileDialogCtrl.FileName);
			}
			catch (Exception ex)
			{
				info(ModoExibicaoMensagemRodape.Normal);
				avisoErro(ex.ToString());
				Close();
				return;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ processaConsultaMagentoSoapApi ]
		private bool processaConsultaMagentoSoapApi(ref List<ProdutoConferePreco> vConferePreco, out string msg_erro)
		{
			#region [ Declarações ]
			bool blnEnviouOk;
			string sku;
			string msg_erro_aux;
			string strMsg;
			string strMsgErro;
			string xmlReqSoap;
			string xmlRespSoap;
			string magentoSessionId;
			string sPercProgresso;
			ProductList productList;
			List<ProductList> vProductList;
			ProductInfo productInfo;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consulta dados no Magento v1 via API SOAP ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "Consultando relação de produtos cadastrados no Magento");

					#region [ Login no Magento ]
					xmlReqSoap = Magento.montaRequisicaoLogin(Global.Cte.Magento.USER_NAME, Global.Cte.Magento.PASSWORD);
					blnEnviouOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.login, out xmlRespSoap, out msg_erro_aux);
					if (!blnEnviouOk)
					{
						strMsgErro = "Falha ao tentar realizar o login no Magento!!" + (msg_erro_aux.Length > 0 ? "\r\n" + msg_erro_aux : "");
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return false;
					}

					magentoSessionId = Magento.obtemSessionIdFromLoginResponse(xmlRespSoap, out msg_erro_aux);
					if (magentoSessionId.Length == 0)
					{
						strMsgErro = "Falha ao tentar obter o SessionId";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return false;
					}

					strMsg = "Sessão iniciada com sucesso no Magento";
					adicionaDisplay(strMsg);
					#endregion

					try // Finally: executa EndSession no Magento
					{
						#region [ Obtém a relação de produtos no Magento ]
						adicionaDisplay("Obtendo a relação de produtos cadastrados no Magento");
						xmlReqSoap = Magento.montaRequisicaoCallCatalogProductList(magentoSessionId);
						blnEnviouOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.call, out xmlRespSoap, out msg_erro_aux);
						if (!blnEnviouOk)
						{
							strMsgErro = "Falha ao tentar obter a relação de produtos cadastrados no Magento!!" + (msg_erro_aux.Length > 0 ? "\r\n" + msg_erro_aux : "");
							adicionaErro(strMsgErro);
							avisoErro(strMsgErro);
							return false;
						}
						vProductList = Magento.decodificaXmlCatalogProductListResponse(xmlRespSoap, out msg_erro_aux);
						#endregion

						#region [ A partir do SKU, obtém o valor de product_id ]
						info(ModoExibicaoMensagemRodape.EmExecucao, "Analisando relação de produtos do Magento");
						for (int i = 0; i < vConferePreco.Count; i++)
						{
							try
							{
								// O uso de lambda expression com variáveis passadas por parâmetro gera o seguinte erro:
								// Error	CS1628	Cannot use ref, out, or in parameter 'vConferePreco' inside an anonymous method, lambda expression, query expression, or local function
								sku = vConferePreco[i].sku;
								productList = vProductList.Single(p => p.sku.Equals(sku));
							}
							catch (Exception)
							{
								// O método Single() lança uma exception se houver 0 (zero) ou mais do que 1 elemento no resultado
								productList = null;
							}

							if (productList == null)
							{
								strMsg = "O SKU " + vConferePreco[i].sku + " não foi encontrado na relação de produtos do Magento!!";
								adicionaErro(strMsg);
							}
							else
							{
								vConferePreco[i].isCadastradoMagento = true;
								vConferePreco[i].product_id = productList.product_id;
								vConferePreco[i].name = productList.name;
							}
						}
						#endregion

						#region [ Para cada produto que foi encontrado o product_id, realiza a consulta detalhada para obter o preço no Magento ]
						for (int i = 0; i < vConferePreco.Count; i++)
						{
							if (!vConferePreco[i].isCadastradoMagento) continue;
							if (vConferePreco[i].product_id.Trim().Length == 0) continue;

							sPercProgresso = (vConferePreco.Count == 0 ? "" : Global.sqlFormataDouble(100d * ((double)(i + 1) / (double)vConferePreco.Count), 0, ',') + "%");
							strMsg = "Consultando produto no Magento: SKU " + vConferePreco[i].sku + "  (Etapa: " + (i + 1).ToString() + "/" + vConferePreco.Count.ToString() + ", Progresso: " + sPercProgresso + ")";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsg);
							adicionaDisplay(strMsg);
							xmlReqSoap = Magento.montaRequisicaoCallCatalogProductInfo(magentoSessionId, vConferePreco[i].product_id);
							blnEnviouOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.call, out xmlRespSoap, out msg_erro_aux);
							if (!blnEnviouOk)
							{
								strMsgErro = "Falha ao tentar consultar SKU " + vConferePreco[i].sku + " no Magento!!" + (msg_erro_aux.Length > 0 ? "\r\n" + msg_erro_aux : "");
								adicionaErro(strMsgErro);
								avisoErro(strMsgErro);
								return false;
							}

							productInfo = Magento.decodificaXmlCatalogProductInfoResponse(xmlRespSoap, out msg_erro_aux);
							if ((productInfo.price ?? "").Length == 0)
							{
								strMsg = "Falha ao tentar obter o preço do SKU " + vConferePreco[i].sku + " na resposta do Magento!";
								adicionaErro(strMsg);
							}
							else
							{
								vConferePreco[i].priceMagento = productInfo.price;
								vConferePreco[i].vlPriceMagento = Global.converteNumeroDecimal(vConferePreco[i].priceMagento);
							}

							Application.DoEvents();
						}
						#endregion
					}
					finally
					{
						xmlReqSoap = Magento.montaRequisicaoEndSession(magentoSessionId);
						blnEnviouOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.endSession, out xmlRespSoap, out msg_erro_aux);
						if (blnEnviouOk)
						{
							strMsg = "Sessão finalizada com sucesso no Magento";
							adicionaDisplay(strMsg);
						}
						else
						{
							strMsg = "Falha ao tentar finalizar sessão no Magento" + (msg_erro_aux.Length > 0 ? "\r\n" + MARGEM_MSG_NIVEL_2 + msg_erro_aux : "");
							adicionaErro(strMsg);
						}
					}
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					adicionaErro(ex.Message);
					avisoErro(ex.ToString());
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ processaConsultaMagento2RestApi ]
		private bool processaConsultaMagento2RestApi(ref List<ProdutoConferePreco> vConferePreco, Loja lojaLoginParameters, out string msg_erro)
		{
			#region [ Declarações ]
			bool blnEnviouOk;
			string sku;
			string strMsg;
			string msg_erro_aux;
			string strMsgErro;
			string urlParamReqRest;
			string urlBaseAddress = "";
			string respJson;
			Magento2Product product;
			List<Magento2SearchCriteriaFilterGroups> filtros;
			Magento2SearchCriteriaFilterGroups filter_group;
			Magento2SearchCriteriaFilterGroupsFilters filtro;
			Magento2ProductSearchResponse productSearchResponse;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consulta dados no Magento 2 via API REST ]
				try
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "Consultando relação de produtos cadastrados no Magento");

					#region [ Obtém a relação de produtos no Magento ]

					#region [ Monta o filtro da consulta ]
					// Lógica de operação dos filtros:
					// https://devdocs.magento.com/guides/v2.4/rest/performing-searches.html
					// The filter_groups array defines one or more filters. Each filter defines a search term, and the field, value, and condition_type of a search term must be assigned the same index number, starting with 0. Increment additional terms as needed.
					// When constructing a search, keep the following in mind:
					//		To perform a logical OR, specify multiple filters within a filter_groups.
					//		To perform a logical AND, specify multiple filter_groups.
					//		You cannot perform a logical OR across different filter_groups, such as (A AND B) OR (X AND Y). ORs can be performed only within the context of a single filter_groups.
					//		You can only search top-level attributes.
					filtros = new List<Magento2SearchCriteriaFilterGroups>();
					filter_group = new Magento2SearchCriteriaFilterGroups();
					filter_group.filters = new List<Magento2SearchCriteriaFilterGroupsFilters>();
					filtro = new Magento2SearchCriteriaFilterGroupsFilters();
					filtro.field = "type_id";
					filtro.value = "simple";
					filtro.condition_type = "eq";
					filter_group.filters.Add(filtro);
					filtro = new Magento2SearchCriteriaFilterGroupsFilters();
					filtro.field = "type_id";
					filtro.value = "virtual";
					filtro.condition_type = "eq";
					filter_group.filters.Add(filtro);
					filtros.Add(filter_group);
					#endregion

					adicionaDisplay("Obtendo a relação de produtos cadastrados no Magento");
					urlParamReqRest = Magento2.montaRequisicaoGetProducts(filtros, lojaLoginParameters.magento_api_rest_endpoint, out urlBaseAddress);
					blnEnviouOk = Magento2.enviaRequisicaoGetComRetry(urlParamReqRest, lojaLoginParameters.magento_api_rest_access_token, urlBaseAddress, out respJson, out msg_erro_aux);
					if (!blnEnviouOk)
					{
						strMsgErro = "Falha ao tentar obter a relação de produtos cadastrados no Magento via API REST!" + (msg_erro_aux.Length > 0 ? "\r\n" + msg_erro_aux : "");
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return false;
					}
					productSearchResponse = Magento2.decodificaJsonProductSearchResponse(respJson, out msg_erro_aux);
					if (productSearchResponse == null)
					{
						strMsgErro = "Falha ao analisar a resposta com a relação de produtos do Magento: não há nenhum produto ou ocorreu erro na decodificação dos dados da resposta!" + (msg_erro_aux.Length > 0 ? "\r\n" + msg_erro_aux : "");
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return false;
					}

					if (productSearchResponse.total_count == 0)
					{
						strMsgErro = "A consulta ao Magento da lista de produtos não retornou nenhum produto!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return false;
					}

					for (int i = 0; i < vConferePreco.Count; i++)
					{
						try
						{
							// O uso de lambda expression com variáveis passadas por parâmetro gera o seguinte erro:
							// Error	CS1628	Cannot use ref, out, or in parameter 'vConferePreco' inside an anonymous method, lambda expression, query expression, or local function
							sku = vConferePreco[i].sku;
							product = productSearchResponse.items.Single(p => p.sku.Equals(sku));
						}
						catch (Exception)
						{
							// O método Single() lança uma exception se houver 0 (zero) ou mais do que 1 elemento no resultado
							product = null;
						}

						if (product == null)
						{
							strMsg = "O SKU " + vConferePreco[i].sku + " não foi encontrado na relação de produtos do Magento!";
							adicionaErro(strMsg);
						}
						else
						{
							vConferePreco[i].isCadastradoMagento = true;
							vConferePreco[i].product_id = product.id;
							vConferePreco[i].name = product.name;
							vConferePreco[i].priceMagento = product.price;
							vConferePreco[i].vlPriceMagento = Global.converteNumeroDecimal(vConferePreco[i].priceMagento);
						}
					}
					#endregion
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					adicionaErro(ex.Message);
					avisoErro(ex.ToString());
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ trataBotaoIniciaProcessamento ]
		private void trataBotaoIniciaProcessamento()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FConferenciaPreco.trataBotaoIniciaProcessamento";
			int IDX_PRIORIDADE_SKU_NOVO;
			int IDX_PRIORIDADE_REDUCAO_PRECO;
			int IDX_PRIORIDADE_AUMENTO_PRECO;
			int IDX_PRIORIDADE_IGUALDADE_PRECO;
			int qtdeLinhaDadosArquivo = 0;
			int idxSku;
			int idxPrice;
			string msg_erro_aux;
			string strMsg;
			string strMsgErro;
			string strMsgErroLog = "";
			string strNomeArquivo;
			string strNomeArquivoHistorico;
			string strNomeArquivoResultadoHistorico;
			string strNomeArquivoResultadoHistoricoBR;
			string strPathHistorico;
			string linhaHeader;
			string linha;
			string sPerc;
			string[] linhasCSV;
			string[] camposHeader;
			string[] camposCSV;
			StringBuilder sbErro;
			DateTime dtInicioProcessamento;
			TimeSpan tsDuracaoProcessamento;
			ProdutoConferePreco conferePreco;
			List<ProdutoConferePreco> vConferePreco = new List<ProdutoConferePreco>();
			List<ProdutoConferePreco> vConferePrecoOrdenado;
			List<string> vResultadoHistorico;
			List<string> vResultadoHistoricoBR;
			Loja lojaLoginParameters;
			Log log = new Log();
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			#endregion

			try
			{
				limpaCamposMensagem();
				grid.Rows.Clear();
				for (int i = 0; i < grid.Columns.Count; i++)
				{
					grid.Columns[i].HeaderCell.SortGlyphDirection = SortOrder.None;
				}

				#region [ Obtém o nome do arquivo ]
				strNomeArquivo = txtArquivo.Text.Trim();
				#endregion

				#region [ Consistências ]
				if (strNomeArquivo.Length == 0)
				{
					strMsgErro = "É necessário selecionar o arquivo CSV a ser conferido!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}

				if (!File.Exists(strNomeArquivo))
				{
					strMsgErro = "O arquivo CSV não existe!!\r\n" + strNomeArquivo;
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}

				if (Global.IsFileLocked(strNomeArquivo))
				{
					strMsgErro = "O arquivo CSV '" + Path.GetFileName(strNomeArquivo) + "' está aberto e em uso!!\r\nNão é possível prosseguir com o processamento!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}
				#endregion

				#region [ Confirmação ]
				if (!confirma("Confirma a execução da conferência de preços?"))
				{
					adicionaErro("Operação cancelada!");
					return;
				}
				#endregion

				#region [ Inicialização do processamento ]
				dtInicioProcessamento = DateTime.Now;
				strMsg = "Início do processamento\r\n" +
						MARGEM_MSG_NIVEL_2 + "Arquivo: " + strNomeArquivo;
				adicionaDisplay(strMsg);

				lojaLoginParameters = FMain.contextoBD.AmbienteBase.lojaDAO.GetLoja(FMain.contextoBD.AmbienteBase.NumeroLojaArclube, out msg_erro_aux);
				if (lojaLoginParameters == null)
				{
					strMsgErro = "Falha ao tentar recuperar os parâmetros de login da API do Magento para a loja " + FMain.contextoBD.AmbienteBase.NumeroLojaArclube + "!";
					if (msg_erro_aux.Length > 0) strMsgErro += "\n" + msg_erro_aux;
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}
				#endregion

				#region [ Carrega dados do arquivo CSV ]
				try
				{
					#region [ Lê dados do arquivo ]
					info(ModoExibicaoMensagemRodape.EmExecucao, "Lendo dados do arquivo CSV");
					linhasCSV = File.ReadAllLines(strNomeArquivo, encode);
					adicionaDisplay("Registros para processar: " + Global.formataInteiro(linhasCSV.Length - 1));
					#endregion

					#region [ Verifica linha com títulos ]
					idxSku = -1;
					idxPrice = -1;
					linhaHeader = linhasCSV[0];
					camposHeader = linhaHeader.Split(';');
					for (int i = 0; i < camposHeader.Length; i++)
					{
						if (camposHeader[i].Equals("sku"))
						{
							idxSku = i;
						}
						else if (camposHeader[i].Equals("price"))
						{
							idxPrice = i;
						}
					}

					sbErro = new StringBuilder("");
					if (idxSku == -1)
					{
						sbErro.AppendLine("Não foi encontrada a coluna 'sku'!");
					}

					if (idxPrice == -1)
					{
						sbErro.AppendLine("Não foi encontrada a coluna 'price'!");
					}

					if (sbErro.Length > 0)
					{
						strMsgErro = "Falha ao analisar o header do arquivo CSV '" + Path.GetFileName(strNomeArquivo) + "'\r\n" + sbErro.ToString() + "\r\nNão é possível prosseguir com o processamento!!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion

					#region [ Verifica se possui linha de dados ]
					if (linhasCSV.Length <= 1)
					{
						strMsgErro = "Arquivo CSV '" + Path.GetFileName(strNomeArquivo) + "' não possui dados!!\r\nNão é possível prosseguir com o processamento!!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion

					#region [ Carrega dados em uma lista ]
					for (int i = 1; i < linhasCSV.Length; i++)
					{
						if (linhasCSV[i].Trim().Length == 0) continue;

						qtdeLinhaDadosArquivo++;
						camposCSV = linhasCSV[i].Split(';');

						if (camposCSV[idxSku].Trim().Length > 0)
						{
							conferePreco = new ProdutoConferePreco();
							conferePreco.sku = camposCSV[idxSku].Trim();
							conferePreco.skuFormatado = Global.normalizaCodigoProduto(conferePreco.sku);
							conferePreco.priceCsv = camposCSV[idxPrice].Trim();
							conferePreco.vlPriceCsv = Global.converteNumeroDecimal(conferePreco.priceCsv);
							vConferePreco.Add(conferePreco);
						}
					}
					#endregion
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					adicionaErro(ex.Message);
					avisoErro(ex.ToString());
					return;
				}
				#endregion

				#region [ Consulta dados no Magento via API ]
				if (lojaLoginParameters.magento_api_versao == Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON)
				{
					if (!processaConsultaMagento2RestApi(ref vConferePreco, lojaLoginParameters, out msg_erro_aux))
					{
						strMsgErro = "Falha no processamento dos dados obtidos através da API do Magento (API REST)!";
						if (msg_erro_aux.Length > 0) strMsgErro += "\n" + msg_erro_aux;
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
				}
				else
				{
					if (!processaConsultaMagentoSoapApi(ref vConferePreco, out msg_erro_aux))
					{
						strMsgErro = "Falha no processamento dos dados obtidos através da API do Magento (API SOAP)!";
						if (msg_erro_aux.Length > 0) strMsgErro += "\n" + msg_erro_aux;
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
				}
				#endregion

				#region [ Processa dados ]
				try
				{
					#region [ Compara os preços e monta o campo de ordenação ]
					info(ModoExibicaoMensagemRodape.EmExecucao, "Analisando os preços");
					// Define a prioridade na ordenação p/ exibição, lembrando que ordenação será decrescente visando exibir primeiro as maiores variações no preço
					IDX_PRIORIDADE_SKU_NOVO = 4;
					IDX_PRIORIDADE_REDUCAO_PRECO = 3;
					IDX_PRIORIDADE_AUMENTO_PRECO = 2;
					IDX_PRIORIDADE_IGUALDADE_PRECO = 1;

					for (int i = 0; i < vConferePreco.Count; i++)
					{
						if (!vConferePreco[i].isCadastradoMagento)
						{
							#region [ Produto não cadastrado no Magento ]
							vConferePreco[i].campoOrdenacao = montaCampoOrdenacao(IDX_PRIORIDADE_SKU_NOVO, 0m, vConferePreco[i].sku);
							#endregion
						}
						else
						{
							if (vConferePreco[i].vlPriceCsv < vConferePreco[i].vlPriceMagento)
							{
								#region [ Houve redução de preço ]
								vConferePreco[i].vlDiferenca = vConferePreco[i].vlPriceCsv - vConferePreco[i].vlPriceMagento;
								if (vConferePreco[i].vlPriceMagento != 0)
								{
									vConferePreco[i].percDiferenca = (double)(100m * (vConferePreco[i].vlPriceCsv - vConferePreco[i].vlPriceMagento) / vConferePreco[i].vlPriceMagento);
								}
								vConferePreco[i].campoOrdenacao = montaCampoOrdenacao(IDX_PRIORIDADE_REDUCAO_PRECO, vConferePreco[i].vlDiferenca, vConferePreco[i].sku);
								#endregion
							}
							else if (vConferePreco[i].vlPriceCsv > vConferePreco[i].vlPriceMagento)
							{
								#region [ Houve aumento de preço ]
								vConferePreco[i].vlDiferenca = vConferePreco[i].vlPriceCsv - vConferePreco[i].vlPriceMagento;
								if (vConferePreco[i].vlPriceMagento != 0)
								{
									vConferePreco[i].percDiferenca = (double)(100m * (vConferePreco[i].vlPriceCsv - vConferePreco[i].vlPriceMagento) / vConferePreco[i].vlPriceMagento);
								}
								vConferePreco[i].campoOrdenacao = montaCampoOrdenacao(IDX_PRIORIDADE_AUMENTO_PRECO, vConferePreco[i].vlDiferenca, vConferePreco[i].sku);
								#endregion
							}
							else
							{
								#region [ Preço permanece igual ]
								vConferePreco[i].vlDiferenca = vConferePreco[i].vlPriceCsv - vConferePreco[i].vlPriceMagento;
								vConferePreco[i].percDiferenca = 0d;
								vConferePreco[i].campoOrdenacao = montaCampoOrdenacao(IDX_PRIORIDADE_IGUALDADE_PRECO, 0m, vConferePreco[i].sku);
								#endregion
							}
						}
					}
					#endregion

					#region [ Ordena a lista ]
					info(ModoExibicaoMensagemRodape.EmExecucao, "Ordenando o resultado");
					vConferePrecoOrdenado = vConferePreco.OrderByDescending(o => o.campoOrdenacao).ToList();
					#endregion

					#region [ Preenche o grid ]
					info(ModoExibicaoMensagemRodape.EmExecucao, "Carregando os dados no grid");
					grid.SuspendLayout();
					grid.Rows.Add(vConferePrecoOrdenado.Count);
					for (int i = 0; i < vConferePrecoOrdenado.Count; i++)
					{
						grid.Rows[i].Cells[GRID_COL_VISIBLE_ORDENACAO_PADRAO].Value = null;
						grid.Rows[i].Cells[GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO].Value = vConferePrecoOrdenado[i].campoOrdenacao;
						grid.Rows[i].Cells[GRID_COL_SKU].Value = vConferePrecoOrdenado[i].sku;
						grid.Rows[i].Cells[GRID_COL_DESCRICAO].Value = vConferePrecoOrdenado[i].name;
						// Se o produto não foi encontrado no Magento, exibe a célula vazia ao invés de 0,00
						grid.Rows[i].Cells[GRID_COL_PRECO_MAGENTO].Value = (!vConferePrecoOrdenado[i].isCadastradoMagento ? "" : Global.formataMoeda(vConferePrecoOrdenado[i].vlPriceMagento));
						grid.Rows[i].Cells[GRID_COL_PRECO_CSV].Value = Global.formataMoeda(vConferePrecoOrdenado[i].vlPriceCsv);
						// Se o produto não foi encontrado no Magento, exibe a célula vazia ao invés de 0,00
						grid.Rows[i].Cells[GRID_COL_DIFERENCA_VALOR].Value = (!vConferePrecoOrdenado[i].isCadastradoMagento ? "" : Global.formataMoeda(vConferePrecoOrdenado[i].vlDiferenca));
						if (vConferePrecoOrdenado[i].percDiferenca == null)
						{
							grid.Rows[i].Cells[GRID_COL_DIFERENCA_PERC].Value = "";
						}
						else
						{
							grid.Rows[i].Cells[GRID_COL_DIFERENCA_PERC].Value = Global.formataPercentual((double)vConferePrecoOrdenado[i].percDiferenca) + " %";
						}

						if ((vConferePrecoOrdenado[i].vlDiferenca < 0) || (vConferePrecoOrdenado[i].percDiferenca < 0))
						{
							grid.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
							grid.Rows[i].DefaultCellStyle.SelectionForeColor = grid.Rows[i].DefaultCellStyle.ForeColor;
							grid.Rows[i].DefaultCellStyle.SelectionBackColor = Color.LightYellow;
						}
						else if ((vConferePrecoOrdenado[i].vlDiferenca > 0) || (vConferePrecoOrdenado[i].percDiferenca > 0))
						{
							grid.Rows[i].DefaultCellStyle.ForeColor = Color.Green;
							grid.Rows[i].DefaultCellStyle.SelectionForeColor = grid.Rows[i].DefaultCellStyle.ForeColor;
							grid.Rows[i].DefaultCellStyle.SelectionBackColor = Color.LightYellow;
						}
						else
						{
							grid.Rows[i].DefaultCellStyle.ForeColor = Color.Black;
							grid.Rows[i].DefaultCellStyle.SelectionForeColor = grid.Rows[i].DefaultCellStyle.ForeColor;
							grid.Rows[i].DefaultCellStyle.SelectionBackColor = Color.LightYellow;
						}
					}
					grid.ResumeLayout();
					#endregion

					#region [ Grava cópia do CSV em uma pasta de histórico ]
					info(ModoExibicaoMensagemRodape.EmExecucao, "Criando arquivo de histórico");
					strPathHistorico = Application.StartupPath + "\\Historico";
					try
					{
						if (!Directory.Exists(strPathHistorico))
						{
							Directory.CreateDirectory(strPathHistorico);
						}
					}
					catch (Exception ex)
					{
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Falha ao tentar criar o diretório para gravar o histórico do processamento!!\r\n" + ex.ToString());
					}

					if (Directory.Exists(strPathHistorico))
					{
						#region [ Faz uma cópia do arquivo CSV ]
						strNomeArquivoHistorico = strPathHistorico + "\\" +
												Path.GetFileNameWithoutExtension(strNomeArquivo) +
												"_" + Global.formataDataYyyyMmDdComSeparador(dtInicioProcessamento, "-") +
												"_" + Global.formataHoraHhMmSsComSimbolo(dtInicioProcessamento) +
												Path.GetExtension(strNomeArquivo);
						File.Copy(strNomeArquivo, strNomeArquivoHistorico);
						if (!File.Exists(strNomeArquivoHistorico))
						{
							adicionaErro("Falha ao tentar criar a cópia do arquivo CSV no histórico de processamento");
						}
						else
						{
							strMsg = "Sucesso na cópia do arquivo no histórico de processamento";
							adicionaDisplay(strMsg);
						}
						#endregion
					}
					#endregion

					#region [ Grava o resultado do processamento em um arquivo na pasta de histórico ]
					strNomeArquivoResultadoHistorico = strPathHistorico + "\\" +
												"ResultadoConferenciaPreco" +
												"_" + Global.formataDataYyyyMmDdComSeparador(dtInicioProcessamento, "-") +
												"_" + Global.formataHoraHhMmSsComSimbolo(dtInicioProcessamento) +
												".csv";
					strNomeArquivoResultadoHistoricoBR = strPathHistorico + "\\" +
												"ResultadoConferenciaPreco" +
												"_BR" +
												"_" + Global.formataDataYyyyMmDdComSeparador(dtInicioProcessamento, "-") +
												"_" + Global.formataHoraHhMmSsComSimbolo(dtInicioProcessamento) +
												".csv";
					vResultadoHistorico = new List<string>();
					vResultadoHistoricoBR = new List<string>();
					linha = "sku;descricao;preco_magento;preco_novo;dif_valor;dif_perc;";
					vResultadoHistorico.Add(linha);
					vResultadoHistoricoBR.Add(linha);
					for (int i = 0; i < vConferePrecoOrdenado.Count; i++)
					{
						#region [ Formato internacional ]
						sPerc = (vConferePrecoOrdenado[i].percDiferenca == null ? "" : Global.sqlFormataDouble((double)vConferePrecoOrdenado[i].percDiferenca, 4));
						linha = vConferePrecoOrdenado[i].sku + ";" +
								vConferePrecoOrdenado[i].name.Replace(";", ",") + ";" +
								Global.sqlFormataDecimal(vConferePrecoOrdenado[i].vlPriceMagento, 2) + ";" +
								Global.sqlFormataDecimal(vConferePrecoOrdenado[i].vlPriceCsv, 2) + ";" +
								Global.sqlFormataDecimal(vConferePrecoOrdenado[i].vlDiferenca, 2) + ";" +
								sPerc + ";";
						vResultadoHistorico.Add(linha);
						#endregion

						#region [ Formato Brasileiro ]
						sPerc = (vConferePrecoOrdenado[i].percDiferenca == null ? "" : Global.sqlFormataDouble((double)vConferePrecoOrdenado[i].percDiferenca, 4, ','));
						linha = vConferePrecoOrdenado[i].sku + ";" +
								vConferePrecoOrdenado[i].name.Replace(";", ",") + ";" +
								Global.sqlFormataDecimal(vConferePrecoOrdenado[i].vlPriceMagento, 2, ',') + ";" +
								Global.sqlFormataDecimal(vConferePrecoOrdenado[i].vlPriceCsv, 2, ',') + ";" +
								Global.sqlFormataDecimal(vConferePrecoOrdenado[i].vlDiferenca, 2, ',') + ";" +
								sPerc + ";";
						vResultadoHistoricoBR.Add(linha);
						#endregion
					}

					#region [ Grava o arquivo de resultado no formato internacional ]
					System.IO.File.WriteAllLines(strNomeArquivoResultadoHistorico, vResultadoHistorico.ToArray());
					if (!File.Exists(strNomeArquivoResultadoHistorico))
					{
						adicionaErro("Falha ao tentar gravar o resultado no histórico de processamento");
					}
					else
					{
						strMsg = "Sucesso na gravação do resultado no histórico de processamento";
						adicionaDisplay(strMsg);
					}
					#endregion

					#region [ Grava o arquivo de resultado no formato brasileiro ]
					System.IO.File.WriteAllLines(strNomeArquivoResultadoHistoricoBR, vResultadoHistoricoBR.ToArray());
					if (!File.Exists(strNomeArquivoResultadoHistoricoBR))
					{
						adicionaErro("Falha ao tentar gravar o resultado no histórico de processamento (formato BR)");
					}
					else
					{
						strMsg = "Sucesso na gravação do resultado no histórico de processamento (formato BR)";
						adicionaDisplay(strMsg);
					}
					#endregion

					#endregion
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					adicionaErro(ex.Message);
					avisoErro(ex.ToString());
					return;
				}
				#endregion

				tsDuracaoProcessamento = DateTime.Now - dtInicioProcessamento;

				#region [ Grava log ]
				strMsg = "Sucesso na conferência de preços (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + "): arquivo = " + strNomeArquivo + " (" + qtdeLinhaDadosArquivo.ToString() + " linhas de dados)";
				Global.gravaLogAtividade(strMsg);
				log.usuario = Global.Usuario.usuario;
				log.operacao = Global.Cte.CXLSEC.LogOperacao.CONFERENCIA_PRECO;
				log.complemento = strMsg;
				FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
				#endregion

				#region [ Mensagem de sucesso ]
				info(ModoExibicaoMensagemRodape.Normal);
				strMsg = "Processamento concluído com sucesso (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + ")!!";
				adicionaDisplay(strMsg);
				aviso(strMsg);
				#endregion

				grid.Focus();
			}
			catch (Exception ex)
			{
				Global.gravaLogAtividade(ex.ToString());
				adicionaErro(ex.Message);
				avisoErro(ex.ToString());
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FConferenciaPreco ]

		#region [ FConferenciaPreco_Load ]
		private void FConferenciaPreco_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

			try
			{
				txtArquivo.Text = "";
				grid.Rows.Clear();

				limpaCamposMensagem();
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

		#region [ FConferenciaPreco_Shown ]
		private void FConferenciaPreco_Shown(object sender, EventArgs e)
		{
			#region [ Declarações ]
			string strFileNameArquivo;
			#endregion

			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					btnDummy.Focus();
					#endregion

					strFileNameArquivo = pathArquivoValorDefault() + "\\" + fileNameArquivoValorDefault();
					if (File.Exists(strFileNameArquivo)) txtArquivo.Text = strFileNameArquivo;

					openFileDialogCtrl.InitialDirectory = pathArquivoValorDefault();
					openFileDialogCtrl.FileName = fileNameArquivoValorDefault();

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

		#region [ FConferenciaPreco_FormClosing ]
		private void FConferenciaPreco_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#endregion

		#region [ btnSelecionaArquivo ]

		#region [ btnSelecionaArquivo_Click ]
		private void btnSelecionaArquivo_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaArquivo();
		}

		#endregion

		#endregion

		#region [ btnAbreArquivo ]

		#region [ btnAbreArquivo_Click ]
		private void btnAbreArquivo_Click(object sender, EventArgs e)
		{
			trataBotaoAbreArquivo();
		}
		#endregion

		#endregion

		#region [ btnIniciaProcessamento ]

		#region [ btnIniciaProcessamento_Click ]
		private void btnIniciaProcessamento_Click(object sender, EventArgs e)
		{
			trataBotaoIniciaProcessamento();
		}
		#endregion

		#endregion

		#region [ txtArquivo ]

		#region [ txtArquivo_Enter ]
		private void txtArquivo_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtArquivo_DoubleClick ]
		private void txtArquivo_DoubleClick(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}

		#endregion

		#endregion

		#region [ lbMensagem ]

		#region [ lbMensagem_DoubleClick ]
		private void lbMensagem_DoubleClick(object sender, EventArgs e)
		{
			if (lbMensagem.Items.Count == 0) return;
			if (lbMensagem.SelectedIndex < 0) return;
			aviso(lbMensagem.Items[lbMensagem.SelectedIndex].ToString());
		}
		#endregion

		#endregion

		#region [ lbErro ]

		#region [ lbErro_DoubleClick ]
		private void lbErro_DoubleClick(object sender, EventArgs e)
		{
			if (lbErro.Items.Count == 0) return;
			if (lbErro.SelectedIndex < 0) return;
			aviso(lbErro.Items[lbErro.SelectedIndex].ToString());
		}
		#endregion

		#endregion

		#region [ grid ]

		#region [ grid_SortCompare ]
		private void grid_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
		{
			#region [ Declarações ]
			string sValue1;
			string sValue2;
			#endregion

			switch (e.Column.Name)
			{
				case GRID_COL_SKU:
					e.SortResult = String.Compare(Global.normalizaCodigoProduto(e.CellValue1.ToString()), Global.normalizaCodigoProduto(e.CellValue2.ToString()));
					e.Handled = true;
					break;
				case GRID_COL_DESCRICAO:
					// NOP
					break;
				case GRID_COL_PRECO_MAGENTO:
					e.SortResult = Decimal.Compare(Global.converteNumeroDecimal(e.CellValue1.ToString()), Global.converteNumeroDecimal(e.CellValue2.ToString()));
					e.Handled = true;
					break;
				case GRID_COL_PRECO_CSV:
					e.SortResult = Decimal.Compare(Global.converteNumeroDecimal(e.CellValue1.ToString()), Global.converteNumeroDecimal(e.CellValue2.ToString()));
					e.Handled = true;
					break;
				case GRID_COL_DIFERENCA_VALOR:
					e.SortResult = Decimal.Compare(Global.converteNumeroDecimal(e.CellValue1.ToString()), Global.converteNumeroDecimal(e.CellValue2.ToString()));
					e.Handled = true;
					break;
				case GRID_COL_DIFERENCA_PERC:
					sValue1 = e.CellValue1.ToString().Replace("%", "").Trim();
					sValue2 = e.CellValue2.ToString().Replace("%", "").Trim();
					e.SortResult = Decimal.Compare(Global.converteNumeroDecimal(sValue1), Global.converteNumeroDecimal(sValue2));
					e.Handled = true;
					break;
				case GRID_COL_VISIBLE_ORDENACAO_PADRAO:
					// Obs: a coluna 'ColVisibleOrdenacaoPadrao' é a coluna visível usada p/ poder ser clicada e fazer a ordenação conforme o padrão inicial, sendo que as células dessa coluna ficam vazias.
					// E a coluna 'ColHiddenValorOrdenacaoPadrao' é a coluna invisível que possui os dados usados p/ a ordenação padrão.
					sValue1 = grid.Rows[e.RowIndex1].Cells[GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO].Value.ToString();
					sValue2 = grid.Rows[e.RowIndex2].Cells[GRID_COL_HIDDEN_VALOR_ORDENACAO_PADRAO].Value.ToString();
					e.SortResult = String.Compare(sValue1, sValue2);
					e.Handled = true;
					break;
				default:
					break;
			}
		}
		#endregion

		#endregion

		#endregion
	}
}
