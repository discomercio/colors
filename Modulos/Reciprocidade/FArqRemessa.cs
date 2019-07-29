#region [ using ]
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
#endregion

namespace Reciprocidade
{
	public partial class FArqRemessa : FModelo
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

		private DataSet _dsConsulta = null;
		private ArqRemessa _arquivoRemessa = new ArqRemessa();
		#endregion

		#region [ Constantes ]
		private const int ST_GERACAO_EM_ANDAMENTO = 0;
		private const int ST_GERACAO_SUCESSO = 1;
		private const int ST_GERACAO_FALHA = 2;
		private const int ST_ENVIADO_SERASA_SUCESSO = 1;
		#endregion

		#region [ Construtor ]
		public FArqRemessa()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ descricaoOcorrencia ]
		private string descricaoOcorrencia(string identificacaoOcorrencia)
		{
			#region [ Declarações ]
			int intIdentificacaoOcorrencia;
			string strResp;
			#endregion

			if (identificacaoOcorrencia == null) return "";
			intIdentificacaoOcorrencia = (int)Global.converteInteiro(identificacaoOcorrencia);

			switch (intIdentificacaoOcorrencia)
			{
				case 2:
					strResp = "Inclusão";
					break;

				case 6:
				case 15:
					strResp = "Pagamento";
					break;

				case 9:
				case 10:
					strResp = "Baixado";
					break;

				case 12:
				case 13:
					strResp = "Alteração (valor)";
					break;

				case 14:
					strResp = "Alteração (vencto)";
					break;

				default:
					strResp = identificacaoOcorrencia;
					break;
			}

			return strResp;
		}
		#endregion

		#region [ pathTituloArquivoRemessaValorDefault ]
		private String pathTituloArquivoRemessaValorDefault()
		{
			#region [ Declarações ]
			String strResp;
			#endregion

			strResp = Global.PATH_DEFAULT_TITULO_ARQUIVO_REMESSA;
			if (Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa))
				{
					strResp = Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa;
				}
			}
			return strResp;
		}
		#endregion

		#region [ ajustaPosicaoLblTotalGridBoletos ]
		private void ajustaPosicaoLblTotalGridBoletos()
		{
			// NOP
		}
		#endregion

		#region [ limpaCamposResposta ]
		private void limpaCamposResposta()
		{
			lblTotalGridBoletos.Text = "";
			lblTotalRegistros.Text = "";
			grdBoletos.Rows.Clear();
			ajustaPosicaoLblTotalGridBoletos();
		}
		#endregion

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtDiretorio.Text = "";
			limpaCamposResposta();
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Diretório ]
			if (txtDiretorio.Text.Trim().Length == 0)
			{
				avisoErro("É necessário informar o diretório em que o arquivo de remessa será gerado!!");
				return false;
			}
			if (!Directory.Exists(txtDiretorio.Text))
			{
				avisoErro("O diretório selecionado para gerar o arquivo de remessa não existe!!");
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ trataBotaoSelecionaDiretorio ]
		private void trataBotaoSelecionaDiretorio()
		{
			DialogResult dr;
			folderBrowserDialog.SelectedPath = txtDiretorio.Text;
			dr = folderBrowserDialog.ShowDialog();
			if (dr != DialogResult.OK) return;
			txtDiretorio.Text = folderBrowserDialog.SelectedPath;
			Global.Usuario.Defaults.FArqRemessa.pathTituloArquivoRemessa = folderBrowserDialog.SelectedPath;
		}
		#endregion

		#region [ obtemPeriodoProximaRemessa ]
		private bool obtemPeriodoProximaRemessa(out DateTime dtInicioPeriodoProxRemessa, out DateTime dtFinalPeriodoProxRemessa,  out String strMsgAlerta, out String strMsgErro)
		{
			#region [ Declarações ]
			int id_arq_remessa_normal;
			DateTime dtInicialPeriodoUltRemessa;
			DateTime dtFinalPeriodoUltRemessa;
			DateTime dtGeracaoUltRemessa;
			#endregion

			#region [ Inicialização ]
			dtInicioPeriodoProxRemessa = DateTime.MinValue;
			dtFinalPeriodoProxRemessa = DateTime.MinValue;
			strMsgAlerta = "";
			strMsgErro = "";
			#endregion

			#region [ Obtém a data da última remessa ]
			if (!ArqRemessaDAO.obtemPeriodoUltRemessa(out id_arq_remessa_normal, out dtInicialPeriodoUltRemessa, out dtFinalPeriodoUltRemessa, out dtGeracaoUltRemessa, out strMsgErro))
			{
				if (strMsgErro.Length > 0)
				{
					avisoErro(strMsgErro);
					return false;
				}
			}

			if (dtFinalPeriodoUltRemessa == DateTime.MinValue) //ocorre quando nenhuma remessa foi gerada
			{
				// A data final do período (até) deve ser no máximo a do dia anterior ao do envio da remessa
				dtFinalPeriodoUltRemessa = DateTime.Now.Date.Subtract(TimeSpan.FromDays(Global.Cte.SerasaReciprocidade.TEMPO_PERIODICIDADE_ARQ_REMESSA_EM_DIAS + 1));
			}
			#endregion

			#region [ Calcula o período da próxima remessa ]
			// A data de início da próxima remessa é o dia seguinte à data final da remessa anterior
			dtInicioPeriodoProxRemessa = dtFinalPeriodoUltRemessa.AddDays(1);
			dtFinalPeriodoProxRemessa = dtInicioPeriodoProxRemessa.AddDays(Global.Cte.SerasaReciprocidade.TEMPO_PERIODICIDADE_ARQ_REMESSA_EM_DIAS - 1);
			#endregion

			#region [ Verifica se completou o ciclo para a remessa dos títulos ]
			// A data final do período (até) deve ser no máximo a do dia anterior ao do envio da remessa
			// Não podem haver lacunas nas remessas, ou seja, é necessário enviar um arquivo de remessa p/ cada período.
			// Se a periodicidade é diária, é necessário enviar um arquivo p/ cada dia. Se ocorrer uma lacuna, o arquivo
			// seguinte que for enviado será desprezado pela Serasa.
			if (dtFinalPeriodoProxRemessa >= DateTime.Now.Date)
			{
				strMsgAlerta = "Não é permitido gerar uma nova remessa porque ainda não completou o ciclo de envio desde a última remessa!!" +
								"\nA remessa anterior foi gerada em " + Global.formataDataDdMmYyyyComSeparador(dtGeracaoUltRemessa) + " com dados do período de " + Global.formataDataDdMmYyyyComSeparador(dtInicialPeriodoUltRemessa) + " a " + Global.formataDataDdMmYyyyComSeparador(dtFinalPeriodoUltRemessa) +
								"\nA próxima remessa poderá ser gerada a partir de " + Global.formataDataDdMmYyyyComSeparador(dtFinalPeriodoProxRemessa.AddDays(1));
			}
			#endregion

			return true;
		}
		#endregion

		#region [ executaConsulta ]
		private bool executaConsulta()
		{
			#region [ Declarações ]
			int intIndiceLinha = 0;
			decimal soma = 0;
			DateTime dtInicioRemessa;
			DateTime dtFinalRemessa;
			String strMsgErro;
			String strMsgAlerta;
			#endregion

			#region [ Consistência ]
			if (!consisteCampos()) return false;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "executando consulta");
			try
			{
				#region [ Limpa o grid ]
				limpaCamposResposta();
				#endregion

				#region [ Obtém o período do arquivo de remessa ]
				if (!obtemPeriodoProximaRemessa(out dtInicioRemessa, out dtFinalRemessa, out strMsgAlerta, out strMsgErro))
				{
					if (strMsgErro.Length == 0) strMsgErro = "Falha ao tentar obter o período do arquivo de remessa!!";
					avisoErro(strMsgErro);
					return false;
				}

				if (strMsgAlerta.Length > 0) aviso(strMsgAlerta);
				#endregion

				#region [ Obtém dados para gerar o arquivo de remessa ]
				_dsConsulta = TituloMovimentoDAO.selecionaBoletosParaArqRemessa(dtFinalRemessa);
				#endregion

				try
				{
					grdBoletos.SuspendLayout();

					#region [ Prepara dados p/ exibição no grid ]
					if (_dsConsulta.Tables["DtbBoleto"].Rows.Count > 0) grdBoletos.Rows.Add(_dsConsulta.Tables["DtbBoleto"].Rows.Count);
					foreach (DataRow rowBoleto in _dsConsulta.Tables["DtbBoleto"].Rows)
					{
						grdBoletos.Rows[intIndiceLinha].Cells["id_registro"].Value = BD.readToString(rowBoleto["id"]);
						grdBoletos.Rows[intIndiceLinha].Cells["cnpj"].Value = Global.formataCnpjCpf((string)rowBoleto.GetParentRow("DtbCliente_DtbBoleto")["cnpj"]);
						grdBoletos.Rows[intIndiceLinha].Cells["num_titulo"].Value = rowBoleto["nosso_numero"] + "-" + rowBoleto["digito_nosso_numero"];
						grdBoletos.Rows[intIndiceLinha].Cells["data_emissao"].Value = Global.formataDataDdMmYyyyComSeparador((DateTime)rowBoleto["dt_emissao"]);
						grdBoletos.Rows[intIndiceLinha].Cells["data_vencimento"].Value = Global.formataDataDdMmYyyyComSeparador((DateTime)rowBoleto["dt_vencto"]);
						grdBoletos.Rows[intIndiceLinha].Cells["valor"].Value = Global.formataMoeda((decimal)rowBoleto["vl_titulo"]);
						grdBoletos.Rows[intIndiceLinha].Cells["tipo_ocorrencia"].Value = BD.readToString(rowBoleto["identificacao_ocorrencia_boleto"]) + " - " + descricaoOcorrencia(BD.readToString(rowBoleto["identificacao_ocorrencia_boleto"]));

						object dtPagto = rowBoleto["dt_pagto"];
						if (dtPagto == DBNull.Value)
							grdBoletos.Rows[intIndiceLinha].Cells["data_pagamento"].Value = "";
						else
							grdBoletos.Rows[intIndiceLinha].Cells["data_pagamento"].Value = Global.formataDataDdMmYyyyComSeparador((DateTime)rowBoleto["dt_pagto"]);

						intIndiceLinha++;
						soma += BD.readToDecimal(rowBoleto["vl_titulo"]);
					}
					#endregion

					#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
					for (int i = 0; i < grdBoletos.Rows.Count; i++)
					{
						if (grdBoletos.Rows[i].Selected) grdBoletos.Rows[i].Selected = false;
					}
					#endregion
				}
				finally
				{
					grdBoletos.ResumeLayout();
				}

				ajustaPosicaoLblTotalGridBoletos();
				lblTotalGridBoletos.Text = Global.formataMoeda(soma);
				lblTotalRegistros.Text = Global.formataInteiro(intIndiceLinha);
				return true;
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoExecutaConsulta ]
		private void trataBotaoExecutaConsulta()
		{
			executaConsulta();
		}
		#endregion

		#region [ radicalCNPJSacadoJaEnviado ]
		private bool radicalCNPJSacadoJaEnviado(string radicalCNPJSacado)
		{
			bool enviado = true;
			enviado = ClienteDAO.radicalCNPJSacadoJaEnviado(radicalCNPJSacado);
			return enviado;
		}
		#endregion

		#region [ trataBotaoGravaArqRemessa ]
		private void trataBotaoGravaArqRemessa()
		{
			#region [ Declarações ]
			bool blnSucesso;
			bool blnGerouNsu;
			int qtdeRegProcessado;
			int percProgressoAtual;
			int percProgressoAnterior;
			String strMsgErro = "";
			String strMsgAlerta;
			String strMsgProgresso;
			String strNomeBasicoArqRemessa;
			String strNomeCompletoArqRemessa;
			String strPathCompleto = "";
			String strMsg;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			StreamWriter sw;
			DateTime dtInicioProcessamento;
			DateTime dtFimProcessamento;
			DateTime dtInicioRemessa;
			DateTime dtFinalRemessa;
			int totalRegTempoRelacPJ = 0;
			int totalRegTitulos = 0;
			HashSet<string> cnpjs = new HashSet<string>();
			const String TITULOS_COM_TEMPO_RELACTO_JA_ENVIADO = "titulos";
			Dictionary<String, DetalheCnpjTituloHelper> titulosPorCnpj = new Dictionary<String, DetalheCnpjTituloHelper>();
			#endregion

			#region [ Já realizou consulta p/ visualizar os dados? ]
			if (_dsConsulta == null)
			{
				avisoErro("Nenhuma consulta foi executada!!");
				return;
			}
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");
			try
			{
				#region [ Obtém o período do arquivo de remessa ]
				if (!obtemPeriodoProximaRemessa(out dtInicioRemessa, out dtFinalRemessa, out strMsgAlerta, out strMsgErro))
				{
					avisoErro("Falha ao tentar obter o período do arquivo de remessa!!\n" + strMsgErro);
					return;
				}

				if (strMsgAlerta.Length > 0)
				{
					avisoErro(strMsgAlerta);
					return;
				}
				#endregion

				#region [ Obtém dados para gerar o arquivo de remessa ]
				_dsConsulta = TituloMovimentoDAO.selecionaBoletosParaArqRemessa(dtFinalRemessa);
				#endregion
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}

			#region [ Há dados? ]
			if (_dsConsulta.Tables["DtbBoleto"].Rows.Count == 0)
			{
				strMsg = "Não há nenhum registro para o período de " + Global.formataDataDdMmYyyyComSeparador(dtInicioRemessa) + " a " + Global.formataDataDdMmYyyyComSeparador(dtFinalRemessa) + "!!" +
						"\nContinua mesmo assim?";
				if (!confirma(strMsg)) return;
			}
			#endregion

			#region [ Confirmação ]
			strMsg = "Confirma a geração do arquivo de remessa para o período de " + Global.formataDataDdMmYyyyComSeparador(dtInicioRemessa) + " a " + Global.formataDataDdMmYyyyComSeparador(dtFinalRemessa) + "?";
			if (!confirma(strMsg)) return;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "gerando arquivo de remessa");
			try
			{
				dtInicioProcessamento = DateTime.Now;

				_arquivoRemessa = new ArqRemessa();

				#region [ Prepara nome do arquivo de remessa ]
				strNomeBasicoArqRemessa = "RemessaSerasa_" +
										  Global.digitos(Global.formataDataYyyyMmDdComSeparador(DateTime.Now)) +
										  ".txt";

				#endregion

				#region [ Obtém path completo ]
				strPathCompleto = txtDiretorio.Text;

				if (!Directory.Exists(strPathCompleto))
				{
					Directory.CreateDirectory(strPathCompleto);
					if (!Directory.Exists(strPathCompleto))
					{
						avisoErro("Falha ao tentar criar o diretório:\n" + strPathCompleto);
						return;
					}
				}
				#endregion

				#region [ Nome completo do arquivo de remessa ]
				strNomeCompletoArqRemessa = Global.barraInvertidaAdd(strPathCompleto) + strNomeBasicoArqRemessa;
				#endregion

				#region [ Verifica se já existe arquivo c/ o mesmo nome ]
				if (File.Exists(strNomeCompletoArqRemessa))
				{
					avisoErro("Já existe um arquivo no diretório especificado com este nome!!\n" + strNomeCompletoArqRemessa);
					return;
				}
				#endregion

				sw = new StreamWriter(strNomeCompletoArqRemessa, true, encode);
				titulosPorCnpj.Add(TITULOS_COM_TEMPO_RELACTO_JA_ENVIADO, new DetalheCnpjTituloHelper());

				try
				{
					#region [ Monta Header ]
					ArqRemessa.LinhaHeader header = new ArqRemessa.LinhaHeader(dtInicioRemessa, dtFinalRemessa);
					_arquivoRemessa.linhaHeader = header;
					sw.WriteLine(header.ToString());
					#endregion

					#region [ Monta os registros do arquivo de remessa ]
					#region [Tempo de Relacionamento PJ]
					qtdeRegProcessado = 0;
					percProgressoAnterior = 0;
					foreach (DataRow rowBoleto in _dsConsulta.Tables["DtbBoleto"].Rows)
					{
						qtdeRegProcessado++;
						percProgressoAtual = 100 * qtdeRegProcessado / _dsConsulta.Tables["DtbBoleto"].Rows.Count;
						if (percProgressoAtual != percProgressoAnterior)
						{
							percProgressoAnterior = percProgressoAtual;
							strMsgProgresso = "Processando registros de tempo de relacionamento: " + percProgressoAtual.ToString() + "%";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
							Application.DoEvents();
						}

						string status = rowBoleto.GetParentRow("DtbCliente_DtbBoleto")["st_enviado_serasa"].ToString();
						if (status == "1") continue;

						//verifica se o radical do CNPJ já foi enviado anteriormente
						string radicalCnpjSacado = rowBoleto.GetParentRow("DtbCliente_DtbBoleto")["raiz_cnpj"].ToString();
						if (radicalCNPJSacadoJaEnviado(radicalCnpjSacado)) continue;

						string cnpjSacado = rowBoleto.GetParentRow("DtbCliente_DtbBoleto")["cnpj"].ToString();
						if (cnpjs.Contains(cnpjSacado)) continue;

						cnpjs.Add(cnpjSacado);

						int clienteId = (int)rowBoleto["id_serasa_cliente"];
						String sacadoPJ = (String)rowBoleto.GetParentRow("DtbCliente_DtbBoleto")["cnpj"];
						DateTime clienteDesde = (DateTime)rowBoleto.GetParentRow("DtbCliente_DtbBoleto")["dt_cliente_desde"];
						ArqRemessa.DetalheTempoRelacionamento tr = new ArqRemessa.DetalheTempoRelacionamento(clienteId, sacadoPJ, clienteDesde);
						_arquivoRemessa.addDetalheTempoRelacionamento(tr);

						titulosPorCnpj.Add(sacadoPJ, new DetalheCnpjTituloHelper(tr));
						totalRegTempoRelacPJ++;
					}
					#endregion

					#region [Títulos]
					qtdeRegProcessado = 0;
					percProgressoAnterior = 0;
					foreach (DataRow rowBoleto in _dsConsulta.Tables["DtbBoleto"].Rows)
					{
						qtdeRegProcessado++;
						percProgressoAtual = 100 * qtdeRegProcessado / _dsConsulta.Tables["DtbBoleto"].Rows.Count;
						if (percProgressoAtual != percProgressoAnterior)
						{
							percProgressoAnterior = percProgressoAtual;
							strMsgProgresso = "Processando registros de títulos: " + percProgressoAtual.ToString() + "%";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
							Application.DoEvents();
						}

						int tituloMovimentoId = (int)rowBoleto["id"];
						int clienteId = (int)rowBoleto["id_serasa_cliente"];
						String sacadoPJ = (String)rowBoleto.GetParentRow("DtbCliente_DtbBoleto")["cnpj"];

						StringBuilder sbNumTitulo = new StringBuilder();
						sbNumTitulo.Append(rowBoleto["nosso_numero"]).Append(rowBoleto["digito_nosso_numero"]);

						DateTime dataEmissao = (DateTime)rowBoleto["dt_emissao"];
						Decimal valorTitulo = (Decimal)rowBoleto["vl_titulo"];
						DateTime dataVecimento = (DateTime)rowBoleto["dt_vencto"];

						DateTime dataPagamento = DateTime.MinValue;
						if (rowBoleto["dt_pagto"] != DBNull.Value)
						{
							dataPagamento = (DateTime)rowBoleto["dt_pagto"];
						}

						ArqRemessa.DetalheTitulo dt = new ArqRemessa.DetalheTitulo(tituloMovimentoId, clienteId, sacadoPJ, sbNumTitulo.ToString(), dataEmissao, valorTitulo, dataVecimento, dataPagamento);

						String numOcorrencia = BD.readToString(rowBoleto["identificacao_ocorrencia_boleto"]);
						if (numOcorrencia.Trim().Equals("09")
							|| numOcorrencia.Trim().Equals("10"))
						{
							dt.isTituloBaixado = true;
						}

						_arquivoRemessa.addDetalheTitulo(dt);

						if (titulosPorCnpj.ContainsKey(sacadoPJ))
						{
							titulosPorCnpj[sacadoPJ].adicionaDetalheTitulo(dt);
						}
						else
						{
							titulosPorCnpj[TITULOS_COM_TEMPO_RELACTO_JA_ENVIADO].adicionaDetalheTitulo(dt);
						}

						totalRegTitulos++;
					}
					#endregion

					#region [ Escreve detalhes dos títulos para cada CNPJ enviado ]
					qtdeRegProcessado = 0;
					percProgressoAnterior = 0;
					foreach (KeyValuePair<String, DetalheCnpjTituloHelper> entry in titulosPorCnpj)
					{
						qtdeRegProcessado++;
						percProgressoAtual = 100 * qtdeRegProcessado / titulosPorCnpj.Count;
						if (percProgressoAtual != percProgressoAnterior)
						{
							percProgressoAnterior = percProgressoAtual;
							strMsgProgresso = "Processando dados de detalhe dos títulos: " + percProgressoAtual.ToString() + "%";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
							Application.DoEvents();
						}

						if (!entry.Key.Equals(TITULOS_COM_TEMPO_RELACTO_JA_ENVIADO))
						{
							sw.WriteLine(entry.Value.detalheRelacto);

							foreach (ArqRemessa.DetalheTitulo titulo in entry.Value.titulos)
							{
								sw.WriteLine(titulo);
							}
						}
					}

					//Escreve por último os títulos que já tiveram o CNPJ do cliente enviado
					qtdeRegProcessado = 0;
					percProgressoAnterior = 0;
					DetalheCnpjTituloHelper titulosComCnpjJaEnviado = titulosPorCnpj[TITULOS_COM_TEMPO_RELACTO_JA_ENVIADO];
					foreach (ArqRemessa.DetalheTitulo titulo in titulosComCnpjJaEnviado.titulos)
					{
						qtdeRegProcessado++;
						percProgressoAtual = 100 * qtdeRegProcessado / titulosComCnpjJaEnviado.titulos.Count;
						if (percProgressoAtual != percProgressoAnterior)
						{
							percProgressoAnterior = percProgressoAtual;
							strMsgProgresso = "Processando dados dos títulos: " + percProgressoAtual.ToString() + "%";
							info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
							Application.DoEvents();
						}

						sw.WriteLine(titulo);
					}
					#endregion
					#endregion

					#region [ Monta Trailler ]
					ArqRemessa.LinhaTrailler t = new ArqRemessa.LinhaTrailler(totalRegTempoRelacPJ, totalRegTitulos);
					_arquivoRemessa.linhaTrailler = t;
					sw.Write(t.ToString());
					#endregion
				}
				finally
				{
					sw.Flush();
					sw.Close();
				}

				dtFimProcessamento = DateTime.Now;
				int id_serasa_arq_remessa_normal = 0;

				try
				{
					BD.iniciaTransacao();

					//obtem o NSU e passa por ref na funcao abaixo
					blnGerouNsu = false;
					blnGerouNsu = BD.geraNsu(Global.Cte.t_SERASA_ARQ_REMESSA_NORMAL, ref id_serasa_arq_remessa_normal, ref strMsgErro);
					if (!blnGerouNsu)
					{
						throw new Exception("Falha ao tentar gerar o NSU para o registro de histórico de arquivos de remessa!!\n" + strMsgErro);
					}

					//insere um registro da tabela t_SERASA_ARQ_REMESSA_NORMAL
					if (!ArqRemessaDAO.insere(id_serasa_arq_remessa_normal,
											   dtInicioProcessamento,
											   dtInicioProcessamento,
											   Global.Usuario.usuario,
											   ArqRemessa.LinhaHeader.CNPJ_EMPRESA_CONVENIADA,
											   Global.formataDataYyyyMmDdSemSeparador(_arquivoRemessa.linhaHeader.dataInicio),
											   _arquivoRemessa.linhaHeader.dataInicio,
											   Global.formataDataYyyyMmDdSemSeparador(_arquivoRemessa.linhaHeader.dataFim),
											   _arquivoRemessa.linhaHeader.dataFim,
											   ArqRemessa.LinhaHeader.PERIODICIDADE_REMESSA,
											   null,
											   null,
											   ArqRemessa.LinhaHeader.ID_VERSAO_LAYOUT,
											   ArqRemessa.LinhaHeader.NUM_VERSAO_LAYOUT,
											   _arquivoRemessa.linhaTrailler.qtdeRegTempoRelacionamento,
											   _arquivoRemessa.linhaTrailler.qtdeRegTitulo,
											   dtFimProcessamento.Subtract(dtInicioProcessamento).Seconds,
											   strNomeBasicoArqRemessa,
											   txtDiretorio.Text,
											   ST_GERACAO_EM_ANDAMENTO,
											   null))
					{
						throw new Exception("Falha ao tentar inserir um registro na tabela t_SERASA_ARQ_REMESSA_NORMAL");
					}

					BD.commitTransacao();
					blnSucesso = true;
				}
				catch (Exception e)
				{
					Global.gravaLogAtividade(e.ToString());
					strMsgErro = e.ToString();
					blnSucesso = false;
				}

				if (!blnSucesso)
				{
					BD.rollbackTransacao();

					#region [ Se o arquivo de remessa foi gravado, renomeia para indicar que houve uma falha ]
					if (File.Exists(strNomeCompletoArqRemessa)) File.Move(strNomeCompletoArqRemessa, strNomeCompletoArqRemessa + ".ERR");
					#endregion

					info(ModoExibicaoMensagemRodape.Normal);
					avisoErro(strMsgErro);
				}
				else
				{
					try
					{
						BD.iniciaTransacao();

						foreach (ArqRemessa.DetalheTempoRelacionamento tempoRelacto in _arquivoRemessa.detTempoRelactoList)
						{
							//Obtem NSU
							int id_serasa_det_tempo_relac = 0;
							blnGerouNsu = false;
							blnGerouNsu = BD.geraNsu(Global.Cte.t_SERASA_REMESSA_DET_TEMPO_RELAC, ref id_serasa_det_tempo_relac, ref strMsgErro);

							if (!blnGerouNsu)
							{
								throw new Exception("Falha ao tentar gerar o NSU para o registro de tempo de relacionamento!!\n" + strMsgErro);
							}

							//insere um registro da tabela t_SERASA_REMESSA_DET_TEMPO_RELAC
							if (!DetTempoRelacDAO.insere(id_serasa_det_tempo_relac,
														 id_serasa_arq_remessa_normal,
														 tempoRelacto.clienteId,
														 ArqRemessa.DetalheTempoRelacionamento.ID,
														 tempoRelacto.cnpjCliente,
														 ArqRemessa.DetalheTempoRelacionamento.TIPO_DADOS,
														 tempoRelacto.clienteDesde,
														 tempoRelacto.tipoCliente))
							{
								throw new Exception("Falha ao criar um registro para o tempo de relacionamento com o cliente!");
							}
						}

						//Atualiza t_SERASA_CLIENTE.st_enviado_serasa
						//Atualiza t_SERASA_CLIENTE.dt_enviado_serasa
						//Atualiza t_SERASA_CLIENTE.id_serasa_arq_remessa_normal
						if (_arquivoRemessa.detTempoRelactoList.Count > 0)
						{
							foreach (ArqRemessa.DetalheTempoRelacionamento tempoRelacto in _arquivoRemessa.detTempoRelactoList)
							{
								String numCnpj = tempoRelacto.cnpjCliente;
								if (!ClienteDAO.atualizaInfoEnvioCNPJ(ST_ENVIADO_SERASA_SUCESSO,
																	DateTime.Now,
																	id_serasa_arq_remessa_normal,
																	numCnpj))
								{
									throw new Exception("Falha ao atualizar as informações de envio do CNPJ do sacado! ");
								}
							}
						}

						qtdeRegProcessado = 0;
						percProgressoAnterior = 0;
						foreach (ArqRemessa.DetalheTitulo detTitulo in _arquivoRemessa.detTituloList)
						{
							qtdeRegProcessado++;
							percProgressoAtual = 100 * qtdeRegProcessado / _arquivoRemessa.detTituloList.Count;
							if (percProgressoAtual != percProgressoAnterior)
							{
								percProgressoAnterior = percProgressoAtual;
								strMsgProgresso = "Atualizando banco de dados: " + percProgressoAtual.ToString() + "%";
								info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
								Application.DoEvents();
							}

							//Obtem NSU
							int id_serasa_det_titulo = 0;
							blnGerouNsu = false;
							blnGerouNsu = BD.geraNsu(Global.Cte.t_SERASA_REMESSA_DET_TITULO, ref id_serasa_det_titulo, ref strMsgErro);

							if (!blnGerouNsu)
							{
								throw new Exception("Falha ao tentar gerar o NSU para o registro de titulos da remessa!!\n" + strMsgErro);
							}

							//insere um registro da tabela t_SERASA_REMESSA_DET_TITULO
							if (!DetTituloDAO.insere(id_serasa_det_titulo,
													 id_serasa_arq_remessa_normal,
													 detTitulo.tituloMovimentoId,
													 detTitulo.clienteId,
													 ArqRemessa.DetalheTitulo.ID,
													 detTitulo.cnpjSacado,
													 ArqRemessa.DetalheTitulo.TIPO_DADOS,
													 detTitulo.numeroTitulo.Remove(10),
													 detTitulo.dataEmissao,
													 detTitulo.valorTitulo,
													 detTitulo.dataVencimento,
													 detTitulo.dataPagamento,
													 "D#",
													 detTitulo.numeroTitulo))
							{
								throw new Exception("Falha ao criar um registro para o título do arquivo de remessa!");
							}

							//atualiza campo t_SERASA_TITULO_MOVIMENTO.st_enviado_serasa
							//atualiza campo t_SERASA_TITULO_MOVIMENTO.id_serasa_arq_remessa_normal
							if (!TituloMovimentoDAO.atualizaStatusEnvioEIdArqRemessa(ST_ENVIADO_SERASA_SUCESSO,
																					 id_serasa_arq_remessa_normal,
																					 detTitulo.tituloMovimentoId))
							{
								throw new Exception("Falha ao atualizar o status de envio do título com ID " + detTitulo.tituloMovimentoId);
							}
						}

						if (!ArqRemessaDAO.atualizaStatusGeracao(ST_GERACAO_SUCESSO, null, id_serasa_arq_remessa_normal)) //1 = Gerado com sucesso
						{
							throw new Exception("Falha ao atualizar o status da geração do arquivo de remessa!");
						}

						blnSucesso = true;
					}
					catch (Exception e)
					{
						Global.gravaLogAtividade(e.ToString());
						strMsgErro = e.ToString();
						blnSucesso = false;
					}

					if (blnSucesso)
					{
						BD.commitTransacao();
						Global.gravaLogAtividade("Arquivo de remessa ID " + id_serasa_arq_remessa_normal + " gerado com sucesso!");
						info(ModoExibicaoMensagemRodape.Normal);
						aviso("Arquivo de remessa gerado com sucesso!!\n\n" + strNomeCompletoArqRemessa);
						Close();
					}
					else
					{
						BD.rollbackTransacao();

						if (File.Exists(strNomeCompletoArqRemessa)) File.Move(strNomeCompletoArqRemessa, strNomeCompletoArqRemessa + ".ERR");
						info(ModoExibicaoMensagemRodape.Normal);
						avisoErro("Não foi possível gerar o arquivo de remessa!");

						//tenta atualizar o status de geração do arquivo da tabela t_SERASA_ARQ_REMESSA_NORMAL
						try
						{
							BD.iniciaTransacao();
							if (!ArqRemessaDAO.atualizaStatusGeracao(ST_GERACAO_FALHA, "Falha na geração do arquivo", id_serasa_arq_remessa_normal))
							{
								throw new Exception("Falha na tentativa de atualizar o status da geração do arquivo de remessa!");
							}
							BD.commitTransacao();
						}
						catch (Exception e)
						{
							Global.gravaLogAtividade(e.ToString());
						}
					}
				}
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoCancelaEnvioTitulo ]
		private void trataBotaoCancelaEnvioTitulo()
		{
			#region [ Consistência ]
			if (grdBoletos.SelectedRows.Count == 0)
			{
				avisoErro("Nenhum registro foi selecionado!!");
				return;
			}

			if (grdBoletos.SelectedRows.Count > 1)
			{
				avisoErro("Não é permitida a seleção de múltiplos registros!!");
				return;
			}
			#endregion

			#region [ Confirma operação com o usuário ]
			if (!confirma("Confirma o cancelamento do envio do título selecionado?"))
			{
				return;
			}
			#endregion

			try
			{
				#region [ Marca o título selecionado como cancelado para envio ]
				int id = Convert.ToInt32(grdBoletos.SelectedRows[0].Cells["id_registro"].Value);
				if (!TituloMovimentoDAO.cancelaEnvio(id))
				{
					throw new Exception("Não foi possível cancelar o registro selecionado!!");
				}
				#endregion

				#region [ Atualiza o grid ]
				executaConsulta();
				#endregion
			}
			catch (Exception e)
			{
				avisoErro(e.Message);
				Global.gravaLogAtividade(e.ToString());
			}
		}
		#endregion
		#endregion

		#region [ Eventos ]
		#region [ FArqRemessa ]
		#region [ FArqRemessa_Load ]
		private void FArqRemessa_Load(object sender, EventArgs e)
		{
			bool blnSucesso = false;

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

		#region [ FArqRemessa_Shown ]
		private void FArqRemessa_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

					#region [ Preenchimento dos campos ]

					txtDiretorio.Text = pathTituloArquivoRemessaValorDefault();
					#endregion

					#region [ Ajusta o label com o valor total ]
					ajustaPosicaoLblTotalGridBoletos();
					#endregion

					#region [ Posiciona foco ]
					btnDummy.Focus();
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

		#region [ FArqRemessa_KeyDown ]
		private void FArqRemessa_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F5)
			{
				trataBotaoExecutaConsulta();
				return;
			}
		}
		#endregion

		#region [ FArqRemessa_FormClosing ]
		private void FArqRemessa_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain._fMain.Location = this.Location;
			FMain._fMain.Visible = true;
			this.Visible = false;
		}
		#endregion
		#endregion

		#region [ btnSelecionaDiretorio ]
		#region [ btnSelecionaDiretorio_Click ]
		private void btnSelecionaDiretorio_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaDiretorio();
		}
		#endregion
		#endregion

		#region [ btnExecutaConsulta ]
		#region [ btnExecutaConsulta_Click ]
		private void btnExecutaConsulta_Click(object sender, EventArgs e)
		{
			trataBotaoExecutaConsulta();
		}
		#endregion
		#endregion

		#region [ btnCancelaEnvioTitulo ]
		private void btnCancelaEnvioTitulo_Click(object sender, EventArgs e)
		{
			trataBotaoCancelaEnvioTitulo();
		}
		#endregion

		#region [ btnGravaArqRemessa ]
		#region [ btnGravaArqRemessa_Click ]
		private void btnGravaArqRemessa_Click(object sender, EventArgs e)
		{
			trataBotaoGravaArqRemessa();
		}
		#endregion
		#endregion

		#region [ txtDiretorio ]
		#region [ txtDiretorio_Enter ]
		private void txtDiretorio_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtDiretorio_DoubleClick ]
		private void txtDiretorio_DoubleClick(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion
		#endregion
		#endregion
	}
}
