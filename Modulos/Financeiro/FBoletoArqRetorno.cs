#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	public partial class FBoletoArqRetorno : Financeiro.FModelo
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

		string _numBancoArqRetorno = "";
		String[] _linhasArqRetorno;
		B237HeaderArqRetorno _b237LinhaHeader;
		B237TraillerArqRetorno _b237LinhaTrailler;
		B422HeaderArqRetorno _b422LinhaHeader;
		B422TraillerArqRetorno _b422LinhaTrailler;
		BoletoCedente _boletoCedente;
		#endregion

		#region [ Construtor ]
		public FBoletoArqRetorno()
		{
			InitializeComponent();
		}
		#endregion

		#region [ Métodos Privados ]

		#region [ adicionaDisplay ]
		private void adicionaDisplay(String mensagem)
		{
			String strMensagem;
			strMensagem = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ":  " + mensagem;
			lbMensagem.Items.Add(strMensagem.Replace('\n', ' '));
			lbMensagem.SelectedIndex = lbMensagem.Items.Count - 1;
			Global.gravaLogAtividade(mensagem);
		}
		#endregion

		#region [ adicionaErro ]
		private void adicionaErro(String mensagem)
		{
			String strMensagem;
			strMensagem = Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + ":  " + mensagem;
			lbErro.Items.Add(strMensagem.Replace('\n', ' '));
			lbErro.SelectedIndex = lbErro.Items.Count - 1;
			Global.gravaLogAtividade("ERRO: " + mensagem);
		}
		#endregion

		#region [ pathBoletoArquivoRetornoValorDefault ]
		private String pathBoletoArquivoRetornoValorDefault()
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
			if (Global.Usuario.Defaults.FBoletoArqRetorno.pathBoletoArquivoRetorno.Length > 0)
			{
				if (Directory.Exists(Global.Usuario.Defaults.FBoletoArqRetorno.pathBoletoArquivoRetorno))
				{
					strResp = Global.Usuario.Defaults.FBoletoArqRetorno.pathBoletoArquivoRetorno;
				}
			}
			return strResp;
		}
		#endregion

		#region [ limpaCampos ]
		private void limpaCampos()
		{
			txtArqRetorno.Text = "";
			grdBoletos.Rows.Clear();
			lblTotalRegistros.Text = "";
			lbMensagem.Items.Clear();
			lbErro.Items.Clear();
			lblCedenteNome.Text = "";
			lblCedenteCarteira.Text = "";
			lblCedenteAgencia.Text = "";
			lblCedenteConta.Text = "";
			lblCedenteCodigoEmpresa.Text = "";
			lblCedenteNumAvisoBancario.Text = "";
			lblCedenteDataBanco.Text = "";
			lblCedenteDataCredito.Text = "";
		}
		#endregion

		#region [ consisteCampos ]
		private bool consisteCampos()
		{
			#region [ Arquivo de Retorno ]
			if (txtArqRetorno.Text.Trim().Length == 0)
			{
				avisoErro("É necessário selecionar o arquivo de retorno que será carregado!!");
				return false;
			}
			if (!File.Exists(txtArqRetorno.Text))
			{
				avisoErro("O arquivo de retorno informado não existe!!");
				return false;
			}
			#endregion

			return true;
		}
		#endregion

		#region [ preencheCamposInfArqRetorno ]
		private void preencheCamposInfArqRetorno()
		{
			if (_numBancoArqRetorno.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
			{
				#region [ Safra ]
				lblCedenteNome.Text = _b422LinhaHeader.nomeEmpresa.valor;
				lblCedenteNumAvisoBancario.Text = _b422LinhaTrailler.numAvisoBancarioSimples.valor;
				lblCedenteDataBanco.Text = Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_b422LinhaHeader.dataGravacaoArquivo.valor));
				lblCedenteDataCredito.Text = "";
				#endregion
			}
			else
			{
				#region [ Bradesco ]
				lblCedenteNome.Text = _b237LinhaHeader.nomeEmpresa.valor;
				lblCedenteNumAvisoBancario.Text = _b237LinhaHeader.numAvisoBancario.valor;
				lblCedenteDataBanco.Text = Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_b237LinhaHeader.dataGravacaoArquivo.valor));
				lblCedenteDataCredito.Text = Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(_b237LinhaHeader.dataCredito.valor));
				#endregion
			}

			lblCedenteCarteira.Text = _boletoCedente.carteira;
			lblCedenteAgencia.Text = _boletoCedente.agencia + '-' + _boletoCedente.digito_agencia;
			lblCedenteConta.Text = _boletoCedente.conta + '-' + _boletoCedente.digito_conta;
			lblCedenteCodigoEmpresa.Text = _boletoCedente.codigo_empresa;
		}
		#endregion

		#region [ carregaGridBoletos ]
		private void carregaGridBoletos()
		{
			#region [ Declarações ]
			String[] v;
			String[] linhasArqRetorno;
			String strDescricaoMotivoOcorrencia;
			String strIdBoletoItem;
			String strNomeArquivoCompleto;
			LinhaHeaderArquivoRetorno linhaHeader = new LinhaHeaderArquivoRetorno();
			LinhaTraillerArquivoRetorno linhaTrailler = new LinhaTraillerArquivoRetorno();
			LinhaRegistroTipo1ArquivoRetorno linhaRegistro = new LinhaRegistroTipo1ArquivoRetorno();
			int intLinhaGrid = 0;
			List<Global.TipoDescricaoMotivoOcorrencia> listaMotivoOcorrencia;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			#endregion

			strNomeArquivoCompleto = txtArqRetorno.Text;

			info(ModoExibicaoMensagemRodape.EmExecucao, "lendo o arquivo de retorno");
			try
			{
				#region [ Limpa campos ]
				grdBoletos.Rows.Clear();
				lblTotalRegistros.Text = "";
				#endregion

				#region [ Carrega dados do arquivo em array ]
				linhasArqRetorno = File.ReadAllLines(strNomeArquivoCompleto, encode);
				#endregion

				#region [ Consistência do header e trailler ]
				linhaHeader.CarregaDados(linhasArqRetorno[0]);
				linhaTrailler.CarregaDados(linhasArqRetorno[linhasArqRetorno.Length - 1]);

				if ((!linhaHeader.identificacaoRegistro.valor.Equals("0")) ||
					(!linhaHeader.identificacaoArquivoRetorno.valor.Equals("2")) ||
					(!linhaHeader.literalRetorno.valor.ToUpper().Equals("RETORNO")))
				{
					avisoErro("Arquivo de retorno com header inválido!!");
					return;
				}

				if ((!linhaTrailler.identificacaoRegistro.valor.Equals("9")) ||
					(!linhaTrailler.identificacaoRetorno.valor.Equals("2")) ||
					(!linhaTrailler.codigoBanco.valor.Equals("237")))
				{
					avisoErro("Arquivo de retorno com trailler inválido!!");
					return;
				}
				#endregion

				#region [ Preenche grid ]
				if (linhasArqRetorno.Length > 2) grdBoletos.Rows.Add(linhasArqRetorno.Length - 2);
				for (int i = 1; i < (linhasArqRetorno.Length - 1); i++)
				{
					linhaRegistro.CarregaDados(linhasArqRetorno[i]);
					grdBoletos.Rows[intLinhaGrid].Cells["numeroDocumento"].Value = linhaRegistro.numeroDocumento.valor;
					grdBoletos.Rows[intLinhaGrid].Cells["dataVenctoTitulo"].Value = Global.formataDataCampoArquivoDdMmYyParaDDMMYYYYComSeparador(linhaRegistro.dataVenctoTitulo.valor);
					grdBoletos.Rows[intLinhaGrid].Cells["valorTitulo"].Value = Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor));
					grdBoletos.Rows[intLinhaGrid].Cells["identificacaoOcorrencia"].Value = linhaRegistro.identificacaoOcorrencia.valor + " - " + Global.decodificaIdentificacaoOcorrencia(linhaRegistro.identificacaoOcorrencia.valor);

					#region [ Motivos da ocorrência ]
					strDescricaoMotivoOcorrencia = "";
					if (linhaRegistro.identificacaoOcorrencia.valor.Equals("19"))
					{
						strDescricaoMotivoOcorrencia = linhaRegistro.motivoCodigoOcorrencia19 + " - " + Global.decodificaMotivoOcorrencia19(linhaRegistro.motivoCodigoOcorrencia19.valor);
					}
					else
					{
						listaMotivoOcorrencia = Global.decodificaMotivoOcorrencia(linhaRegistro.identificacaoOcorrencia.valor, linhaRegistro.motivosRejeicoes.valor);
						for (int j = 0; j < listaMotivoOcorrencia.Count; j++)
						{
							if (listaMotivoOcorrencia[j].descricaoMotivoOcorrencia.Length > 0)
							{
								if (strDescricaoMotivoOcorrencia.Length > 0) strDescricaoMotivoOcorrencia += "\n";
								strDescricaoMotivoOcorrencia += listaMotivoOcorrencia[j].motivoOcorrencia + " - " + listaMotivoOcorrencia[j].descricaoMotivoOcorrencia;
							}
						}
					}
					grdBoletos.Rows[intLinhaGrid].Cells["motivosRejeicoes"].Value = strDescricaoMotivoOcorrencia;

					strIdBoletoItem = "";
					if (linhaRegistro.numControleParticipante.valor.IndexOf('=') != -1)
					{
						v = linhaRegistro.numControleParticipante.valor.Split('=');
						strIdBoletoItem = v[1];
					}
					grdBoletos.Rows[intLinhaGrid].Cells["id_boleto_item"].Value = strIdBoletoItem;
					#endregion

					intLinhaGrid++;
				}
				#endregion

				lblTotalRegistros.Text = Global.formataInteiro(intLinhaGrid);

				Global.Usuario.Defaults.FBoletoArqRetorno.pathBoletoArquivoRetorno = Path.GetDirectoryName(strNomeArquivoCompleto);
			}
			finally
			{
				info(ModoExibicaoMensagemRodape.Normal);
			}
		}
		#endregion

		#region [ trataBotaoSelecionaArqRetorno ]
		private void trataBotaoSelecionaArqRetorno()
		{
			#region [ Declarações ]
			DialogResult dr;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
			#endregion

			try
			{
				openFileDialog.InitialDirectory = pathBoletoArquivoRetornoValorDefault();
				dr = openFileDialog.ShowDialog();
				if (dr != DialogResult.OK) return;

				#region [ É o mesmo arquivo já selecionado? ]
				if ((openFileDialog.FileName.Length > 0) && (txtArqRetorno.Text.Length > 0))
				{
					if (openFileDialog.FileName.ToUpper().Equals(txtArqRetorno.Text.ToUpper())) return;
				}
				#endregion

				#region [ Limpa campos ]
				limpaCampos();
				#endregion

				txtArqRetorno.Text = openFileDialog.FileName;

				#region [ Carrega dados do arquivo em array ]
				info(ModoExibicaoMensagemRodape.EmExecucao, "lendo dados do arquivo");
				_linhasArqRetorno = File.ReadAllLines(txtArqRetorno.Text, encode);
				#endregion

				#region [ Consistência ]
				if (_linhasArqRetorno == null)
				{
					avisoErro("É necessário selecionar um arquivo de retorno!!");
					return;
				}
				if (_linhasArqRetorno.Length <= 2)
				{
					avisoErro("Arquivo de retorno selecionado não possui dados!!");
					return;
				}

				if (!_linhasArqRetorno[0].Substring(11, 8).ToUpper().Equals("COBRANCA"))
				{
					avisoErro("O arquivo selecionado não é de cobrança de boletos!!");
					return;
				}
				if (!_linhasArqRetorno[0].Substring(2, 7).ToUpper().Equals("RETORNO"))
				{
					avisoErro("O arquivo selecionado não é um arquivo de retorno!!");
					return;
				}

				_numBancoArqRetorno = Global.digitos(_linhasArqRetorno[0].Substring(76, 3));
				if (_numBancoArqRetorno.Length < 3)
				{
					avisoErro("O arquivo selecionado é inválido: não foi possível identificar o número do banco!!");
					return;
				}

				if ((!_numBancoArqRetorno.Equals(Global.Cte.FIN.NumeroBanco.BRADESCO)) && (!_numBancoArqRetorno.Equals(Global.Cte.FIN.NumeroBanco.SAFRA)))
				{
					avisoErro("O arquivo selecionado é de um banco inválido (" + _numBancoArqRetorno + ")!!");
					return;
				}
				#endregion

				#region [ Carrega header/trailler ]
				if (_numBancoArqRetorno.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
				{
					#region [ Safra ]
					_b422LinhaHeader = new B422HeaderArqRetorno();
					_b422LinhaTrailler = new B422TraillerArqRetorno();
					_b422LinhaHeader.CarregaDados(_linhasArqRetorno[0]);
					_b422LinhaTrailler.CarregaDados(_linhasArqRetorno[_linhasArqRetorno.Length - 1]);
					_boletoCedente = BoletoCedenteDAO.getBoletoCedenteByCodigoEmpresa(_b422LinhaHeader.codigoEmpresa.valor);
					#endregion
				}
				else
				{
					#region [ Bradesco ]
					_b237LinhaHeader = new B237HeaderArqRetorno();
					_b237LinhaTrailler = new B237TraillerArqRetorno();
					_b237LinhaHeader.CarregaDados(_linhasArqRetorno[0]);
					_b237LinhaTrailler.CarregaDados(_linhasArqRetorno[_linhasArqRetorno.Length - 1]);
					_boletoCedente = BoletoCedenteDAO.getBoletoCedenteByCodigoEmpresa(_b237LinhaHeader.codigoEmpresa.valor);
					#endregion
				}
				#endregion

				preencheCamposInfArqRetorno();
				carregaGridBoletos();
				grdBoletos.Focus();
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

		#region [ trataBotaoCarregaArqRetorno ]
		private void trataBotaoCarregaArqRetorno()
		{
			#region [ Declarações ]
			int intNsuBoletoArqRetorno = 0;
			int intTotalRegistrosArqRetorno;
			int intDuracaoProcessamentoEmSeg = 0;
			bool blnGerouNsu;
			bool blnSucesso;
			bool blnSerasaGravouDados;
			bool blnSerasaOcorrenciaIgnorada;
			string numBanco;
			string codEmpresa = "";
			string sDataGravacaoArquivo = "";
			String strAux;
			String strMsg;
			String strMsgProgresso;
			String strMsgErro = "";
			String strMsgErroAux = "";
			String strMsgErroLog = "";
			String strDescricaoLog = "";
			String strNomeArquivoCompleto;
			String strNomeArquivoCompletoRenomeado;
			String strNomeArquivoCompletoRenomeadoAux;
			String usuarioProcessamentoAnterior;
			String usuarioProcessamentoUltArqCarregadoComSucesso;
			String strNomeUltArqRetornoCarregadoComSucesso;
			String strMotivoNaoGravarDadosSerasa;
			String[] linhasArqRetorno;
			BoletoArqRetorno boletoArqRetorno = new BoletoArqRetorno();
			DateTime dtInicioProcessamento;
			DateTime dtGravacaoArquivo;
			DateTime dtGravacaoUltArqCarregadoComSucesso;
			DateTime dtHrProcessamentoAnterior;
			DateTime dtHrProcessamentoUltArqCarregadoComSucesso;
			FinLog finLog = new FinLog();
			B237HeaderArqRetorno b237LinhaHeader = null;
			B237TraillerArqRetorno b237LinhaTrailler = null;
			B237RegTipo1ArqRetorno b237LinhaRegistro = null;
			B422HeaderArqRetorno b422LinhaHeader = null;
			B422TraillerArqRetorno b422LinhaTrailler = null;
			B422RegTipo1ArqRetorno b422LinhaRegistro = null;
			FAutorizacao fAutorizacao;
			DialogResult drAutorizacao;
			Encoding encode = Encoding.GetEncoding("Windows-1252");
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

			#region [ Obtém nome do arquivo de retorno ]
			strNomeArquivoCompleto = txtArqRetorno.Text;
			#endregion

			#region [ Consistência ]
			if (strNomeArquivoCompleto.Length == 0)
			{
				avisoErro("É necessário selecionar o arquivo de retorno a ser carregado!!");
				return;
			}
			if (!File.Exists(strNomeArquivoCompleto))
			{
				avisoErro("O arquivo de retorno selecionado não existe!!\n\n" + strNomeArquivoCompleto);
				return;
			}
			#endregion

			#region [ Confirmação ]
			if (!confirma("Confirma a carga do arquivo de retorno?")) return;
			#endregion

			info(ModoExibicaoMensagemRodape.EmExecucao, "lendo o arquivo de retorno");
			try
			{
				dtInicioProcessamento = DateTime.Now;

				#region [ Carrega dados do arquivo em array ]
				adicionaDisplay("Leitura dos registros do arquivo de retorno " + Path.GetFileName(strNomeArquivoCompleto));
				linhasArqRetorno = File.ReadAllLines(strNomeArquivoCompleto, encode);
				adicionaDisplay("Registros para processar: " + Global.formataInteiro(linhasArqRetorno.Length - 2));
				#endregion

				#region [ Consistência do arquivo ]
				if (!linhasArqRetorno[0].Substring(11, 8).ToUpper().Equals("COBRANCA"))
				{
					strMsgErro = "O arquivo selecionado não é de cobrança de boletos!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}

				if (!linhasArqRetorno[0].Substring(2, 7).ToUpper().Equals("RETORNO"))
				{
					strMsgErro = "O arquivo selecionado não é um arquivo de retorno!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}

				numBanco = Global.digitos(linhasArqRetorno[0].Substring(76, 3));
				if (numBanco.Length < 3)
				{
					strMsgErro = "O arquivo selecionado é inválido: não foi possível identificar o número do banco!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}

				if ((!numBanco.Equals(Global.Cte.FIN.NumeroBanco.BRADESCO)) && (!numBanco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA)))
				{
					strMsgErro = "O arquivo selecionado é de um banco inválido(" + numBanco + ")!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}
				#endregion

				#region [ Inicializa variáveis conforme layout do banco ]
				if (numBanco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
				{
					#region [ Safra ]
					b422LinhaHeader = new B422HeaderArqRetorno();
					b422LinhaTrailler = new B422TraillerArqRetorno();
					b422LinhaRegistro = new B422RegTipo1ArqRetorno();
					b422LinhaHeader.CarregaDados(linhasArqRetorno[0]);
					b422LinhaTrailler.CarregaDados(linhasArqRetorno[linhasArqRetorno.Length - 1]);
					#endregion
				}
				else
				{
					#region [ Bradesco ]
					b237LinhaHeader = new B237HeaderArqRetorno();
					b237LinhaTrailler = new B237TraillerArqRetorno();
					b237LinhaRegistro = new B237RegTipo1ArqRetorno();
					b237LinhaHeader.CarregaDados(linhasArqRetorno[0]);
					b237LinhaTrailler.CarregaDados(linhasArqRetorno[linhasArqRetorno.Length - 1]);
					#endregion
				}
				#endregion

				#region [ Consistência do header e trailler ]
				if (numBanco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
				{
					#region [ Safra ]
					if ((!b422LinhaHeader.identificacaoRegistro.valor.Equals("0")) ||
						(!b422LinhaHeader.identificacaoArquivoRetorno.valor.Equals("2")) ||
						(!b422LinhaHeader.literalRetorno.valor.ToUpper().Equals("RETORNO")))
					{
						strMsgErro = "Arquivo de retorno com header inválido!!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}

					if ((!b422LinhaTrailler.identificacaoRegistro.valor.Equals("9")) ||
						(!b422LinhaTrailler.identificacaoRetorno.valor.Equals("2")) ||
						(!b422LinhaTrailler.codigoBanco.valor.Equals(Global.Cte.FIN.NumeroBanco.SAFRA)))
					{
						strMsgErro = "Arquivo de retorno com trailler inválido!!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion
				}
				else
				{
					#region [ Bradesco ]
					if ((!b237LinhaHeader.identificacaoRegistro.valor.Equals("0")) ||
						(!b237LinhaHeader.identificacaoArquivoRetorno.valor.Equals("2")) ||
						(!b237LinhaHeader.literalRetorno.valor.ToUpper().Equals("RETORNO")))
					{
						strMsgErro = "Arquivo de retorno com header inválido!!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}

					if ((!b237LinhaTrailler.identificacaoRegistro.valor.Equals("9")) ||
						(!b237LinhaTrailler.identificacaoRetorno.valor.Equals("2")) ||
						(!b237LinhaTrailler.codigoBanco.valor.Equals(Global.Cte.FIN.NumeroBanco.BRADESCO)))
					{
						strMsgErro = "Arquivo de retorno com trailler inválido!!";
						adicionaErro(strMsgErro);
						avisoErro(strMsgErro);
						return;
					}
					#endregion
				}
				#endregion

				#region [ Obtém dados do cedente ]
				if (numBanco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
				{
					codEmpresa = b422LinhaHeader.codigoEmpresa.valor;
					sDataGravacaoArquivo = b422LinhaHeader.dataGravacaoArquivo.valor;
				}
				else
				{
					codEmpresa = b237LinhaHeader.codigoEmpresa.valor;
					sDataGravacaoArquivo = b237LinhaHeader.dataGravacaoArquivo.valor;
				}

				_boletoCedente = BoletoCedenteDAO.getBoletoCedenteByCodigoEmpresa(codEmpresa);

				if (_boletoCedente == null)
				{
					strMsgErro = "Não foi localizado o registro da empresa cedente: " + codEmpresa + "!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}

				if (_boletoCedente.id == 0)
				{
					strMsgErro = "Falha ao recuperar o registro da empresa cedente: " + codEmpresa + "!!";
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
					return;
				}
				#endregion

				#region [ Verifica se este arquivo já foi carregado anteriormente ]
				if (BoletoDAO.boletoArqRetornoJaCarregado(Path.GetFileName(strNomeArquivoCompleto), codEmpresa, sDataGravacaoArquivo, out dtHrProcessamentoAnterior, out usuarioProcessamentoAnterior))
				{
					strMsg = "O arquivo de retorno " + Path.GetFileName(strNomeArquivoCompleto) + " já foi carregado em " + Global.formataDataDdMmYyyyHhMmSsComSeparador(dtHrProcessamentoAnterior) + " por " + usuarioProcessamentoAnterior + "!!" +
							 "\nNão é possível carregá-lo novamente!!";
					adicionaErro(strMsg);
					avisoErro(strMsg);
					return;
				}
				#endregion

				#region [ Consiste data de geração do arquivo ]
				dtGravacaoArquivo = Global.converteDdMmYyParaDateTime(sDataGravacaoArquivo);

				#region [ Verifica se os arquivos de retorno estão sendo carregados fora de ordem ]
				if (BoletoDAO.boletoArqRetornoObtemDtGravacaoUltArqCarregadoComSucesso(codEmpresa, out dtGravacaoUltArqCarregadoComSucesso, out strNomeUltArqRetornoCarregadoComSucesso, out dtHrProcessamentoUltArqCarregadoComSucesso, out usuarioProcessamentoUltArqCarregadoComSucesso))
				{
					if (dtGravacaoUltArqCarregadoComSucesso > dtGravacaoArquivo)
					{
						strAux = "O arquivo de retorno " + Path.GetFileName(strNomeArquivoCompleto) +
								 " foi gravado em " + Global.formataDataDdMmYyyyComSeparador(dtGravacaoArquivo) +
								 ", mas o último arquivo de retorno carregado foi gravado em " +
								 Global.formataDataDdMmYyyyComSeparador(dtGravacaoUltArqCarregadoComSucesso) +
								 " (" + strNomeUltArqRetornoCarregadoComSucesso + " carregado por " + usuarioProcessamentoUltArqCarregadoComSucesso + " em " + Global.formataDataDdMmYyyyHhMmComSeparador(dtHrProcessamentoUltArqCarregadoComSucesso) + ")!!" +
								 "\n" +
								 "Isso pode ser um indício de que os arquivos de retorno estão sendo carregados fora de ordem!!" +
								 "\n" +
								 "Digite a senha para confirmar que autoriza a carga deste arquivo de retorno!!";
						fAutorizacao = new FAutorizacao(strAux);
						drAutorizacao = fAutorizacao.ShowDialog();
						if (drAutorizacao != DialogResult.OK)
						{
							strMsg = "Operação não confirmada!!\nA carga do arquivo de retorno não será realizada!!";
							adicionaErro(strMsg);
							avisoErro(strMsg);
							return;
						}
						if (fAutorizacao.senha.ToUpper() != Global.Usuario.senhaDescriptografada.ToUpper())
						{
							strMsg = "Senha inválida!!\nA carga do arquivo de retorno não será realizada!!";
							adicionaErro(strMsg);
							avisoErro(strMsg);
							return;
						}
						this.Refresh();
					}
				}
				#endregion

				if (dtGravacaoArquivo < DateTime.Now.AddDays(-7))
				{
					strAux = "O arquivo de retorno " + Path.GetFileName(strNomeArquivoCompleto) + " foi gravado em " + Global.formataDataDdMmYyyyComSeparador(dtGravacaoArquivo) + "!!" +
							 "\n" +
							 "Digite a senha para confirmar a carga deste arquivo de retorno que parece ser ANTIGO!!";
					fAutorizacao = new FAutorizacao(strAux);
					drAutorizacao = fAutorizacao.ShowDialog();
					if (drAutorizacao != DialogResult.OK)
					{
						strMsg = "Operação não confirmada!!\nA carga do arquivo de retorno não será realizada!!";
						adicionaErro(strMsg);
						avisoErro(strMsg);
						return;
					}
					if (fAutorizacao.senha.ToUpper() != Global.Usuario.senhaDescriptografada.ToUpper())
					{
						strMsg = "Senha inválida!!\nA carga do arquivo de retorno não será realizada!!";
						adicionaErro(strMsg);
						avisoErro(strMsg);
						return;
					}
					this.Refresh();
				}
				#endregion

				blnSucesso = false;
				try
				{
					BD.iniciaTransacao();

					info(ModoExibicaoMensagemRodape.EmExecucao, "processando o arquivo de retorno");

					#region [ Gera o NSU para o novo registro que será gravado em t_FIN_BOLETO_ARQ_RETORNO ]
					blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_BOLETO_ARQ_RETORNO, ref intNsuBoletoArqRetorno, ref strMsgErro);
					if (!blnGerouNsu)
					{
						throw new FinanceiroException("Falha ao tentar gerar o NSU para o registro de histórico de arquivos de retorno!!\n" + strMsgErro);
					}
					#endregion

					if (numBanco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
					{
						#region [  Safra ]

						#region [ Grava o registro em t_FIN_BOLETO_ARQ_RETORNO ]
						boletoArqRetorno.id = intNsuBoletoArqRetorno;
						boletoArqRetorno.id_boleto_cedente = _boletoCedente.id;
						boletoArqRetorno.st_processamento = Global.Cte.FIN.CodBoletoArqRetornoStProcessamento.EM_PROCESSAMENTO;
						boletoArqRetorno.qtde_registros = linhasArqRetorno.Length - 2;
						boletoArqRetorno.codigo_empresa = b422LinhaHeader.codigoEmpresa.valor;
						boletoArqRetorno.nome_empresa = b422LinhaHeader.nomeEmpresa.valor;
						boletoArqRetorno.num_banco = b422LinhaHeader.numBanco.valor;
						boletoArqRetorno.nome_banco = b422LinhaHeader.nomeBanco.valor;
						boletoArqRetorno.data_gravacao_arquivo = b422LinhaHeader.dataGravacaoArquivo.valor;
						boletoArqRetorno.numero_aviso_bancario = b422LinhaTrailler.numAvisoBancarioSimples.valor;
						boletoArqRetorno.data_credito = "";
						boletoArqRetorno.qtdeTitulosEmCobranca = b422LinhaTrailler.qtdeTitulosEmCobrancaSimples.valor;
						boletoArqRetorno.valorTotalEmCobranca = b422LinhaTrailler.valorTotalEmCobrancaSimples.valor;
						boletoArqRetorno.nome_arq_retorno = Path.GetFileName(strNomeArquivoCompleto);
						boletoArqRetorno.caminho_arq_retorno = Path.GetDirectoryName(strNomeArquivoCompleto);
						if (!BoletoDAO.boletoArqRetornoInsere(Global.Usuario.usuario, boletoArqRetorno, ref strMsgErroAux))
						{
							throw new FinanceiroException("Falha ao gravar o histórico de arquivos de retorno no banco de dados!!" + "\n" + strMsgErroAux);
						}
						#endregion

						#endregion
					}
					else
					{
						#region [ Bradesco ]

						#region [ Grava o registro em t_FIN_BOLETO_ARQ_RETORNO ]
						boletoArqRetorno.id = intNsuBoletoArqRetorno;
						boletoArqRetorno.id_boleto_cedente = _boletoCedente.id;
						boletoArqRetorno.st_processamento = Global.Cte.FIN.CodBoletoArqRetornoStProcessamento.EM_PROCESSAMENTO;
						boletoArqRetorno.qtde_registros = linhasArqRetorno.Length - 2;
						boletoArqRetorno.codigo_empresa = b237LinhaHeader.codigoEmpresa.valor;
						boletoArqRetorno.nome_empresa = b237LinhaHeader.nomeEmpresa.valor;
						boletoArqRetorno.num_banco = b237LinhaHeader.numBanco.valor;
						boletoArqRetorno.nome_banco = b237LinhaHeader.nomeBanco.valor;
						boletoArqRetorno.data_gravacao_arquivo = b237LinhaHeader.dataGravacaoArquivo.valor;
						boletoArqRetorno.numero_aviso_bancario = b237LinhaHeader.numAvisoBancario.valor;
						boletoArqRetorno.data_credito = b237LinhaHeader.dataCredito.valor;
						boletoArqRetorno.qtdeTitulosEmCobranca = b237LinhaTrailler.qtdeTitulosEmCobranca.valor;
						boletoArqRetorno.valorTotalEmCobranca = b237LinhaTrailler.valorTotalEmCobranca.valor;
						boletoArqRetorno.qtdeRegsOcorrencia02ConfirmacaoEntradas = b237LinhaTrailler.qtdeRegsOcorrencia02ConfirmacaoEntradas.valor;
						boletoArqRetorno.valorRegsOcorrencia02ConfirmacaoEntradas = b237LinhaTrailler.valorRegsOcorrencia02ConfirmacaoEntradas.valor;
						boletoArqRetorno.valorRegsOcorrencia06Liquidacao = b237LinhaTrailler.valorRegsOcorrencia06Liquidacao.valor;
						boletoArqRetorno.qtdeRegsOcorrencia06Liquidacao = b237LinhaTrailler.qtdeRegsOcorrencia06Liquidacao.valor;
						boletoArqRetorno.valorRegsOcorrencia06 = b237LinhaTrailler.valorRegsOcorrencia06.valor;
						boletoArqRetorno.qtdeRegsOcorrencia09e10TitulosBaixados = b237LinhaTrailler.qtdeRegsOcorrencia09e10TitulosBaixados.valor;
						boletoArqRetorno.valorRegsOcorrencia09e10TitulosBaixados = b237LinhaTrailler.valorRegsOcorrencia09e10TitulosBaixados.valor;
						boletoArqRetorno.qtdeRegsOcorrencia13AbatimentoCancelado = b237LinhaTrailler.qtdeRegsOcorrencia13AbatimentoCancelado.valor;
						boletoArqRetorno.valorRegsOcorrencia13AbatimentoCancelado = b237LinhaTrailler.valorRegsOcorrencia13AbatimentoCancelado.valor;
						boletoArqRetorno.qtdeRegsOcorrencia14VenctoAlterado = b237LinhaTrailler.qtdeRegsOcorrencia14VenctoAlterado.valor;
						boletoArqRetorno.valorRegsOcorrencia14VenctoAlterado = b237LinhaTrailler.valorRegsOcorrencia14VenctoAlterado.valor;
						boletoArqRetorno.qtdeRegsOcorrencia12AbatimentoConcedido = b237LinhaTrailler.qtdeRegsOcorrencia12AbatimentoConcedido.valor;
						boletoArqRetorno.valorRegsOcorrencia12AbatimentoConcedido = b237LinhaTrailler.valorRegsOcorrencia12AbatimentoConcedido.valor;
						boletoArqRetorno.qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto = b237LinhaTrailler.qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto.valor;
						boletoArqRetorno.valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto = b237LinhaTrailler.valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto.valor;
						boletoArqRetorno.valorTotalRateiosEfetuados = b237LinhaTrailler.valorTotalRateiosEfetuados.valor;
						boletoArqRetorno.qtdeTotalRateiosEfetuados = b237LinhaTrailler.qtdeTotalRateiosEfetuados.valor;
						boletoArqRetorno.nome_arq_retorno = Path.GetFileName(strNomeArquivoCompleto);
						boletoArqRetorno.caminho_arq_retorno = Path.GetDirectoryName(strNomeArquivoCompleto);
						if (!BoletoDAO.boletoArqRetornoInsere(Global.Usuario.usuario, boletoArqRetorno, ref strMsgErroAux))
						{
							throw new FinanceiroException("Falha ao gravar o histórico de arquivos de retorno no banco de dados!!" + "\n" + strMsgErroAux);
						}
						#endregion

						#endregion
					}

					#region [ Grava o log no BD ]
					strDescricaoLog = "Início da carga do arquivo de retorno: " + strNomeArquivoCompleto;
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_CARREGA_ARQ_RETORNO;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO;
					finLog.id_boleto_cedente = _boletoCedente.id;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
					#endregion

					#region [ Mensagem no log se a empresa participa ou não do Reciprocidade Serasa ]
					if (_boletoCedente.st_participante_serasa_reciprocidade != 0)
					{
						Global.gravaLogAtividade("Serasa Reciprocidade: cedente (id=" + _boletoCedente.id.ToString() + ") participa do programa de Reciprocidade");
					}
					else
					{
						Global.gravaLogAtividade("Serasa Reciprocidade: cedente (id=" + _boletoCedente.id.ToString() + ") NÃO participa do programa de Reciprocidade");
					}
					#endregion

					#region [ Trata cada um dos registros ]
					intTotalRegistrosArqRetorno = linhasArqRetorno.Length - 2;
					for (int intLinha = 1; intLinha < (linhasArqRetorno.Length - 1); intLinha++)
					{
						if (numBanco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
						{
							#region [ Safra ]
							b422LinhaRegistro.CarregaDados(linhasArqRetorno[intLinha]);
							strMsgProgresso = "Processando registro " + intLinha.ToString() + ": " +
											  b422LinhaRegistro.identificacaoOcorrencia.valor + " - " +
											  Global.b422DecodificaIdentificacaoOcorrencia(b422LinhaRegistro.identificacaoOcorrencia.valor) +
											  "   (" + (100 * intLinha / intTotalRegistrosArqRetorno).ToString() + "%)";
							#endregion
						}
						else
						{
							#region [ Bradesco ]
							b237LinhaRegistro.CarregaDados(linhasArqRetorno[intLinha]);
							strMsgProgresso = "Processando registro " + intLinha.ToString() + ": " +
											  b237LinhaRegistro.identificacaoOcorrencia.valor + " - " +
											  Global.b237DecodificaIdentificacaoOcorrencia(b237LinhaRegistro.identificacaoOcorrencia.valor) +
											  "   (" + (100 * intLinha / intTotalRegistrosArqRetorno).ToString() + "%)";
							#endregion
						}

						adicionaDisplay(strMsgProgresso);

						#region [ DoEvents ]
						Application.DoEvents();
						#endregion

						#region [ Tratamento para cada tipo de ocorrência ]
						if (numBanco.Equals(Global.Cte.FIN.NumeroBanco.SAFRA))
						{
							#region [ Safra ]
							if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("02"))
							{
								#region [ Tratamento p/ ocorrência 02: entrada confirmada ]
								if (!b422TrataOcorrencia02EntradaConfirmada(b422LinhaHeader, b422LinhaRegistro, ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 02 (entrada confirmada)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("03"))
							{
								#region [ Tratamento p/ ocorrência 03: entrada rejeitada ]
								if (!b422TrataOcorrenciaValaComum(intNsuBoletoArqRetorno, b422LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 03 (entrada rejeitada)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nosso número: " + Global.formataBoletoNossoNumero(b422LinhaRegistro.nossoNumeroSemDigito.valor, b422LinhaRegistro.digitoNossoNumero.valor) + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("06"))
							{
								#region [ Tratamento p/ ocorrência 06: liquidação normal ]
								if (!b422TrataOcorrencia06LiquidacaoNormal(intNsuBoletoArqRetorno, b422LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 06 (liquidação normal)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("09"))
							{
								#region [ Tratamento p/ ocorrência 09: baixado automaticamente ]
								// 21/09/2009: Rogério definiu que as baixas devem apenas ser exibidas na listagem 
								// =========== de ocorrências para serem tratadas manualmente. Ou seja, a localização
								// e cancelamento do lançamento de fluxo de caixa associado ao boleto será feito
								// de forma manual. Isso permite que o tratamento fique coerente enquanto não
								// se implementa a operação de identificação de depósito desconhecido. Além disso,
								// isso obriga que se tome ciência de tudo o que o banco está baixando.
								if (!b422TrataOcorrencia09BaixadoAutomaticamente(intNsuBoletoArqRetorno, b422LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 09 (baixado automaticamente)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("10"))
							{
								#region [ Tratamento p/ ocorrência 10: baixado conforme instruções ]
								// 21/09/2009: Rogério definiu que as baixas devem apenas ser exibidas na listagem 
								// =========== de ocorrências para serem tratadas manualmente. Ou seja, a localização
								// e cancelamento do lançamento de fluxo de caixa associado ao boleto será feito
								// de forma manual. Isso permite que o tratamento fique coerente enquanto não
								// se implementa a operação de identificação de depósito desconhecido. Além disso,
								// isso obriga que se tome ciência de tudo o que o banco está baixando.
								if (!b422TrataOcorrencia10BaixadoConfInstrucoes(intNsuBoletoArqRetorno, b422LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 10 (baixado conforme instruções)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("12"))
							{
								#region [ Tratamento p/ ocorrência 12: abatimento concedido ]
								if (!b422TrataOcorrencia12AbatimentoConcedido(intNsuBoletoArqRetorno, b422LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 12 (abatimento concedido)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("13"))
							{
								#region [ Tratamento p/ ocorrência 13: abatimento cancelado ]
								if (!b422TrataOcorrencia13AbatimentoCancelado(intNsuBoletoArqRetorno, b422LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 13 (abatimento cancelado)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("14"))
							{
								#region [ Tratamento p/ ocorrência 14: vencimento alterado ]
								if (!b422TrataOcorrencia14VenctoAlterado(intNsuBoletoArqRetorno, b422LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 14 (vencimento alterado)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("15"))
							{
								#region [ Tratamento p/ ocorrência 15: liquidação em cartório ]
								if (!b422TrataOcorrencia15LiquidacaoEmCartorio(intNsuBoletoArqRetorno, b422LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 15 (liquidação em cartório)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b422LinhaRegistro.identificacaoOcorrencia.valor.Equals("16"))
							{
								#region [ Tratamento p/ ocorrência 16: título pago em cheque ]
								if (!b422TrataOcorrencia16TituloPagoEmCheque(intNsuBoletoArqRetorno, b422LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 16 (título pago em cheque)!!\nNº documento: " + b422LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b422LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (linhaRegistro.identificacaoOcorrencia.valor.Equals("17"))
							{
								#region [ Tratamento p/ ocorrência 17: liquidação após baixa ou título não registrado ]
								if (!trataOcorrencia17LiqAposBaixaOuTitNaoRegistrado(intNsuBoletoArqRetorno, linhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 17 (liquidação após baixa ou título não registrado)!!\nNº documento: " + linhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + linhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (linhaRegistro.identificacaoOcorrencia.valor.Equals("19"))
							{
								#region [ Tratamento p/ ocorrência 19: confirmação recebimento instrução de protesto ]
								if (!trataOcorrencia19ConfirmacaoRecebInstProtesto(intNsuBoletoArqRetorno, linhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 19 (confirmação receb. inst. de protesto)!!\nNº documento: " + linhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + linhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (linhaRegistro.identificacaoOcorrencia.valor.Equals("22"))
							{
								#region [ Tratamento p/ ocorrência 22: título com pagamento cancelado ]
								if (!trataOcorrencia22TituloComPagamentoCancelado(intNsuBoletoArqRetorno, linhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 22 (título com pagamento cancelado)!!\nNº documento: " + linhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + linhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (linhaRegistro.identificacaoOcorrencia.valor.Equals("23"))
							{
								#region [ Tratamento p/ ocorrência 23: entrada do título em cartório ]
								if (!trataOcorrencia23EntradaTituloEmCartorio(intNsuBoletoArqRetorno, linhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 23 (entrada do título em cartório)!!\nNº documento: " + linhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + linhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (linhaRegistro.identificacaoOcorrencia.valor.Equals("24"))
							{
								#region [ Tratamento p/ ocorrência 24: entrada rejeitada por CEP irregular ]
								if (!trataOcorrencia24EntradaRejeitadaCepIrregular(intNsuBoletoArqRetorno, linhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 24 (entrada rejeitada por CEP irregular)!!\nNº documento: " + linhaRegistro.numeroDocumento.valor + ", nosso número: " + Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor) + ", nº ctrl: " + linhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (linhaRegistro.identificacaoOcorrencia.valor.Equals("28"))
							{
								#region [ Tratamento p/ ocorrência 28: débito de tarifas/custas ]
								if (!trataOcorrencia28DebitoTarifasCustas(intNsuBoletoArqRetorno, linhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 28 (débito de tarifas/custas)!!\nNº documento: " + linhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + linhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (linhaRegistro.identificacaoOcorrencia.valor.Equals("34"))
							{
								#region [ Tratamento p/ ocorrência 34: retirado de cartório e manutenção carteira ]
								if (!trataOcorrencia34RetiradoCartorioManutencaoCarteira(intNsuBoletoArqRetorno, linhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 34 (retirado de cartório e manutenção carteira)!!\nNº documento: " + linhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + linhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else
							{
								#region [ Tratamento p/ demais casos ]
								if (!trataOcorrenciaValaComum(intNsuBoletoArqRetorno, linhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha no tratamento de vala comum para a ocorrência " + linhaRegistro.identificacaoOcorrencia.valor + "!!\nNosso número: " + Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor) + "\n\n" + strMsgErro);
								}
								#endregion
							}
							#endregion
						}
						else
						{
							#region [ Bradesco ]
							if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("02"))
							{
								#region [ Tratamento p/ ocorrência 02: entrada confirmada ]
								if (!b237TrataOcorrencia02EntradaConfirmada(b237LinhaHeader, b237LinhaRegistro, ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 02 (entrada confirmada)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("03"))
							{
								#region [ Tratamento p/ ocorrência 03: entrada rejeitada ]
								if (!b237TrataOcorrenciaValaComum(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 03 (entrada rejeitada)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nosso número: " + Global.formataBoletoNossoNumero(b237LinhaRegistro.nossoNumeroSemDigito.valor, b237LinhaRegistro.digitoNossoNumero.valor) + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("06"))
							{
								#region [ Tratamento p/ ocorrência 06: liquidação normal ]
								if (!b237TrataOcorrencia06LiquidacaoNormal(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 06 (liquidação normal)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("09"))
							{
								#region [ Tratamento p/ ocorrência 09: baixado automat. via arquivo ]
								// 21/09/2009: Rogério definiu que as baixas devem apenas ser exibidas na listagem 
								// =========== de ocorrências para serem tratadas manualmente. Ou seja, a localização
								// e cancelamento do lançamento de fluxo de caixa associado ao boleto será feito
								// de forma manual. Isso permite que o tratamento fique coerente enquanto não
								// se implementa a operação de identificação de depósito desconhecido. Além disso,
								// isso obriga que se tome ciência de tudo o que o banco está baixando.
								if (!b237TrataOcorrencia09BaixadoAutoViaArq(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 09 (baixado automaticamente via arquivo)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("10"))
							{
								#region [ Tratamento p/ ocorrência 10: baixado conforme instruções da agência ]
								// 21/09/2009: Rogério definiu que as baixas devem apenas ser exibidas na listagem 
								// =========== de ocorrências para serem tratadas manualmente. Ou seja, a localização
								// e cancelamento do lançamento de fluxo de caixa associado ao boleto será feito
								// de forma manual. Isso permite que o tratamento fique coerente enquanto não
								// se implementa a operação de identificação de depósito desconhecido. Além disso,
								// isso obriga que se tome ciência de tudo o que o banco está baixando.
								if (!b237TrataOcorrencia10BaixadoConfInstrAgencia(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 10 (baixado conforme instruções da agência)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("12"))
							{
								#region [ Tratamento p/ ocorrência 12: abatimento concedido ]
								if (!b237TrataOcorrencia12AbatimentoConcedido(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 12 (abatimento concedido)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("13"))
							{
								#region [ Tratamento p/ ocorrência 13: abatimento cancelado ]
								if (!b237TrataOcorrencia13AbatimentoCancelado(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 13 (abatimento cancelado)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("14"))
							{
								#region [ Tratamento p/ ocorrência 14: vencimento alterado ]
								if (!b237TrataOcorrencia14VenctoAlterado(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 14 (vencimento alterado)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("15"))
							{
								#region [ Tratamento p/ ocorrência 15: liquidação em cartório ]
								if (!b237TrataOcorrencia15LiquidacaoEmCartorio(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 15 (liquidação em cartório)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("16"))
							{
								#region [ Tratamento p/ ocorrência 16: título pago em cheque ]
								if (!b237TrataOcorrencia16TituloPagoEmCheque(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 16 (título pago em cheque)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("17"))
							{
								#region [ Tratamento p/ ocorrência 17: liquidação após baixa ou título não registrado ]
								if (!b237TrataOcorrencia17LiqAposBaixaOuTitNaoRegistrado(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 17 (liquidação após baixa ou título não registrado)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("19"))
							{
								#region [ Tratamento p/ ocorrência 19: confirmação recebimento instrução de protesto ]
								if (!b237TrataOcorrencia19ConfirmacaoRecebInstProtesto(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 19 (confirmação receb. inst. de protesto)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("22"))
							{
								#region [ Tratamento p/ ocorrência 22: título com pagamento cancelado ]
								if (!b237TrataOcorrencia22TituloComPagamentoCancelado(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 22 (título com pagamento cancelado)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("23"))
							{
								#region [ Tratamento p/ ocorrência 23: entrada do título em cartório ]
								if (!b237TrataOcorrencia23EntradaTituloEmCartorio(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 23 (entrada do título em cartório)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("24"))
							{
								#region [ Tratamento p/ ocorrência 24: entrada rejeitada por CEP irregular ]
								if (!b237TrataOcorrencia24EntradaRejeitadaCepIrregular(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 24 (entrada rejeitada por CEP irregular)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor + ", nosso número: " + Global.formataBoletoNossoNumero(b237LinhaRegistro.nossoNumeroSemDigito.valor, b237LinhaRegistro.digitoNossoNumero.valor) + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("28"))
							{
								#region [ Tratamento p/ ocorrência 28: débito de tarifas/custas ]
								if (!b237TrataOcorrencia28DebitoTarifasCustas(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 28 (débito de tarifas/custas)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else if (b237LinhaRegistro.identificacaoOcorrencia.valor.Equals("34"))
							{
								#region [ Tratamento p/ ocorrência 34: retirado de cartório e manutenção carteira ]
								if (!b237TrataOcorrencia34RetiradoCartorioManutencaoCarteira(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha ao tratar a ocorrência 34 (retirado de cartório e manutenção carteira)!!\nNº documento: " + b237LinhaRegistro.numeroDocumento.valor.Trim() + ", nº ctrl: " + b237LinhaRegistro.numControleParticipante.valor + "\n\n" + strMsgErro);
								}
								#endregion
							}
							else
							{
								#region [ Tratamento p/ demais casos ]
								if (!b237TrataOcorrenciaValaComum(intNsuBoletoArqRetorno, b237LinhaRegistro, linhasArqRetorno[intLinha], ref strMsgErro))
								{
									throw new FinanceiroException("Falha no tratamento de vala comum para a ocorrência " + b237LinhaRegistro.identificacaoOcorrencia.valor + "!!\nNosso número: " + Global.formataBoletoNossoNumero(b237LinhaRegistro.nossoNumeroSemDigito.valor, b237LinhaRegistro.digitoNossoNumero.valor) + "\n\n" + strMsgErro);
								}
								#endregion
							}
							#endregion
						}
						#endregion

						#region [ Processa a inclusão dos dados na tabela de reciprocidade com a Serasa, se for o caso ]
						// Cada empresa participante (mesmo que seja uma filial) do programa de Reciprocidade deve trocar
						// seu próprio conjunto de arquivos c/ a Serasa.
						if (_boletoCedente.st_participante_serasa_reciprocidade != 0)
						{
							if (trataInclusaoDadosSerasaTituloMovimento(intNsuBoletoArqRetorno, _boletoCedente.id, linhaRegistro, out blnSerasaGravouDados, out blnSerasaOcorrenciaIgnorada, out strMotivoNaoGravarDadosSerasa, out strMsgErro))
							{
								if (blnSerasaGravouDados)
								{
									Global.gravaLogAtividade("Serasa Reciprocidade: dados gravados (ocorrência: " + linhaRegistro.identificacaoOcorrencia.valor + ")");
								}
							}
							else
							{
								if (blnSerasaOcorrenciaIgnorada)
								{
									Global.gravaLogAtividade("Serasa Reciprocidade: registro ignorado devido ao tipo de ocorrência (" + linhaRegistro.identificacaoOcorrencia.valor + ")");
								}
								if (strMotivoNaoGravarDadosSerasa.Length > 0)
								{
									Global.gravaLogAtividade("Serasa Reciprocidade: registro ignorado (motivo: " + strMotivoNaoGravarDadosSerasa + ")");
								}
								if (strMsgErro.Length > 0)
								{
									strMsgErro = "Serasa Reciprocidade: falha ao tratar a inclusão dos dados do título na tabela de integração com o sistema de Reciprocidade da Serasa!" + "\n" + strMsgErro;
									Global.gravaLogAtividade(strMsgErro);
								}
							}
						}
						#endregion
					} // for
					#endregion

					#region [ Atualiza os dados no registro em t_FIN_BOLETO_ARQ_RETORNO ]
					intDuracaoProcessamentoEmSeg = Global.calculaTimeSpanSegundos(DateTime.Now - dtInicioProcessamento);
					if (!BoletoDAO.boletoArqRetornoAtualiza(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										Global.Cte.FIN.CodBoletoArqRetornoStProcessamento.SUCESSO,
										intDuracaoProcessamentoEmSeg,
										"",
										ref strMsgErroAux))
					{
						throw new FinanceiroException("Falha ao gravar o histórico de arquivos de retorno no banco de dados!!");
					}
					#endregion

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					Global.gravaLogAtividade(ex.ToString());
					strMsgErro = ex.ToString();
					blnSucesso = false;
				}

				if (blnSucesso)
				{
					BD.commitTransacao();

					adicionaDisplay("Arquivo de retorno carregado com sucesso!!");

					#region [ Renomeia o arquivo de retorno ]
					strNomeArquivoCompletoRenomeado = Path.ChangeExtension(strNomeArquivoCompleto, ".PRC");
					if (File.Exists(strNomeArquivoCompletoRenomeado))
					{
						int intFileIndex = 0;
						while (File.Exists(strNomeArquivoCompletoRenomeado))
						{
							intFileIndex++;
							strNomeArquivoCompletoRenomeadoAux = strNomeArquivoCompletoRenomeado + ".OLD." + intFileIndex.ToString().PadLeft(3, '0');
							if (!File.Exists(strNomeArquivoCompletoRenomeadoAux)) File.Move(strNomeArquivoCompletoRenomeado, strNomeArquivoCompletoRenomeadoAux);
						}
					}
					File.Move(strNomeArquivoCompleto, strNomeArquivoCompletoRenomeado);
					#endregion

					adicionaDisplay("Arquivo de retorno renomeado para " + Path.GetFileName(strNomeArquivoCompletoRenomeado));

					#region [ Grava o log no BD ]
					strDescricaoLog = "Sucesso na carga do arquivo de retorno: " + strNomeArquivoCompleto + " (duração do processamento: " + intDuracaoProcessamentoEmSeg.ToString() + " segundos)";
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_CARREGA_ARQ_RETORNO;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO;
					finLog.id_boleto_cedente = _boletoCedente.id;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
					#endregion

					Global.Usuario.Defaults.FBoletoArqRetorno.pathBoletoArquivoRetorno = Path.GetDirectoryName(strNomeArquivoCompleto);

					info(ModoExibicaoMensagemRodape.Normal);
					aviso("Arquivo de retorno carregado com sucesso!!\n\n" + strNomeArquivoCompleto);
				}
				else
				{
					BD.rollbackTransacao();

					#region [ Grava o log no BD ]
					strDescricaoLog = "Falha na carga do arquivo de retorno: " + strNomeArquivoCompleto + " (" + strMsgErro + ")";
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_CARREGA_ARQ_RETORNO;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO;
					finLog.id_boleto_cedente = _boletoCedente.id;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroLog);
					#endregion

					strMsgErro = "Falha na carga do arquivo de retorno!!\n\n" + strMsgErro;
					adicionaErro(strMsgErro);
					avisoErro(strMsgErro);
				}
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

		#region [ trataInclusaoDadosSerasaTituloMovimento ]
		/// <summary>
		/// Rotina que analisa os dados do arquivo de retorno para determinar o que deve ser enviado para a Serasa através do módulo de Reciprocidade.
		/// É fundamental que esta rotina seja executada após as rotinas que tratam os boletos e o fluxo de caixa (para cada linha do arquivo de retorno), pois
		/// está sendo presumido que esses tratamentos já foram realiazados. Por exemplo, no caso de uma ocorrência '02' (entrada confirmada), esta rotina assume
		/// que o lançamento no fluxo de caixa já foi criado.
		/// </summary>
		/// <param name="intNsuBoletoArqRetorno">Identificação do registro em t_FIN_BOLETO_ARQ_RETORNO referente à esta carga do arquivo de retorno</param>
		/// <param name="idBoletoCedente">Identificação do cedente a quem pertence o arquivo de retorno que está sendo carregado</param>
		/// <param name="linhaRegistro">Dados estruturados da linha do arquivo de retorno a ser processada</param>
		/// <param name="blnSerasaGravouDados">Parâmetro de retorno que indica se esta linha do arquivo de retorno gerou dados (cadastro inicial ou atualização) para serem enviados à Serasa</param>
		/// <param name="blnSerasaOcorrenciaIgnorada">Parâmetro de retorno que indica se esta linha do arquivo de retorno foi ignorada porque a ocorrência não tem impacto sobre os dados enviados à Serasa</param>
		/// <param name="strMotivoNaoGravarDadosSerasa">Texto explicativo informando o motivo pelo qual os dados não devem ser enviados à Serasa</param>
		/// <param name="strMsgErro">Retorna a mensagem de erro, se for o caso</param>
		/// <returns>
		/// true: processamento efetuado com sucesso (não implica necessariamente que os dados foram gravados p/ envio à Serasa)
		/// false: erro no processamento
		/// </returns>
		private bool trataInclusaoDadosSerasaTituloMovimento(int intNsuBoletoArqRetorno, int idBoletoCedente, LinhaRegistroTipo1ArquivoRetorno linhaRegistro, out bool blnSerasaGravouDados, out bool blnSerasaOcorrenciaIgnorada, out String strMotivoNaoGravarDadosSerasa, out String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			bool blnTituloMovimentoCadastrar = false;
			bool blnTituloMovimentoAtualizar = false;
			bool blnTituloCadastradoSerasa;
			String[] vId;
			String[] vCnpj;
			String strIdBoletoItem;
			String strIdentificacaoOcorrencia;
			DateTime dt_cliente_desde;
			String strListaCnpjEmpresa;
			String strCnpjEmpresa;
			String strRaizCnpjEmpresa;
			String strRaizCnpjCliente;
			Cliente cliente;
			SerasaCliente serasaCliente;
			SerasaTituloMovimento tituloMovimento = new SerasaTituloMovimento();
			BoletoCedente boletoCedente;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixa;
			DsDataSource.DtbFinFluxoCaixaRow rowLancamento = null;
			List<SerasaTituloMovimento> listaTituloMovimento;
			#endregion

			strMsgErro = "";
			blnSerasaGravouDados = false;
			blnSerasaOcorrenciaIgnorada = false;
			strMotivoNaoGravarDadosSerasa = "";

			#region [ Apenas boletos do cedente que participa da Reciprocidade (apenas um) ]
			// Cada empresa participante (mesmo que seja uma filial) do programa de Reciprocidade deve trocar
			// seu próprio conjunto de arquivos c/ a Serasa.
			if (_boletoCedente.st_participante_serasa_reciprocidade == 0)
			{
				strMotivoNaoGravarDadosSerasa = "Cedente (id=" + _boletoCedente.id.ToString() + ") não participa do programa de Reciprocidade";
				return false;
			}
			#endregion

			strIdentificacaoOcorrencia = linhaRegistro.identificacaoOcorrencia.valor;

			#region [ Verifica se é uma ocorrência que necessita de tratamento ]
			// Observação: a ocorrência 17 (liquidação após baixa ou título não registrado) será ignorada neste tratamento
			// ==========  porque o título será baixado em um arquivo de retorno futuro devido ao tratamento manual feito
			//             pelo departamento financeiro.
			if (strIdentificacaoOcorrencia.Equals("02"))
			{
				// Entrada confirmada
				blnTituloMovimentoCadastrar = true;
			}
			else if (strIdentificacaoOcorrencia.Equals("06"))
			{
				// Liquidação normal
				blnTituloMovimentoAtualizar = true;
			}
			else if (strIdentificacaoOcorrencia.Equals("09"))
			{
				// Baixado automat. via arquivo
				blnTituloMovimentoAtualizar = true;
			}
			else if (strIdentificacaoOcorrencia.Equals("10"))
			{
				// Baixado conforme instruções da agência
				blnTituloMovimentoAtualizar = true;
			}
			else if (strIdentificacaoOcorrencia.Equals("12"))
			{
				// Abatimento concedido
				blnTituloMovimentoAtualizar = true;
			}
			else if (strIdentificacaoOcorrencia.Equals("13"))
			{
				// Abatimento cancelado
				blnTituloMovimentoAtualizar = true;
			}
			else if (strIdentificacaoOcorrencia.Equals("14"))
			{
				// Vencimento alterado
				blnTituloMovimentoAtualizar = true;
			}
			else if (strIdentificacaoOcorrencia.Equals("15"))
			{
				// Liquidação em cartório
				blnTituloMovimentoAtualizar = true;
			}
			else
			{
				blnSerasaOcorrenciaIgnorada = true;
				return false;
			}
			#endregion

			#region [ Consistências ]
			if (linhaRegistro == null)
			{
				strMsgErro = "A linha do registro informada é nula!!";
				return false;
			}
			if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
			{
				strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
				return false;
			}
			if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
			{
				strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
				return false;
			}
			if (intNsuBoletoArqRetorno <= 0)
			{
				strMsgErro = "Não foi fornecido o NSU do registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
				return false;
			}
			#endregion

			#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
			vId = linhaRegistro.numControleParticipante.valor.Split('=');
			strIdBoletoItem = vId[1];
			#endregion

			#region [ Consiste o valor do campo com o Id ]
			if (strIdBoletoItem == null)
			{
				strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
				return false;
			}

			if (strIdBoletoItem.Trim().Length == 0)
			{
				strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
				return false;
			}

			if (Global.converteInteiro(strIdBoletoItem) <= 0)
			{
				strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
				return false;
			}
			#endregion

			idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

			rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
			rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

			cliente = ClienteDAO.getCliente(rowBoletoPrincipal.id_cliente);
			boletoCedente = BoletoCedenteDAO.getBoletoCedente(idBoletoCedente);

			dtbFinFluxoCaixa = LancamentoFluxoCaixaDAO.obtemRegistroLancamentoByCtrlPagtoIdParcela(rowBoletoItem.id, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
			if (dtbFinFluxoCaixa.Rows.Count == 0)
			{
				strMsgErro = "Não foi localizado o lançamento no fluxo de caixa referente ao boleto (id_boleto_item = " + rowBoletoItem.id.ToString() + ")!!";
				return false;
			}
			else if (dtbFinFluxoCaixa.Rows.Count > 1)
			{
				strMsgErro = "Há mais de um lançamento no fluxo de caixa associado ao boleto (id_boleto_item = " + rowBoletoItem.id.ToString() + ")!!";
				return false;
			}
			else
			{
				rowLancamento = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixa.Rows[0];
			}

			#region [ Apenas títulos de PJ são enviados p/ Serasa ]
			if (cliente.tipo != Global.Cte.Etc.ID_PJ)
			{
				strMotivoNaoGravarDadosSerasa = "O cliente " + Global.formataCnpjCpf(cliente.cnpj_cpf) + " é pessoa física";
				return false;
			}
			#endregion

			#region [ Boletos AV não são enviados ]
			// Os boletos AV frequentemente são pagos poucos dias após a emissão do boleto e isso cria uma situação em que o cadastramento do boleto e a confirmação do pagamento
			// possam ser enviados no mesmo arquivo de remessa para o Serasa. Quando essa situação ocorre, o Serasa trata o cadastramento do boleto, mas não trata a confirmação
			// do pagamento, fazendo com que o título fique pendente e precise ser tratado através da conciliação. Por isso a Lilian decidiu não enviar os boletos AV.
			// Os boletos AV são identificados através da letra 'A' no início do número do documento.
			if (blnTituloMovimentoCadastrar && rowBoletoItem.numero_documento.Trim().ToUpper().StartsWith("A"))
			{
				strMotivoNaoGravarDadosSerasa = "O título não será enviado por se tratar de um boleto AV (documento: " + rowBoletoItem.numero_documento + ")";
				return false;
			}
			#endregion

			#region [ Títulos emitidos p/ o CNPJ da própria empresa não podem ser enviados ]
			strListaCnpjEmpresa = ComumDAO.getCampoStringTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.SERASA_RECIPROCIDADE_CNPJ_IGNORADOS);
			vCnpj = strListaCnpjEmpresa.Split('|');
			foreach (var item in vCnpj)
			{
				if (item == null) continue;
				if (Global.digitos(item).Length == 0) continue;
				strCnpjEmpresa = Global.digitos(item);
				strRaizCnpjEmpresa = Texto.leftStr(strCnpjEmpresa, 8);
				strRaizCnpjCliente = Texto.leftStr(cliente.cnpj_cpf, 8);
				if (strRaizCnpjCliente.Equals(strRaizCnpjEmpresa))
				{
					strMotivoNaoGravarDadosSerasa = "Não é permitido enviar um título cujo CNPJ do cliente seja igual ao CNPJ da própria empresa (" + Global.formataCnpjCpf(cliente.cnpj_cpf) + ")";
					return false;
				}
			}
			#endregion

			#region [ Em caso de atualização, verifica se o título foi enviado anteriormente à Serasa ]
			if (blnTituloMovimentoAtualizar)
			{
				listaTituloMovimento = SerasaDAO.getTituloMovimentoByIdBoletoItem(rowBoletoItem.id);
				blnTituloCadastradoSerasa = false;
				foreach (var titulo in listaTituloMovimento)
				{
					if (titulo.st_envio_serasa_cancelado == 0)
					{
						// IMPORTANTE: o envio dos dados à Serasa pode ser feito com diferentes periodicidades: diário, semanal, quinzenal ou mensal.
						// Portanto, pode acontecer de ficarem acumulados p/ uma mesma remessa os movimentos de cadastro (ocorrência 02-entrada confirmada)
						// e de atualização (ocorrências de pagamento, baixa, alteração de data ou valor).
						// Isso significa que, para os movimentos de atualização, não se pode testar a flag que indica se o título já está cadastrado na
						// Serasa (t_SERASA_TITULO_MOVIMENTO.st_enviado_serasa).
						blnTituloCadastradoSerasa = true;
						break;
					}
				}

				// TODO - HOMOLOGAÇÃO/DESENVOLVIMENTO
#if (DESENVOLVIMENTO || HOMOLOGACAO)
				blnTituloCadastradoSerasa = true;
#endif

				if (!blnTituloCadastradoSerasa)
				{
					strMotivoNaoGravarDadosSerasa = "Os dados de atualização do boleto (id=" + idBoletoItem.ToString() + ") não foram enviados para a Serasa porque o título não foi enviado anteriormente";
					return false;
				}
			}
			#endregion

			serasaCliente = SerasaDAO.getSerasaClienteByRaizCnpj(cliente.cnpj_cpf);

			#region [ Cliente ainda não está cadastrado em t_SERASA_CLIENTE ]
			if (serasaCliente == null)
			{
				dt_cliente_desde = SerasaDAO.getDataClienteDesde(cliente.cnpj_cpf);
				if (dt_cliente_desde == DateTime.MinValue) dt_cliente_desde = rowBoletoPrincipal.dt_cadastro;
				serasaCliente = new SerasaCliente();
				serasaCliente.dt_cliente_desde = dt_cliente_desde;
				serasaCliente.id_cliente = cliente.id;
				serasaCliente.cnpj = cliente.cnpj_cpf;
				if (!SerasaDAO.clienteInsere(Global.Usuario.usuario, serasaCliente, out strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao gravar os dados do cliente na tabela de Clientes do módulo de integração com a Serasa (CNPJ: " + Global.formataCnpjCpf(cliente.cnpj_cpf) + ")!!" + strMsgErro;
					Global.gravaLogAtividade(strMsgErro);
					throw new FinanceiroException(strMsgErro);
				}
			}
			#endregion

			#region [ Prepara dados p/ gravação em t_SERASA_TITULO_MOVIMENTO ]
			tituloMovimento.id_boleto_arq_retorno = intNsuBoletoArqRetorno;
			tituloMovimento.id_boleto_item = rowBoletoItem.id;
			tituloMovimento.id_serasa_cliente = serasaCliente.id;
			tituloMovimento.cnpj = cliente.cnpj_cpf;
			tituloMovimento.identificacao_ocorrencia_boleto = linhaRegistro.identificacaoOcorrencia.valor;
			tituloMovimento.numero_documento = BD.readToString(rowBoletoItem.numero_documento);
			tituloMovimento.nosso_numero = BD.readToString(rowBoletoItem.nosso_numero);
			tituloMovimento.digito_nosso_numero = BD.readToString(rowBoletoItem.digito_nosso_numero);
			tituloMovimento.dt_emissao = BD.readToDateTime(rowBoletoItem.dt_entrada_confirmada);
			tituloMovimento.vl_titulo = BD.readToDecimal(rowBoletoItem.valor);
			tituloMovimento.dt_vencto = BD.readToDateTime(rowBoletoItem.dt_vencto);
			#endregion

			#region [ Tratamento específico de alguns tipos de ocorrência ]
			// Observação: a ocorrência 17 (liquidação após baixa ou título não registrado) será ignorada neste tratamento
			// ==========  porque o título será baixado em um arquivo de retorno futuro devido ao tratamento manual feito
			//             pelo departamento financeiro.
			if (strIdentificacaoOcorrencia.Equals("06"))
			{
				#region [ Liquidação normal ]
				tituloMovimento.dt_pagto = rowLancamento.dt_competencia;
				tituloMovimento.vl_pago = rowLancamento.valor;
				#endregion
			}
			else if (strIdentificacaoOcorrencia.Equals("12"))
			{
				#region [ Abatimento concedido ]
				tituloMovimento.vl_titulo = rowLancamento.valor;
				#endregion
			}
			else if (strIdentificacaoOcorrencia.Equals("13"))
			{
				#region [ Abatimento cancelado ]
				tituloMovimento.vl_titulo = rowLancamento.valor;
				#endregion
			}
			else if (strIdentificacaoOcorrencia.Equals("14"))
			{
				#region [ Vencimento alterado ]
				tituloMovimento.dt_vencto = rowLancamento.dt_competencia;
				#endregion
			}
			else if (strIdentificacaoOcorrencia.Equals("15"))
			{
				#region [ Liquidação em cartório ]
				tituloMovimento.dt_pagto = rowLancamento.dt_competencia;
				tituloMovimento.vl_pago = rowLancamento.valor;
				#endregion
			}
			#endregion

			#region [ Tratamento p/ cadastramento/atualização do título na Serasa ]
			if (blnTituloMovimentoCadastrar)
			{
				#region [ Grava os dados do novo título p/ envio à Serasa ]
				if (SerasaDAO.tituloMovimentoInsere(Global.Usuario.usuario, tituloMovimento, out strMsgErro))
				{
					blnSerasaGravouDados = true;
				}
				else
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao gravar os dados do título na tabela de movimentos do módulo de integração com a Serasa [cadastramento] (id_boleto_item: " + rowBoletoItem.id.ToString() + ")!!" + strMsgErro;
					Global.gravaLogAtividade(strMsgErro);
					throw new FinanceiroException(strMsgErro);
				}
				#endregion
			}
			else if (blnTituloMovimentoAtualizar)
			{
				#region [ Grava o movimento p/ atualizar os dados na Serasa ]
				if (SerasaDAO.tituloMovimentoInsere(Global.Usuario.usuario, tituloMovimento, out strMsgErro))
				{
					blnSerasaGravouDados = true;
				}
				else
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao gravar os dados do título na tabela de movimentos do módulo de integração com a Serasa [atualização] (id_boleto_item: " + rowBoletoItem.id.ToString() + ")!!" + strMsgErro;
					Global.gravaLogAtividade(strMsgErro);
					throw new FinanceiroException(strMsgErro);
				}
				#endregion
			}
			#endregion

			return true;
		}
		#endregion

		#region [ BRADESCO: Rotinas para tratamento de cada ocorrência ]

		#region [ b237TrataOcorrencia02EntradaConfirmada ]
		private bool b237TrataOcorrencia02EntradaConfirmada(B237HeaderArqRetorno linhaHeader, B237RegTipo1ArqRetorno linhaRegistro, ref String strMsgErro)
		{
			#region [ Declarações ]
			int id_fluxo_caixa;
			int intCounter;
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strCodigoBarras = "";
			String strLinhaDigitavel = "";
			String strDescricaoLancamento;
			String strPedido;
			Global.eTipoAtualizacaoEfetuada tipoAtualizacaoEfetuada;
			DsDataSource.DtbFinBoletoItemRateioDataTable dtbFinBoletoItemRateio;
			DsDataSource.DtbFinBoletoItemRateioRow rowFinBoletoItemRateio;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistência ]
				if (linhaHeader == null)
				{
					strMsgErro = "A linha do header informada é nula!!";
					return false;
				}
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Calcula o código de barras e a linha digitável ]
				if (!Global.b237MontaLinhaDigitavelECodigoBarras(linhaHeader.numBanco.valor.Trim(), linhaRegistro.identifCedenteAgencia.valor.Trim(), linhaRegistro.identifCedenteCtaCorrente.valor.Trim(), linhaRegistro.identifCedenteCarteira.valor.Trim(), linhaRegistro.nossoNumeroSemDigito.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor), Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor), ref strCodigoBarras, ref strLinhaDigitavel))
				{
					strMsgErro = "Falha ao calcular o código de barras e a linha digitável!!";
					return false;
				}
				if (strCodigoBarras.Length == 0)
				{
					strMsgErro = "Não foi possível calcular o código de barras!!";
					return false;
				}
				if (strLinhaDigitavel.Length == 0)
				{
					strMsgErro = "Não foi possível calcular a linha digitável!!";
					return false;
				}
				#endregion

				#region [ Atualiza dados do boleto ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia02EntradaConfirmada(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strCodigoBarras,
										strLinhaDigitavel,
										Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 02 (entrada confirmada)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Gera o lançamento no fluxo de caixa ]
				strDescricaoLancamento = linhaRegistro.numeroDocumento.valor.Trim();
				if (!LancamentoFluxoCaixaDAO.insereLancamentoDevidoBoletoOcorrencia02(
									Global.Usuario.usuario,
									idBoletoItem,
									Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
									Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
									strDescricaoLancamento,
									out id_fluxo_caixa,
									out tipoAtualizacaoEfetuada,
									ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao gerar automaticamente o lançamento do fluxo de caixa durante o tratamento da ocorrência 02 (entrada confirmada)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Gera o registro no histório de pagamento do pedido, se houver pedido associado ]
				if (!PedidoHistPagtoDAO.inserePagtoDevidoBoletoOcorrencia02(
									Global.Usuario.usuario,
									id_fluxo_caixa,
									idBoletoItem,
									Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
									Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
									strDescricaoLancamento,
									ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao gerar automaticamente o registro no histórico de pagamentos do pedido durante o tratamento da ocorrência 02 (entrada confirmada)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Registra no(s) pedido(s) associado(s), se houver, que os boletos foram emitidos ]

				// Obtém os pedidos envolvidos no rateio
				dtbFinBoletoItemRateio = BoletoDAO.obtemBoletoItemRateio(idBoletoItem);

				for (intCounter = 0; intCounter < dtbFinBoletoItemRateio.Rows.Count; intCounter++)
				{
					rowFinBoletoItemRateio = (DsDataSource.DtbFinBoletoItemRateioRow)dtbFinBoletoItemRateio.Rows[intCounter];
					strPedido = rowFinBoletoItemRateio.pedido;
					if (Global.isNumeroPedido(strPedido))
					{
						if (!PedidoDAO.marcaPedidoStatusBoletoConfeccionado(
														Global.Usuario.usuario,
														strPedido,
														ref strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao assinalar no pedido que os boletos já estão confeccionados durante o tratamento da ocorrência 02 (entrada confirmada)!!" + strMsgErro;
							return false;
						}
					}
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia06LiquidacaoNormal ]
		private bool b237TrataOcorrencia06LiquidacaoNormal(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int intColNome;
			int intColSinal;
			int intColValor;
			decimal vlTitulo;
			decimal vlAbatimento;
			decimal vlDesconto;
			decimal vlMora;
			decimal vlDevido;
			decimal vlPago;
			decimal vlDiferenca;
			decimal vlDiferencaAux;
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados do pagamento na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia06LiquidacaoNormal(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Insere registro do pagamento na tabela t_FIN_BOLETO_MOVIMENTO ]
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				if (!BoletoDAO.b237BoletoMovimentoInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										rowBoletoPrincipal.id,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorIofDevido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorMora.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro na tabela t_FIN_BOLETO_MOVIMENTO durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia06(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Assinala o pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia06(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Se houver divergência no valor pago, grava ocorrência ]
				// IMPORTANTE: QUANDO O CLIENTE PAGA INDEVIDAMENTE O BOLETO COM UM VALOR MENOR, O BANCO
				// ==========  INFORMA NORMALMENTE COMO OCORRÊNCIA 06 (LIQUIDAÇÃO NORMAL) E ATRIBUI A
				// DIFERENÇA COMO DESCONTO!!
				vlPago = Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor);
				vlTitulo = Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor);
				vlAbatimento = Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor);
				vlDesconto = Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor);
				vlMora = Global.decodificaCampoMonetario(linhaRegistro.valorMora.valor);
				vlDevido = vlTitulo - vlAbatimento - vlDesconto + vlMora;
				vlDiferenca = vlDevido - vlPago;
				if ((vlDiferenca != 0) ||
					(vlPago < vlTitulo))
				{
					intColNome = 14;
					intColSinal = 5;
					intColValor = 14;
					if (vlPago < vlTitulo)
					{
						vlDiferencaAux = vlTitulo - vlAbatimento + vlMora - vlPago;
						strObsOcorrencia = "Divergência de valor" +
										   "\n" +
										   "VL Título".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlTitulo).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Abatimento".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlAbatimento).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Mora".PadLeft(intColNome, ' ') + " (+) " + Global.formataMoeda(vlMora).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Pago".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlPago).PadLeft(intColValor, ' ') +
										   "\n" +
										   "".PadLeft(intColNome + intColSinal + intColValor, '=') +
										   "\n" +
										   "Diferença".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlDiferencaAux).PadLeft(intColValor, ' ');
					}
					else
					{
						strObsOcorrencia = "Divergência no valor pago de " +
										   Global.Cte.Etc.SIMBOLO_MONETARIO + " " + Global.formataMoeda(vlDiferenca) +
										   "\n" +
										   "VL Título".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlTitulo).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Abatimento".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlAbatimento).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Desconto".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlDesconto).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Mora".PadLeft(intColNome, ' ') + " (+) " + Global.formataMoeda(vlMora).PadLeft(intColValor, ' ') +
										   "\n" +
										   "".PadLeft(intColNome + intColSinal + intColValor, '=') +
										   "\n" +
										   "VL Devido".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlDevido).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Pago".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlPago).PadLeft(intColValor, ' ') +
										   "\n" +
										   "".PadLeft(intColNome + intColSinal + intColValor, '=') +
										   "\n" +
										   "Diferença".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlDiferenca).PadLeft(intColValor, ' ');
					}

					if (!BoletoDAO.b237BoletoOcorrenciaInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										_boletoCedente.id,
										rowBoletoPrincipal.id,
										idBoletoItem,
										Global.Cte.FIN.StCampoFlag.FLAG_LIGADO,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao inserir ocorrência devido à divergência no valor pago durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
						return false;
					}
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia09BaixadoAutoViaArq ]
		private bool b237TrataOcorrencia09BaixadoAutoViaArq(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia09BaixadoAutoViaArq(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 09 (baixado automaticamente via arquivo)!!" + strMsgErro;
					return false;
				}
				#endregion

				// OBSERVAÇÃO!!
				// Por decisão do Rogério, a localização e cancelamento do lançamento de fluxo de caixa 
				// associado ao boleto será feito de forma manual. Isso obriga que se tome ciência de tudo 
				// o que o banco está baixando.
				// Somente serão atualizados no registro do lançamento do fluxo de caixa, os campos que
				// indicam se o boleto está baixado ou não, mas são campos apenas informativos (histórico), não são
				// considerados ao contabilizar o saldo do fluxo de caixa.
				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia09(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 09 (baixado automaticamente via arquivo)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia09(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 09 (baixado automaticamente via arquivo)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = "Boleto baixado sem nenhum processamento automático sobre o fluxo de caixa";
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 09 (baixado automaticamente via arquivo)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia10BaixadoConfInstrAgencia ]
		private bool b237TrataOcorrencia10BaixadoConfInstrAgencia(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia10BaixadoConfInstrAgencia(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 10 (baixado conforme instruções da agência)!!" + strMsgErro;
					return false;
				}
				#endregion

				// OBSERVAÇÃO!!
				// Por decisão do Rogério, a localização e cancelamento do lançamento de fluxo de caixa 
				// associado ao boleto será feito de forma manual. Isso obriga que se tome ciência de tudo 
				// o que o banco está baixando.
				// Somente serão atualizados no registro do lançamento do fluxo de caixa, os campos que
				// indicam se o boleto está baixado ou não, mas são campos apenas informativos (histórico), não são
				// considerados ao contabilizar o saldo do fluxo de caixa.
				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia10(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 10 (baixado conforme instruções da agência)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia10(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 10 (baixado conforme instruções da agência)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = "Boleto baixado sem nenhum processamento automático sobre o fluxo de caixa";
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 10 (baixado conforme instruções da agência)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia12AbatimentoConcedido ]
		private bool b237TrataOcorrencia12AbatimentoConcedido(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			decimal novoValorTitulo;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia12AbatimentoConcedido(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 12 (abatimento concedido)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				// Apesar de que aparentemente no arquivo de retorno o campo referente ao valor do título
				// conter a informação do valor do título inicial, ou seja, sem subtrair o valor do abatimento,
				// optou-se por calcular usando a informação armazenada no banco de dados.
				novoValorTitulo = rowBoletoItem.valor -
								  Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor);

				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia12(
										Global.Usuario.usuario,
										idBoletoItem,
										novoValorTitulo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 12 (abatimento concedido)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia12(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 12 (abatimento concedido)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = "Ocorrência apenas a título informativo. Processamento já foi devidamente realizado. Abatimento concedido: " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor));
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 12 (abatimento concedido)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia13AbatimentoCancelado ]
		private bool b237TrataOcorrencia13AbatimentoCancelado(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			String strValorComAbatimento;
			String strDtCompetencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixa;
			DsDataSource.DtbFinFluxoCaixaRow rowLancamento = null;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				dtbFinFluxoCaixa = LancamentoFluxoCaixaDAO.obtemRegistroLancamentoByCtrlPagtoIdParcela(rowBoletoItem.id, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				if (dtbFinFluxoCaixa.Rows.Count == 1) rowLancamento = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixa.Rows[0];

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia13AbatimentoCancelado(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 13 (abatimento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia13(
										Global.Usuario.usuario,
										idBoletoItem,
										rowBoletoItem.valor,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 13 (abatimento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia13(
										Global.Usuario.usuario,
										idBoletoItem,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 13 (abatimento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				if (rowLancamento == null)
				{
					strValorComAbatimento = "(?)";
					strDtCompetencia = Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_vencto);
				}
				else
				{
					strValorComAbatimento = Global.formataMoeda(rowLancamento.valor);
					strDtCompetencia = Global.formataDataDdMmYyyyComSeparador(rowLancamento.dt_competencia);
				}

				strObsOcorrencia = "Ocorrência apenas a título informativo. Processamento já foi devidamente realizado. O lançamento do fluxo de caixa em " + strDtCompetencia + " foi alterado de " + strValorComAbatimento + " para " + Global.formataMoeda(rowBoletoItem.valor);
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 13 (abatimento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia14VenctoAlterado ]
		private bool b237TrataOcorrencia14VenctoAlterado(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DateTime dtNovoVencto;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Tenta converter a nova data de vencimento ]
				if (linhaRegistro.dataVenctoTitulo.valor == null)
				{
					strMsgErro = "A nova data de vencimento não foi informada!!";
					return false;
				}
				if (linhaRegistro.dataVenctoTitulo.valor.Trim().Length == 0)
				{
					strMsgErro = "A nova data de vencimento não foi preenchida!!";
					return false;
				}
				dtNovoVencto = Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor);
				if (dtNovoVencto == DateTime.MinValue)
				{
					strMsgErro = "Não foi possível converter a nova data de vencimento em uma variável do tipo DateTime (" + linhaRegistro.dataVenctoTitulo.valor + ")!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia14VenctoAlterado(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										dtNovoVencto,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 14 (vencimento alterado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia14(
										Global.Usuario.usuario,
										idBoletoItem,
										dtNovoVencto,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 14 (vencimento alterado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia14(
										Global.Usuario.usuario,
										idBoletoItem,
										dtNovoVencto,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 14 (vencimento alterado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = "Ocorrência apenas a título informativo. Processamento já foi devidamente realizado. Vencimento alterado de " + Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_vencto) + " para " + Global.formataDataDdMmYyyyComSeparador(dtNovoVencto);
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 14 (vencimento alterado)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia15LiquidacaoEmCartorio ]
		private bool b237TrataOcorrencia15LiquidacaoEmCartorio(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoleto = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia = "";
			String strLancamentoDtCompetencia;
			String strLancamentoValor;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem = null;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixa;
			DsDataSource.DtbFinFluxoCaixaRow rowLancamento = null;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ O campo nº controle participante contém o id do registro do boleto? ]
				if (linhaRegistro.numControleParticipante.valor.Trim().Length > 0)
				{
					if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") != -1)
					{
						#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
						vId = linhaRegistro.numControleParticipante.valor.Split('=');
						strIdBoletoItem = vId[1];
						if (strIdBoletoItem != null)
						{
							if (Global.converteInteiro(strIdBoletoItem) > 0)
							{
								idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);
								rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
							}
						}
						#endregion
					}
				}
				#endregion

				#region [ É necessário descobrir o id do registro do boleto através do nosso número? ]
				if (idBoletoItem == 0)
				{
					if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor);
						if (rowBoletoItem != null)
						{
							idBoletoItem = rowBoletoItem.id;
						}
					}
				}
				#endregion

				#region [ Não conseguiu determinar o número de identificação do registro do boleto ]
				if (idBoletoItem == 0)
				{
					strMsgErro = "Não foi possível determinar o número de identificação do registro do boleto, nem mesmo através do campo nosso número!!";
					return false;
				}
				#endregion

				#region [ Obtém dados do registro principal do boleto ]
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				idBoleto = rowBoletoPrincipal.id;
				#endregion

				#region [ Obtém dados do lançamento do fluxo de caixa associado ao boleto ]
				dtbFinFluxoCaixa = LancamentoFluxoCaixaDAO.obtemRegistroLancamentoByCtrlPagtoIdParcela(rowBoletoItem.id, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				if (dtbFinFluxoCaixa.Rows.Count == 1) rowLancamento = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixa.Rows[0];
				#endregion

				#region [ Atualiza os dados do pagamento na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia15LiquidacaoEmCartorio(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Insere registro do pagamento na tabela t_FIN_BOLETO_MOVIMENTO ]
				if (!BoletoDAO.b237BoletoMovimentoInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										idBoleto,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorIofDevido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorMora.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro na tabela t_FIN_BOLETO_MOVIMENTO durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia15(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Assinala o pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia15(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência ]
				if (rowLancamento != null)
				{
					strLancamentoDtCompetencia = Global.formataDataDdMmYyyyComSeparador(rowLancamento.dt_competencia);
					strLancamentoValor = Global.formataMoeda(rowLancamento.valor);
				}
				else
				{
					strLancamentoDtCompetencia = "(?)";
					strLancamentoValor = "(?)";
				}

				strObsOcorrencia = "O lançamento do fluxo de caixa foi alterado de " +
									strLancamentoDtCompetencia +
									" = " +
									strLancamentoValor +
									" para " +
									Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor)) +
									" = " +
									Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor));

				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										_boletoCedente.id,
										idBoleto,
										idBoletoItem,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro no relatório de ocorrências de boletos durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia16TituloPagoEmCheque ]
		private bool b237TrataOcorrencia16TituloPagoEmCheque(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoletoPrincipal = 0;
			bool blnBoletoOcorrencia06 = false;
			bool blnBoletoOcorrencia15 = false;
			bool blnBoletoBaixado = false;
			DateTime dtBoletoOcorrencia06 = DateTime.MinValue;
			DateTime dtBoletoOcorrencia15 = DateTime.MinValue;
			DateTime dtBoletoBaixado = DateTime.MinValue;
			String strMsgErroAux = "";
			String strObsOcorrencia = "";
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem = null;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Tenta identificar o boleto através do "número de controle do participante" ]
				idBoletoItem = Global.decodificaBoletoNumeroControleParticipante(linhaRegistro.numControleParticipante.valor, ref strMsgErroAux);
				if (idBoletoItem > 0)
				{
					rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
					if (rowBoletoItem == null)
					{
						strMsgErro = "Não foi localizado no banco de dados o registro do boleto (id=" + idBoletoItem.ToString() + ")";
						return false;
					}
				}
				#endregion

				#region [ Tenta identificar o boleto pelo campo "nosso número"? ]
				// Há casos em que no arquivo de retorno não é informado o campo "número de controle do participante"
				if (idBoletoItem <= 0)
				{
					if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor);
						if (rowBoletoItem != null)
						{
							idBoletoItem = rowBoletoItem.id;
						}
					}
				}
				#endregion

				#region [ Conseguiu identificar o boleto? ]
				if (idBoletoItem <= 0)
				{
					strMsgErro = "Boleto não cadastrado no sistema (sem informação necessária no arquivo de retorno)";
					return false;
				}
				#endregion

				#region [ Tenta recuperar dados do boleto principal ]
				if (idBoletoItem > 0)
				{
					rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
					if (rowBoletoPrincipal != null)
					{
						idBoletoPrincipal = rowBoletoPrincipal.id;
					}
				}
				#endregion

				#region [ Atualiza os dados do pagto em cheque (valor ainda vinculado) na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia16TituloPagoEmCheque(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 16 (título pago em cheque)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Analisa se o boleto já foi liquidado ou baixado ]
				if (rowBoletoItem != null)
				{
					if (rowBoletoItem.st_boleto_ocorrencia_06 == Global.Cte.FIN.StCampoFlag.FLAG_LIGADO)
					{
						blnBoletoOcorrencia06 = true;
						dtBoletoOcorrencia06 = BD.readToDateTime(rowBoletoItem.dt_ocorrencia_banco_boleto_ocorrencia_06);
					}
					else if (rowBoletoItem.st_boleto_ocorrencia_15 == Global.Cte.FIN.StCampoFlag.FLAG_LIGADO)
					{
						blnBoletoOcorrencia15 = true;
						dtBoletoOcorrencia15 = BD.readToDateTime(rowBoletoItem.dt_ocorrencia_banco_boleto_ocorrencia_15);
					}
					else if (rowBoletoItem.st_boleto_baixado == Global.Cte.FIN.StCampoFlag.FLAG_LIGADO)
					{
						blnBoletoBaixado = true;
						dtBoletoBaixado = BD.readToDateTime(rowBoletoItem.dt_ocorrencia_banco_boleto_baixado);
					}
				}
				#endregion

				#region [ Processa a ocorrência 16 no lançamento do fluxo de caixa e no histórico de pagamentos dos pedidos ]
				if (blnBoletoOcorrencia06 || blnBoletoOcorrencia15 || blnBoletoBaixado)
				{
					#region [ Se o boleto já foi liquidado ou baixado, grava uma ocorrência ]
					if (blnBoletoOcorrencia06)
					{
						strObsOcorrencia = "Ocorrência 16 (título pago em cheque) em boleto que já foi liquidado através de uma ocorrência 06 (liquidação normal) em " + Global.formataDataDdMmYyyyComSeparador(dtBoletoOcorrencia06);
					}
					else if (blnBoletoOcorrencia15)
					{
						strObsOcorrencia = "Ocorrência 16 (título pago em cheque) em boleto que já foi liquidado através de uma ocorrência 15 (liquidação em cartório) em " + Global.formataDataDdMmYyyyComSeparador(dtBoletoOcorrencia15);
					}
					else if (blnBoletoBaixado)
					{
						strObsOcorrencia = "Ocorrência 16 (título pago em cheque) em boleto que já foi baixado em " + Global.formataDataDdMmYyyyComSeparador(dtBoletoBaixado);
					}

					if (!BoletoDAO.b237BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											idBoletoPrincipal,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao inserir registro no relatório de ocorrências de boletos durante o tratamento da ocorrência 16 (título pago em cheque)!!" + strMsgErro;
						return false;
					}
					#endregion
				}
				else
				{
					#region [ Atualiza o lançamento do fluxo de caixa ]
					if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia16(
											Global.Usuario.usuario,
											idBoletoItem,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 16 (título pago em cheque)!!" + strMsgErro;
						return false;
					}
					#endregion

					#region [ Atualiza o histórico de pagamento dos pedidos ]
					if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia16(
											Global.Usuario.usuario,
											idBoletoItem,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 16 (título pago em cheque)!!" + strMsgErro;
						return false;
					}
					#endregion
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia17LiqAposBaixaOuTitNaoRegistrado ]
		private bool b237TrataOcorrencia17LiqAposBaixaOuTitNaoRegistrado(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoleto = 0;
			int idLancamentoFluxoCaixa;
			String strObsOcorrencia = "";
			String strDescricaoPedidoHistPagto;
			String strDescricaoLancamentoFluxoCaixa;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem = null;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Tenta determinar o nº identificação do registro do boleto ]
				if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
				{
					rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor);
					if (rowBoletoItem != null)
					{
						idBoletoItem = rowBoletoItem.id;
						rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
						idBoleto = rowBoletoPrincipal.id;
					}
				}
				#endregion

				#region [ Insere registro do pagamento na tabela t_FIN_BOLETO_MOVIMENTO ]
				if (!BoletoDAO.b237BoletoMovimentoInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										idBoleto,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorIofDevido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorMora.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro na tabela t_FIN_BOLETO_MOVIMENTO durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Insere novo registro de lançamento do fluxo de caixa ]
				if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
				strObsOcorrencia += "Valor título: " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor)) +
									", Valor pago: " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor)) +
									"\n" +
									"Inserido novo lançamento no fluxo de caixa: Competência=" +
									Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor)) +
									", Valor=" + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor));

				if (rowBoletoItem != null)
					strDescricaoLancamentoFluxoCaixa = rowBoletoItem.numero_documento + " (boleto ocorrência 17)";
				else
					strDescricaoLancamentoFluxoCaixa = "Ocorrência 17 - Nosso nº " + Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor);

				if (!LancamentoFluxoCaixaDAO.insereLancamentoDevidoBoletoOcorrencia17(
										Global.Usuario.usuario,
										idBoletoItem,
										_boletoCedente.id,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										strDescricaoLancamentoFluxoCaixa,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										out idLancamentoFluxoCaixa,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Conseguiu determinar o nº identificação do registro do boleto? ]
				if (idBoletoItem > 0)
				{
					#region [ Atualiza os dados do pagamento na tabela t_FIN_BOLETO_ITEM ]
					if (!BoletoDAO.atualizaBoletoItemOcorrencia17LiqAposBaixaOuTitNaoRegistrado(
											Global.Usuario.usuario,
											idBoletoItem,
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
						return false;
					}
					#endregion

					#region [ Insere novo registro no histórico de pagamento dos pedidos ]
					if (rowBoletoItem != null)
						strDescricaoPedidoHistPagto = rowBoletoItem.numero_documento + " (boleto ocorrência 17)";
					else
						strDescricaoPedidoHistPagto = "Ocorrência 17 - Nosso nº " + Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor);

					if (!PedidoHistPagtoDAO.inserePagtoDevidoBoletoOcorrencia17(
											Global.Usuario.usuario,
											idLancamentoFluxoCaixa,
											idBoletoItem,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
											strDescricaoPedidoHistPagto,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
						return false;
					}
					#endregion
				}
				#endregion

				#region [ Grava ocorrência ]
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										_boletoCedente.id,
										idBoleto,
										idBoletoItem,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro no relatório de ocorrências de boletos durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia19ConfirmacaoRecebInstProtesto ]
		private bool b237TrataOcorrencia19ConfirmacaoRecebInstProtesto(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados da ocorrência na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia19ConfirmacaoRecebInstProtesto(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 19 (confirmação receb. inst. de protesto)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = Global.decodificaMotivoOcorrencia19(linhaRegistro.motivoCodigoOcorrencia19.valor);
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 19 (confirmação receb. inst. de protesto)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia22TituloComPagamentoCancelado ]
		private bool b237TrataOcorrencia22TituloComPagamentoCancelado(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados devido ao pagamento cancelado na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia22TituloComPagamentoCancelado(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 22 (título com pagamento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia22(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 22 (título com pagamento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia22(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 22 (título com pagamento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência ]
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				strObsOcorrencia = "Ocorrência 22 (Título com Pagamento Cancelado)";
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
								Global.Usuario.usuario,
								intNsuBoletoArqRetorno,
								_boletoCedente.id,
								rowBoletoPrincipal.id,
								idBoletoItem,
								Global.Cte.FIN.StCampoFlag.FLAG_LIGADO,
								linhaRegistro.numeroDocumento.valor,
								linhaRegistro.nossoNumeroSemDigito.valor,
								linhaRegistro.digitoNossoNumero.valor,
								Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
								Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
								linhaRegistro.identificacaoOcorrencia.valor,
								linhaRegistro.motivosRejeicoes.valor,
								linhaRegistro.motivoCodigoOcorrencia19.valor,
								Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
								strObsOcorrencia,
								linhaTextoRegistroArquivo,
								ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 22 (título com pagamento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia23EntradaTituloEmCartorio ]
		private bool b237TrataOcorrencia23EntradaTituloEmCartorio(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia23EntradaTituloEmCartorio(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 23 (entrada do título em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia23(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 23 (entrada do título em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia23(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 23 (entrada do título em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia24EntradaRejeitadaCepIrregular ]
		private bool b237TrataOcorrencia24EntradaRejeitadaCepIrregular(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados da última ocorrência na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia24CepIrregular(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o boleto com os dados da última ocorrência!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Obtém dados do registro principal do boleto ]
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				#endregion

				#region [ Grava ocorrência ]
				strObsOcorrencia = "CEP inválido";
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										_boletoCedente.id,
										rowBoletoPrincipal.id,
										idBoletoItem,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir novo registro de ocorrência durante o tratamento da ocorrência " + linhaRegistro.identificacaoOcorrencia.valor + "!!\nNº documento: " + linhaRegistro.numeroDocumento.valor + ", nosso número: " + Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor) + ", nº ctrl: " + linhaRegistro.numControleParticipante.valor + "\n" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia28DebitoTarifasCustas ]
		private bool b237TrataOcorrencia28DebitoTarifasCustas(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoletoPrincipal = 0;
			String strMsgErroAux = "";
			String strObsOcorrencia = "";
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Tenta identificar o boleto através do "número de controle do participante" ]
				idBoletoItem = Global.decodificaBoletoNumeroControleParticipante(linhaRegistro.numControleParticipante.valor, ref strMsgErroAux);
				if (idBoletoItem > 0)
				{
					rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				}
				#endregion

				#region [ Tenta identificar o boleto pelo campo "nosso número"? ]
				// Há casos em que no arquivo de retorno não é informado o campo "número de controle do participante"
				if (idBoletoItem <= 0)
				{
					if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor);
						if (rowBoletoItem != null)
						{
							idBoletoItem = rowBoletoItem.id;
						}
					}
				}
				#endregion

				#region [ Conseguiu identificar o boleto? ]
				if (idBoletoItem <= 0)
				{
					if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
					strObsOcorrencia += "Boleto não cadastrado no sistema (sem informação necessária no arquivo de retorno)";
				}
				#endregion

				#region [ Tenta recuperar dados do boleto principal ]
				if (idBoletoItem > 0)
				{
					rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
					if (rowBoletoPrincipal != null)
					{
						idBoletoPrincipal = rowBoletoPrincipal.id;
					}
				}
				#endregion

				#region [ Atualiza os dados da ocorrência na tabela t_FIN_BOLETO_ITEM ]
				if (idBoletoItem > 0)
				{
					if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia28DebitoTarifasCustas(
											Global.Usuario.usuario,
											idBoletoItem,
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											linhaRegistro,
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 28 (débito de tarifas/custas)!!" + strMsgErro;
						return false;
					}
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]

				if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "03"))
				{
					#region [ Obs ocorrência: Tarifa de sustação (motivo 03) usando campo despesas de cobrança ]
					if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
					strObsOcorrencia += "Tarifa de sustação = " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor));
					#endregion
				}
				else if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "04"))
				{
					#region [ Obs ocorrência: Tarifa de protesto (motivo 04) usando campo despesas de cobrança ]
					if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
					strObsOcorrencia += "Tarifa de protesto = " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor));
					#endregion
				}

				#region [ Obs ocorrência: Custas de protesto (motivo 08) usando campo outras despesas ]
				if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "08"))
				{
					if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
					strObsOcorrencia += "Custas de protesto = " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor));
				}
				#endregion

				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											idBoletoPrincipal,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 28 (débito de tarifas/custas)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrencia34RetiradoCartorioManutencaoCarteira ]
		private bool b237TrataOcorrencia34RetiradoCartorioManutencaoCarteira(
								int intNsuBoletoArqRetorno,
								B237RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b237AtualizaBoletoItemOcorrencia34RetiradoCartorioManutencaoCarteira(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 34 (retirado de cartório e manutenção carteira)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia34(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 34 (retirado de cartório e manutenção carteira)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia34(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 34 (retirado de cartório e manutenção carteira)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b237TrataOcorrenciaValaComum ]
		private bool b237TrataOcorrenciaValaComum(int idArqRetorno, B237RegTipo1ArqRetorno linhaRegistro, String linhaTextoRegistroArquivo, ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoleto = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia = "";
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistência ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				#endregion

				#region [ Possui nº controle do participante (t_FIN_BOLETO_ITEM.id)? ]
				if (linhaRegistro.numControleParticipante.valor.Trim().Length > 0)
				{
					if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") != -1)
					{
						vId = linhaRegistro.numControleParticipante.valor.Split('=');
						strIdBoletoItem = vId[1];
						if (strIdBoletoItem != null)
						{
							if (strIdBoletoItem.Trim().Length > 0)
							{
								idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);
							}
						}
					}
				}
				#endregion

				#region [ Se não possui nº controle do participante, pesquisa por 'nosso número' ]
				if (idBoletoItem == 0)
				{
					if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor));
						if (rowBoletoItem != null) idBoletoItem = rowBoletoItem.id;
					}
				}
				#endregion

				#region [ Se não possui nº controle do participante, pesquisa por 'nº documento' ]
				if (idBoletoItem == 0)
				{
					if (linhaRegistro.numeroDocumento.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNumeroDocumento(_boletoCedente.id, linhaRegistro.numeroDocumento.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor));
						if (rowBoletoItem != null) idBoletoItem = rowBoletoItem.id;
					}
				}
				#endregion

				#region [ Se conseguiu identificar o registro do boleto, atualiza os dados ]
				if (idBoletoItem > 0)
				{
					if (!BoletoDAO.atualizaBoletoItemOcorrenciaValaComum(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento de vala comum da ocorrência " + linhaRegistro.identificacaoOcorrencia.valor + "!!" + strMsgErro;
						return false;
					}
				}
				#endregion

				#region [ Se conseguiu identificar o registro do boleto, obtém dados do registro principal ]
				if (idBoletoItem > 0)
				{
					rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
					if (rowBoletoPrincipal != null) idBoleto = rowBoletoPrincipal.id;
				}
				#endregion

				#region [ Insere registro de ocorrência na vala comum ]
				if (!BoletoDAO.b237BoletoOcorrenciaInsere(
										Global.Usuario.usuario,
										idArqRetorno,
										_boletoCedente.id,
										idBoleto,
										idBoletoItem,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir o registro no BD durante o tratamento de vala comum da ocorrência " + linhaRegistro.identificacaoOcorrencia.valor + "!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#endregion

		#region [ SAFRA: Rotinas para tratamento de cada ocorrência ]

		#region [ b422TrataOcorrencia02EntradaConfirmada ]
		private bool b422TrataOcorrencia02EntradaConfirmada(B422HeaderArqRetorno linhaHeader, B422RegTipo1ArqRetorno linhaRegistro, ref String strMsgErro)
		{
			#region [ Declarações ]
			int id_fluxo_caixa;
			int intCounter;
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strCodigoBarras = "";
			String strLinhaDigitavel = "";
			String strDescricaoLancamento;
			String strPedido;
			String strAgencia;
			String strContaCorrente;
			Global.eTipoAtualizacaoEfetuada tipoAtualizacaoEfetuada;
			DsDataSource.DtbFinBoletoItemRateioDataTable dtbFinBoletoItemRateio;
			DsDataSource.DtbFinBoletoItemRateioRow rowFinBoletoItemRateio;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistência ]
				if (linhaHeader == null)
				{
					strMsgErro = "A linha do header informada é nula!!";
					return false;
				}
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Calcula o código de barras e a linha digitável ]
				// O campo 'Cod. Empresa' do layout Safra é formado por: Ag(5) + Cta Cob (9)
				strAgencia = Texto.leftStr(linhaRegistro.codEmpresa.valor, 5);
				strContaCorrente = Texto.rightStr(linhaRegistro.codEmpresa.valor, 9);
				if (!Global.b422MontaLinhaDigitavelECodigoBarras(linhaHeader.numBanco.valor.Trim(), strAgencia.Trim(), strContaCorrente.Trim(), linhaRegistro.carteira.valor.Trim(), linhaRegistro.nossoNumeroSemDigito.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor), Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor), ref strCodigoBarras, ref strLinhaDigitavel))
				{
					strMsgErro = "Falha ao calcular o código de barras e a linha digitável!!";
					return false;
				}
				if (strCodigoBarras.Length == 0)
				{
					strMsgErro = "Não foi possível calcular o código de barras!!";
					return false;
				}
				if (strLinhaDigitavel.Length == 0)
				{
					strMsgErro = "Não foi possível calcular a linha digitável!!";
					return false;
				}
				#endregion

				#region [ Atualiza dados do boleto ]
				if (!BoletoDAO.b422AtualizaBoletoItemOcorrencia02EntradaConfirmada(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strCodigoBarras,
										strLinhaDigitavel,
										Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor),
										linhaRegistro.bancoCobrador.valor,
										linhaRegistro.agenciaCobradora.valor,
										linhaRegistro.dataCredito.valor,
										linhaRegistro.codBeneficiarioTransferido.valor,
										linhaRegistro.indicadorEntradaDDA.valor,
										linhaRegistro.meioLiquidacao.valor,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 02 (entrada confirmada)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Gera o lançamento no fluxo de caixa ]
				strDescricaoLancamento = linhaRegistro.numeroDocumento.valor.Trim();
				if (!LancamentoFluxoCaixaDAO.insereLancamentoDevidoBoletoOcorrencia02(
									Global.Usuario.usuario,
									idBoletoItem,
									Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
									Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
									strDescricaoLancamento,
									out id_fluxo_caixa,
									out tipoAtualizacaoEfetuada,
									ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao gerar automaticamente o lançamento do fluxo de caixa durante o tratamento da ocorrência 02 (entrada confirmada)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Gera o registro no histório de pagamento do pedido, se houver pedido associado ]
				if (!PedidoHistPagtoDAO.inserePagtoDevidoBoletoOcorrencia02(
									Global.Usuario.usuario,
									id_fluxo_caixa,
									idBoletoItem,
									Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
									Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
									strDescricaoLancamento,
									ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao gerar automaticamente o registro no histórico de pagamentos do pedido durante o tratamento da ocorrência 02 (entrada confirmada)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Registra no(s) pedido(s) associado(s), se houver, que os boletos foram emitidos ]

				// Obtém os pedidos envolvidos no rateio
				dtbFinBoletoItemRateio = BoletoDAO.obtemBoletoItemRateio(idBoletoItem);

				for (intCounter = 0; intCounter < dtbFinBoletoItemRateio.Rows.Count; intCounter++)
				{
					rowFinBoletoItemRateio = (DsDataSource.DtbFinBoletoItemRateioRow)dtbFinBoletoItemRateio.Rows[intCounter];
					strPedido = rowFinBoletoItemRateio.pedido;
					if (Global.isNumeroPedido(strPedido))
					{
						if (!PedidoDAO.marcaPedidoStatusBoletoConfeccionado(
														Global.Usuario.usuario,
														strPedido,
														ref strMsgErro))
						{
							if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
							strMsgErro = "Falha ao assinalar no pedido que os boletos já estão confeccionados durante o tratamento da ocorrência 02 (entrada confirmada)!!" + strMsgErro;
							return false;
						}
					}
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia06LiquidacaoNormal ]
		private bool b422TrataOcorrencia06LiquidacaoNormal(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int intColNome;
			int intColSinal;
			int intColValor;
			decimal vlTitulo;
			decimal vlAbatimento;
			decimal vlDesconto;
			decimal vlMora;
			decimal vlDevido;
			decimal vlPago;
			decimal vlDiferenca;
			decimal vlDiferencaAux;
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados do pagamento na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b422AtualizaBoletoItemOcorrencia06LiquidacaoNormal(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Insere registro do pagamento na tabela t_FIN_BOLETO_MOVIMENTO ]
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				if (!BoletoDAO.b422BoletoMovimentoInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										rowBoletoPrincipal.id,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorIofDevido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorMora.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro na tabela t_FIN_BOLETO_MOVIMENTO durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia06(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Assinala o pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia06(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Se houver divergência no valor pago, grava ocorrência ]
				// IMPORTANTE: QUANDO O CLIENTE PAGA INDEVIDAMENTE O BOLETO COM UM VALOR MENOR, O BANCO
				// ==========  INFORMA NORMALMENTE COMO OCORRÊNCIA 06 (LIQUIDAÇÃO NORMAL) E ATRIBUI A
				// DIFERENÇA COMO DESCONTO!!
				vlPago = Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor);
				vlTitulo = Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor);
				vlAbatimento = Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor);
				vlDesconto = Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor);
				vlMora = Global.decodificaCampoMonetario(linhaRegistro.valorMora.valor);
				vlDevido = vlTitulo - vlAbatimento - vlDesconto + vlMora;
				vlDiferenca = vlDevido - vlPago;
				if ((vlDiferenca != 0) ||
					(vlPago < vlTitulo))
				{
					intColNome = 14;
					intColSinal = 5;
					intColValor = 14;
					if (vlPago < vlTitulo)
					{
						vlDiferencaAux = vlTitulo - vlAbatimento + vlMora - vlPago;
						strObsOcorrencia = "Divergência de valor" +
										   "\n" +
										   "VL Título".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlTitulo).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Abatimento".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlAbatimento).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Mora".PadLeft(intColNome, ' ') + " (+) " + Global.formataMoeda(vlMora).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Pago".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlPago).PadLeft(intColValor, ' ') +
										   "\n" +
										   "".PadLeft(intColNome + intColSinal + intColValor, '=') +
										   "\n" +
										   "Diferença".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlDiferencaAux).PadLeft(intColValor, ' ');
					}
					else
					{
						strObsOcorrencia = "Divergência no valor pago de " +
										   Global.Cte.Etc.SIMBOLO_MONETARIO + " " + Global.formataMoeda(vlDiferenca) +
										   "\n" +
										   "VL Título".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlTitulo).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Abatimento".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlAbatimento).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Desconto".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlDesconto).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Mora".PadLeft(intColNome, ' ') + " (+) " + Global.formataMoeda(vlMora).PadLeft(intColValor, ' ') +
										   "\n" +
										   "".PadLeft(intColNome + intColSinal + intColValor, '=') +
										   "\n" +
										   "VL Devido".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlDevido).PadLeft(intColValor, ' ') +
										   "\n" +
										   "VL Pago".PadLeft(intColNome, ' ') + " (-) " + Global.formataMoeda(vlPago).PadLeft(intColValor, ' ') +
										   "\n" +
										   "".PadLeft(intColNome + intColSinal + intColValor, '=') +
										   "\n" +
										   "Diferença".PadLeft(intColNome, ' ') + "".PadLeft(intColSinal, ' ') + Global.formataMoeda(vlDiferenca).PadLeft(intColValor, ' ');
					}

					if (!BoletoDAO.b422BoletoOcorrenciaInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										_boletoCedente.id,
										rowBoletoPrincipal.id,
										idBoletoItem,
										Global.Cte.FIN.StCampoFlag.FLAG_LIGADO,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao inserir ocorrência devido à divergência no valor pago durante o tratamento da ocorrência 06 (liquidação normal)!!" + strMsgErro;
						return false;
					}
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia09BaixadoAutomaticamente ]
		private bool b422TrataOcorrencia09BaixadoAutomaticamente(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b422AtualizaBoletoItemOcorrencia09BaixadoAutomaticamente(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 09 (baixado automaticamente)!!" + strMsgErro;
					return false;
				}
				#endregion

				// OBSERVAÇÃO!!
				// Por decisão do Rogério, a localização e cancelamento do lançamento de fluxo de caixa 
				// associado ao boleto será feito de forma manual. Isso obriga que se tome ciência de tudo 
				// o que o banco está baixando.
				// Somente serão atualizados no registro do lançamento do fluxo de caixa, os campos que
				// indicam se o boleto está baixado ou não, mas são campos apenas informativos (histórico), não são
				// considerados ao contabilizar o saldo do fluxo de caixa.
				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia09(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 09 (baixado automaticamente)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia09(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 09 (baixado automaticamente)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = "Boleto baixado sem nenhum processamento automático sobre o fluxo de caixa";
				if (!BoletoDAO.b422BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.codRejeicao.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 09 (baixado automaticamente)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia10BaixadoConfInstrucoes ]
		private bool b422TrataOcorrencia10BaixadoConfInstrucoes(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b422AtualizaBoletoItemOcorrencia10BaixadoConfInstrucoes(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 10 (baixado conforme instruções)!!" + strMsgErro;
					return false;
				}
				#endregion

				// OBSERVAÇÃO!!
				// Por decisão do Rogério, a localização e cancelamento do lançamento de fluxo de caixa 
				// associado ao boleto será feito de forma manual. Isso obriga que se tome ciência de tudo 
				// o que o banco está baixando.
				// Somente serão atualizados no registro do lançamento do fluxo de caixa, os campos que
				// indicam se o boleto está baixado ou não, mas são campos apenas informativos (histórico), não são
				// considerados ao contabilizar o saldo do fluxo de caixa.
				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia10(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 10 (baixado conforme instruções)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia10(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 10 (baixado conforme instruções)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = "Boleto baixado sem nenhum processamento automático sobre o fluxo de caixa";
				if (!BoletoDAO.b422BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.codRejeicao.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 10 (baixado conforme instruções)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia12AbatimentoConcedido ]
		private bool b422TrataOcorrencia12AbatimentoConcedido(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			decimal novoValorTitulo;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b422AtualizaBoletoItemOcorrencia12AbatimentoConcedido(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 12 (abatimento concedido)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				// Apesar de que aparentemente no arquivo de retorno o campo referente ao valor do título
				// conter a informação do valor do título inicial, ou seja, sem subtrair o valor do abatimento,
				// optou-se por calcular usando a informação armazenada no banco de dados.
				novoValorTitulo = rowBoletoItem.valor -
								  Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor);

				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia12(
										Global.Usuario.usuario,
										idBoletoItem,
										novoValorTitulo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 12 (abatimento concedido)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia12(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 12 (abatimento concedido)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = "Ocorrência apenas a título informativo. Processamento já foi devidamente realizado. Abatimento concedido: " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor));
				if (!BoletoDAO.b422BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.codRejeicao.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 12 (abatimento concedido)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia13AbatimentoCancelado ]
		private bool b422TrataOcorrencia13AbatimentoCancelado(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			String strValorComAbatimento;
			String strDtCompetencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixa;
			DsDataSource.DtbFinFluxoCaixaRow rowLancamento = null;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				dtbFinFluxoCaixa = LancamentoFluxoCaixaDAO.obtemRegistroLancamentoByCtrlPagtoIdParcela(rowBoletoItem.id, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				if (dtbFinFluxoCaixa.Rows.Count == 1) rowLancamento = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixa.Rows[0];

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b422AtualizaBoletoItemOcorrencia13AbatimentoCancelado(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 13 (abatimento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia13(
										Global.Usuario.usuario,
										idBoletoItem,
										rowBoletoItem.valor,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 13 (abatimento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia13(
										Global.Usuario.usuario,
										idBoletoItem,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 13 (abatimento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				if (rowLancamento == null)
				{
					strValorComAbatimento = "(?)";
					strDtCompetencia = Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_vencto);
				}
				else
				{
					strValorComAbatimento = Global.formataMoeda(rowLancamento.valor);
					strDtCompetencia = Global.formataDataDdMmYyyyComSeparador(rowLancamento.dt_competencia);
				}

				strObsOcorrencia = "Ocorrência apenas a título informativo. Processamento já foi devidamente realizado. O lançamento do fluxo de caixa em " + strDtCompetencia + " foi alterado de " + strValorComAbatimento + " para " + Global.formataMoeda(rowBoletoItem.valor);
				if (!BoletoDAO.b422BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.codRejeicao.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 13 (abatimento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia14VenctoAlterado ]
		private bool b422TrataOcorrencia14VenctoAlterado(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DateTime dtNovoVencto;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Tenta converter a nova data de vencimento ]
				if (linhaRegistro.dataVenctoTitulo.valor == null)
				{
					strMsgErro = "A nova data de vencimento não foi informada!!";
					return false;
				}
				if (linhaRegistro.dataVenctoTitulo.valor.Trim().Length == 0)
				{
					strMsgErro = "A nova data de vencimento não foi preenchida!!";
					return false;
				}
				dtNovoVencto = Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor);
				if (dtNovoVencto == DateTime.MinValue)
				{
					strMsgErro = "Não foi possível converter a nova data de vencimento em uma variável do tipo DateTime (" + linhaRegistro.dataVenctoTitulo.valor + ")!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b422AtualizaBoletoItemOcorrencia14VenctoAlterado(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										dtNovoVencto,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 14 (vencimento alterado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia14(
										Global.Usuario.usuario,
										idBoletoItem,
										dtNovoVencto,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 14 (vencimento alterado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza no pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia14(
										Global.Usuario.usuario,
										idBoletoItem,
										dtNovoVencto,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 14 (vencimento alterado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = "Ocorrência apenas a título informativo. Processamento já foi devidamente realizado. Vencimento alterado de " + Global.formataDataDdMmYyyyComSeparador(rowBoletoItem.dt_vencto) + " para " + Global.formataDataDdMmYyyyComSeparador(dtNovoVencto);
				if (!BoletoDAO.b422BoletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.codRejeicao.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 14 (vencimento alterado)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia15LiquidacaoEmCartorio ]
		private bool b422TrataOcorrencia15LiquidacaoEmCartorio(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoleto = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia = "";
			String strLancamentoDtCompetencia;
			String strLancamentoValor;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem = null;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinFluxoCaixaDataTable dtbFinFluxoCaixa;
			DsDataSource.DtbFinFluxoCaixaRow rowLancamento = null;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ O campo nº controle participante contém o id do registro do boleto? ]
				if (linhaRegistro.numControleParticipante.valor.Trim().Length > 0)
				{
					if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") != -1)
					{
						#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
						vId = linhaRegistro.numControleParticipante.valor.Split('=');
						strIdBoletoItem = vId[1];
						if (strIdBoletoItem != null)
						{
							if (Global.converteInteiro(strIdBoletoItem) > 0)
							{
								idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);
								rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
							}
						}
						#endregion
					}
				}
				#endregion

				#region [ É necessário descobrir o id do registro do boleto através do nosso número? ]
				if (idBoletoItem == 0)
				{
					if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor);
						if (rowBoletoItem != null)
						{
							idBoletoItem = rowBoletoItem.id;
						}
					}
				}
				#endregion

				#region [ Não conseguiu determinar o número de identificação do registro do boleto ]
				if (idBoletoItem == 0)
				{
					strMsgErro = "Não foi possível determinar o número de identificação do registro do boleto, nem mesmo através do campo nosso número!!";
					return false;
				}
				#endregion

				#region [ Obtém dados do registro principal do boleto ]
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				idBoleto = rowBoletoPrincipal.id;
				#endregion

				#region [ Obtém dados do lançamento do fluxo de caixa associado ao boleto ]
				dtbFinFluxoCaixa = LancamentoFluxoCaixaDAO.obtemRegistroLancamentoByCtrlPagtoIdParcela(rowBoletoItem.id, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
				if (dtbFinFluxoCaixa.Rows.Count == 1) rowLancamento = (DsDataSource.DtbFinFluxoCaixaRow)dtbFinFluxoCaixa.Rows[0];
				#endregion

				#region [ Atualiza os dados do pagamento na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.b422AtualizaBoletoItemOcorrencia15LiquidacaoEmCartorio(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Insere registro do pagamento na tabela t_FIN_BOLETO_MOVIMENTO ]
				if (!BoletoDAO.b422BoletoMovimentoInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										idBoleto,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorIofDevido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorMora.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro na tabela t_FIN_BOLETO_MOVIMENTO durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia15(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Assinala o pagamento no histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia15(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência ]
				if (rowLancamento != null)
				{
					strLancamentoDtCompetencia = Global.formataDataDdMmYyyyComSeparador(rowLancamento.dt_competencia);
					strLancamentoValor = Global.formataMoeda(rowLancamento.valor);
				}
				else
				{
					strLancamentoDtCompetencia = "(?)";
					strLancamentoValor = "(?)";
				}

				strObsOcorrencia = "O lançamento do fluxo de caixa foi alterado de " +
									strLancamentoDtCompetencia +
									" = " +
									strLancamentoValor +
									" para " +
									Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor)) +
									" = " +
									Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor));

				if (!BoletoDAO.b422BoletoOcorrenciaInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										_boletoCedente.id,
										idBoleto,
										idBoletoItem,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro no relatório de ocorrências de boletos durante o tratamento da ocorrência 15 (liquidação em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia16TituloPagoEmCheque ]
		xxxx
		private bool b422TrataOcorrencia16TituloPagoEmCheque(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoletoPrincipal = 0;
			bool blnBoletoOcorrencia06 = false;
			bool blnBoletoOcorrencia15 = false;
			bool blnBoletoBaixado = false;
			DateTime dtBoletoOcorrencia06 = DateTime.MinValue;
			DateTime dtBoletoOcorrencia15 = DateTime.MinValue;
			DateTime dtBoletoBaixado = DateTime.MinValue;
			String strMsgErroAux = "";
			String strObsOcorrencia = "";
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem = null;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Tenta identificar o boleto através do "número de controle do participante" ]
				idBoletoItem = Global.decodificaBoletoNumeroControleParticipante(linhaRegistro.numControleParticipante.valor, ref strMsgErroAux);
				if (idBoletoItem > 0)
				{
					rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
					if (rowBoletoItem == null)
					{
						strMsgErro = "Não foi localizado no banco de dados o registro do boleto (id=" + idBoletoItem.ToString() + ")";
						return false;
					}
				}
				#endregion

				#region [ Tenta identificar o boleto pelo campo "nosso número"? ]
				// Há casos em que no arquivo de retorno não é informado o campo "número de controle do participante"
				if (idBoletoItem <= 0)
				{
					if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor);
						if (rowBoletoItem != null)
						{
							idBoletoItem = rowBoletoItem.id;
						}
					}
				}
				#endregion

				#region [ Conseguiu identificar o boleto? ]
				if (idBoletoItem <= 0)
				{
					strMsgErro = "Boleto não cadastrado no sistema (sem informação necessária no arquivo de retorno)";
					return false;
				}
				#endregion

				#region [ Tenta recuperar dados do boleto principal ]
				if (idBoletoItem > 0)
				{
					rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
					if (rowBoletoPrincipal != null)
					{
						idBoletoPrincipal = rowBoletoPrincipal.id;
					}
				}
				#endregion

				#region [ Atualiza os dados do pagto em cheque (valor ainda vinculado) na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.atualizaBoletoItemOcorrencia16TituloPagoEmCheque(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 16 (título pago em cheque)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Analisa se o boleto já foi liquidado ou baixado ]
				if (rowBoletoItem != null)
				{
					if (rowBoletoItem.st_boleto_ocorrencia_06 == Global.Cte.FIN.StCampoFlag.FLAG_LIGADO)
					{
						blnBoletoOcorrencia06 = true;
						dtBoletoOcorrencia06 = BD.readToDateTime(rowBoletoItem.dt_ocorrencia_banco_boleto_ocorrencia_06);
					}
					else if (rowBoletoItem.st_boleto_ocorrencia_15 == Global.Cte.FIN.StCampoFlag.FLAG_LIGADO)
					{
						blnBoletoOcorrencia15 = true;
						dtBoletoOcorrencia15 = BD.readToDateTime(rowBoletoItem.dt_ocorrencia_banco_boleto_ocorrencia_15);
					}
					else if (rowBoletoItem.st_boleto_baixado == Global.Cte.FIN.StCampoFlag.FLAG_LIGADO)
					{
						blnBoletoBaixado = true;
						dtBoletoBaixado = BD.readToDateTime(rowBoletoItem.dt_ocorrencia_banco_boleto_baixado);
					}
				}
				#endregion

				#region [ Processa a ocorrência 16 no lançamento do fluxo de caixa e no histórico de pagamentos dos pedidos ]
				if (blnBoletoOcorrencia06 || blnBoletoOcorrencia15 || blnBoletoBaixado)
				{
					#region [ Se o boleto já foi liquidado ou baixado, grava uma ocorrência ]
					if (blnBoletoOcorrencia06)
					{
						strObsOcorrencia = "Ocorrência 16 (título pago em cheque) em boleto que já foi liquidado através de uma ocorrência 06 (liquidação normal) em " + Global.formataDataDdMmYyyyComSeparador(dtBoletoOcorrencia06);
					}
					else if (blnBoletoOcorrencia15)
					{
						strObsOcorrencia = "Ocorrência 16 (título pago em cheque) em boleto que já foi liquidado através de uma ocorrência 15 (liquidação em cartório) em " + Global.formataDataDdMmYyyyComSeparador(dtBoletoOcorrencia15);
					}
					else if (blnBoletoBaixado)
					{
						strObsOcorrencia = "Ocorrência 16 (título pago em cheque) em boleto que já foi baixado em " + Global.formataDataDdMmYyyyComSeparador(dtBoletoBaixado);
					}

					if (!BoletoDAO.boletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											idBoletoPrincipal,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao inserir registro no relatório de ocorrências de boletos durante o tratamento da ocorrência 16 (título pago em cheque)!!" + strMsgErro;
						return false;
					}
					#endregion
				}
				else
				{
					#region [ Atualiza o lançamento do fluxo de caixa ]
					if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia16(
											Global.Usuario.usuario,
											idBoletoItem,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 16 (título pago em cheque)!!" + strMsgErro;
						return false;
					}
					#endregion

					#region [ Atualiza o histórico de pagamento dos pedidos ]
					if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia16(
											Global.Usuario.usuario,
											idBoletoItem,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 16 (título pago em cheque)!!" + strMsgErro;
						return false;
					}
					#endregion
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia17LiqAposBaixaOuTitNaoRegistrado ]
		private bool b422TrataOcorrencia17LiqAposBaixaOuTitNaoRegistrado(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoleto = 0;
			int idLancamentoFluxoCaixa;
			String strObsOcorrencia = "";
			String strDescricaoPedidoHistPagto;
			String strDescricaoLancamentoFluxoCaixa;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem = null;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Tenta determinar o nº identificação do registro do boleto ]
				if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
				{
					rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor);
					if (rowBoletoItem != null)
					{
						idBoletoItem = rowBoletoItem.id;
						rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
						idBoleto = rowBoletoPrincipal.id;
					}
				}
				#endregion

				#region [ Insere registro do pagamento na tabela t_FIN_BOLETO_MOVIMENTO ]
				if (!BoletoDAO.boletoMovimentoInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										idBoleto,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorIofDevido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorAbatimentoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorDescontoConcedido.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorMora.valor),
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro na tabela t_FIN_BOLETO_MOVIMENTO durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Insere novo registro de lançamento do fluxo de caixa ]
				if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
				strObsOcorrencia += "Valor título: " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor)) +
									", Valor pago: " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor)) +
									"\n" +
									"Inserido novo lançamento no fluxo de caixa: Competência=" +
									Global.formataDataDdMmYyyyComSeparador(Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor)) +
									", Valor=" + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor));

				if (rowBoletoItem != null)
					strDescricaoLancamentoFluxoCaixa = rowBoletoItem.numero_documento + " (boleto ocorrência 17)";
				else
					strDescricaoLancamentoFluxoCaixa = "Ocorrência 17 - Nosso nº " + Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor);

				if (!LancamentoFluxoCaixaDAO.insereLancamentoDevidoBoletoOcorrencia17(
										Global.Usuario.usuario,
										idBoletoItem,
										_boletoCedente.id,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
										strDescricaoLancamentoFluxoCaixa,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										out idLancamentoFluxoCaixa,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Conseguiu determinar o nº identificação do registro do boleto? ]
				if (idBoletoItem > 0)
				{
					#region [ Atualiza os dados do pagamento na tabela t_FIN_BOLETO_ITEM ]
					if (!BoletoDAO.atualizaBoletoItemOcorrencia17LiqAposBaixaOuTitNaoRegistrado(
											Global.Usuario.usuario,
											idBoletoItem,
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
						return false;
					}
					#endregion

					#region [ Insere novo registro no histórico de pagamento dos pedidos ]
					if (rowBoletoItem != null)
						strDescricaoPedidoHistPagto = rowBoletoItem.numero_documento + " (boleto ocorrência 17)";
					else
						strDescricaoPedidoHistPagto = "Ocorrência 17 - Nosso nº " + Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor);

					if (!PedidoHistPagtoDAO.inserePagtoDevidoBoletoOcorrencia17(
											Global.Usuario.usuario,
											idLancamentoFluxoCaixa,
											idBoletoItem,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataCredito.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorPago.valor),
											strDescricaoPedidoHistPagto,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
						return false;
					}
					#endregion
				}
				#endregion

				#region [ Grava ocorrência ]
				if (!BoletoDAO.boletoOcorrenciaInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										_boletoCedente.id,
										idBoleto,
										idBoletoItem,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir registro no relatório de ocorrências de boletos durante o tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia19ConfirmacaoRecebInstProtesto ]
		private bool b422TrataOcorrencia19ConfirmacaoRecebInstProtesto(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);

				#region [ Atualiza os dados da ocorrência na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.atualizaBoletoItemOcorrencia19ConfirmacaoRecebInstProtesto(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 19 (confirmação receb. inst. de protesto)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]
				strObsOcorrencia = Global.decodificaMotivoOcorrencia19(linhaRegistro.motivoCodigoOcorrencia19.valor);
				if (!BoletoDAO.boletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											rowBoletoPrincipal.id,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 19 (confirmação receb. inst. de protesto)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia22TituloComPagamentoCancelado ]
		private bool b422TrataOcorrencia22TituloComPagamentoCancelado(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados devido ao pagamento cancelado na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.atualizaBoletoItemOcorrencia22TituloComPagamentoCancelado(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 22 (título com pagamento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia22(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 22 (título com pagamento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia22(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 22 (título com pagamento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Grava ocorrência ]
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				strObsOcorrencia = "Ocorrência 22 (Título com Pagamento Cancelado)";
				if (!BoletoDAO.boletoOcorrenciaInsere(
								Global.Usuario.usuario,
								intNsuBoletoArqRetorno,
								_boletoCedente.id,
								rowBoletoPrincipal.id,
								idBoletoItem,
								Global.Cte.FIN.StCampoFlag.FLAG_LIGADO,
								linhaRegistro.numeroDocumento.valor,
								linhaRegistro.nossoNumeroSemDigito.valor,
								linhaRegistro.digitoNossoNumero.valor,
								Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
								Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
								linhaRegistro.identificacaoOcorrencia.valor,
								linhaRegistro.motivosRejeicoes.valor,
								linhaRegistro.motivoCodigoOcorrencia19.valor,
								Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
								strObsOcorrencia,
								linhaTextoRegistroArquivo,
								ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 22 (título com pagamento cancelado)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia23EntradaTituloEmCartorio ]
		private bool b422TrataOcorrencia23EntradaTituloEmCartorio(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.atualizaBoletoItemOcorrencia23EntradaTituloEmCartorio(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 23 (entrada do título em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia23(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 23 (entrada do título em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia23(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 23 (entrada do título em cartório)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia24EntradaRejeitadaCepIrregular ]
		private bool b422TrataOcorrencia24EntradaRejeitadaCepIrregular(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia;
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados da última ocorrência na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.atualizaBoletoItemOcorrencia24CepIrregular(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o boleto com os dados da última ocorrência!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Obtém dados do registro principal do boleto ]
				rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
				#endregion

				#region [ Grava ocorrência ]
				strObsOcorrencia = "CEP inválido";
				if (!BoletoDAO.boletoOcorrenciaInsere(
										Global.Usuario.usuario,
										intNsuBoletoArqRetorno,
										_boletoCedente.id,
										rowBoletoPrincipal.id,
										idBoletoItem,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										linhaRegistro.motivoCodigoOcorrencia19.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir novo registro de ocorrência durante o tratamento da ocorrência " + linhaRegistro.identificacaoOcorrencia.valor + "!!\nNº documento: " + linhaRegistro.numeroDocumento.valor + ", nosso número: " + Global.formataBoletoNossoNumero(linhaRegistro.nossoNumeroSemDigito.valor, linhaRegistro.digitoNossoNumero.valor) + ", nº ctrl: " + linhaRegistro.numControleParticipante.valor + "\n" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia28DebitoTarifasCustas ]
		private bool b422TrataOcorrencia28DebitoTarifasCustas(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoletoPrincipal = 0;
			String strMsgErroAux = "";
			String strObsOcorrencia = "";
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Tenta identificar o boleto através do "número de controle do participante" ]
				idBoletoItem = Global.decodificaBoletoNumeroControleParticipante(linhaRegistro.numControleParticipante.valor, ref strMsgErroAux);
				if (idBoletoItem > 0)
				{
					rowBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				}
				#endregion

				#region [ Tenta identificar o boleto pelo campo "nosso número"? ]
				// Há casos em que no arquivo de retorno não é informado o campo "número de controle do participante"
				if (idBoletoItem <= 0)
				{
					if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor);
						if (rowBoletoItem != null)
						{
							idBoletoItem = rowBoletoItem.id;
						}
					}
				}
				#endregion

				#region [ Conseguiu identificar o boleto? ]
				if (idBoletoItem <= 0)
				{
					if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
					strObsOcorrencia += "Boleto não cadastrado no sistema (sem informação necessária no arquivo de retorno)";
				}
				#endregion

				#region [ Tenta recuperar dados do boleto principal ]
				if (idBoletoItem > 0)
				{
					rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
					if (rowBoletoPrincipal != null)
					{
						idBoletoPrincipal = rowBoletoPrincipal.id;
					}
				}
				#endregion

				#region [ Atualiza os dados da ocorrência na tabela t_FIN_BOLETO_ITEM ]
				if (idBoletoItem > 0)
				{
					if (!BoletoDAO.atualizaBoletoItemOcorrencia28DebitoTarifasCustas(
											Global.Usuario.usuario,
											idBoletoItem,
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											linhaRegistro,
											ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 28 (débito de tarifas/custas)!!" + strMsgErro;
						return false;
					}
				}
				#endregion

				#region [ Grava ocorrência p/ notificar usuário ]

				if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "03"))
				{
					#region [ Obs ocorrência: Tarifa de sustação (motivo 03) usando campo despesas de cobrança ]
					if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
					strObsOcorrencia += "Tarifa de sustação = " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor));
					#endregion
				}
				else if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "04"))
				{
					#region [ Obs ocorrência: Tarifa de protesto (motivo 04) usando campo despesas de cobrança ]
					if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
					strObsOcorrencia += "Tarifa de protesto = " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor));
					#endregion
				}

				#region [ Obs ocorrência: Custas de protesto (motivo 08) usando campo outras despesas ]
				if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "08"))
				{
					if (strObsOcorrencia.Length > 0) strObsOcorrencia += "; ";
					strObsOcorrencia += "Custas de protesto = " + Global.formataMoeda(Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor));
				}
				#endregion

				if (!BoletoDAO.boletoOcorrenciaInsere(
											Global.Usuario.usuario,
											intNsuBoletoArqRetorno,
											_boletoCedente.id,
											idBoletoPrincipal,
											idBoletoItem,
											linhaRegistro.numeroDocumento.valor,
											linhaRegistro.nossoNumeroSemDigito.valor,
											linhaRegistro.digitoNossoNumero.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
											Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
											linhaRegistro.identificacaoOcorrencia.valor,
											linhaRegistro.motivosRejeicoes.valor,
											linhaRegistro.motivoCodigoOcorrencia19.valor,
											Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
											strObsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir ocorrência durante o tratamento da ocorrência 28 (débito de tarifas/custas)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrencia34RetiradoCartorioManutencaoCarteira ]
		private bool b422TrataOcorrencia34RetiradoCartorioManutencaoCarteira(
								int intNsuBoletoArqRetorno,
								B422RegTipo1ArqRetorno linhaRegistro,
								String linhaTextoRegistroArquivo,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			String[] vId;
			String strIdBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui conteúdo!!";
					return false;
				}
				if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") == -1)
				{
					strMsgErro = "O campo que identifica o registro do boleto não está com o conteúdo no formato aguardado!!";
					return false;
				}
				if (intNsuBoletoArqRetorno <= 0)
				{
					strMsgErro = "Não foi fornecido o NSU para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!";
					return false;
				}
				#endregion

				#region [ Recupera Id do registro da tabela t_FIN_BOLETO_ITEM ]
				vId = linhaRegistro.numControleParticipante.valor.Split('=');
				strIdBoletoItem = vId[1];
				#endregion

				#region [ Consiste o valor do campo com o Id ]
				if (strIdBoletoItem == null)
				{
					strMsgErro = "O campo que identifica o registro do boleto não possui a informação necessária!!";
					return false;
				}

				if (strIdBoletoItem.Trim().Length == 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto não contém a informação necessária!!";
					return false;
				}

				if (Global.converteInteiro(strIdBoletoItem) <= 0)
				{
					strMsgErro = "O campo que identifica o registro do boleto possui informação inválida (" + strIdBoletoItem + ")!!";
					return false;
				}
				#endregion

				idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);

				#region [ Atualiza os dados na tabela t_FIN_BOLETO_ITEM ]
				if (!BoletoDAO.atualizaBoletoItemOcorrencia34RetiradoCartorioManutencaoCarteira(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.motivosRejeicoes.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento da ocorrência 34 (retirado de cartório e manutenção carteira)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o lançamento do fluxo de caixa ]
				if (!LancamentoFluxoCaixaDAO.atualizaLancamentoDevidoBoletoOcorrencia34(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o registro do fluxo de caixa durante o tratamento da ocorrência 34 (retirado de cartório e manutenção carteira)!!" + strMsgErro;
					return false;
				}
				#endregion

				#region [ Atualiza o histórico de pagamento dos pedidos ]
				if (!PedidoHistPagtoDAO.atualizaPagtoDevidoBoletoOcorrencia34(
										Global.Usuario.usuario,
										idBoletoItem,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao atualizar o histórico de pagamento do pedido durante o tratamento da ocorrência 34 (retirado de cartório e manutenção carteira)!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#region [ b422TrataOcorrenciaValaComum ]
		private bool b422TrataOcorrenciaValaComum(int idArqRetorno, B422RegTipo1ArqRetorno linhaRegistro, String linhaTextoRegistroArquivo, ref String strMsgErro)
		{
			#region [ Declarações ]
			int idBoletoItem = 0;
			int idBoleto = 0;
			String[] vId;
			String strIdBoletoItem;
			String strObsOcorrencia = "";
			DsDataSource.DtbFinBoletoRow rowBoletoPrincipal;
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistência ]
				if (linhaRegistro == null)
				{
					strMsgErro = "A linha do registro informada é nula!!";
					return false;
				}
				#endregion

				#region [ Possui nº controle do participante (t_FIN_BOLETO_ITEM.id)? ]
				if (linhaRegistro.numControleParticipante.valor.Trim().Length > 0)
				{
					if (linhaRegistro.numControleParticipante.valor.IndexOf(Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=") != -1)
					{
						vId = linhaRegistro.numControleParticipante.valor.Split('=');
						strIdBoletoItem = vId[1];
						if (strIdBoletoItem != null)
						{
							if (strIdBoletoItem.Trim().Length > 0)
							{
								idBoletoItem = (int)Global.converteInteiro(strIdBoletoItem);
							}
						}
					}
				}
				#endregion

				#region [ Se não possui nº controle do participante, pesquisa por 'nosso número' ]
				if (idBoletoItem == 0)
				{
					if (linhaRegistro.nossoNumeroSemDigito.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNossoNumero(_boletoCedente.id, linhaRegistro.nossoNumeroSemDigito.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor));
						if (rowBoletoItem != null) idBoletoItem = rowBoletoItem.id;
					}
				}
				#endregion

				#region [ Se não possui nº controle do participante, pesquisa por 'nº documento' ]
				if (idBoletoItem == 0)
				{
					if (linhaRegistro.numeroDocumento.valor.Trim().Length > 0)
					{
						rowBoletoItem = BoletoDAO.obtemRegistroBoletoItemByNumeroDocumento(_boletoCedente.id, linhaRegistro.numeroDocumento.valor, Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor));
						if (rowBoletoItem != null) idBoletoItem = rowBoletoItem.id;
					}
				}
				#endregion

				#region [ Se conseguiu identificar o registro do boleto, atualiza os dados ]
				if (idBoletoItem > 0)
				{
					if (!BoletoDAO.b422AtualizaBoletoItemOcorrenciaValaComum(
										Global.Usuario.usuario,
										idBoletoItem,
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										ref strMsgErro))
					{
						if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
						strMsgErro = "Falha ao atualizar os dados no BD durante o tratamento de vala comum da ocorrência " + linhaRegistro.identificacaoOcorrencia.valor + "!!" + strMsgErro;
						return false;
					}
				}
				#endregion

				#region [ Se conseguiu identificar o registro do boleto, obtém dados do registro principal ]
				if (idBoletoItem > 0)
				{
					rowBoletoPrincipal = BoletoDAO.obtemRegistroPrincipalBoletoByIdBoletoItem(idBoletoItem);
					if (rowBoletoPrincipal != null) idBoleto = rowBoletoPrincipal.id;
				}
				#endregion

				#region [ Insere registro de ocorrência na vala comum ]
				if (!BoletoDAO.b422BoletoOcorrenciaInsere(
										Global.Usuario.usuario,
										idArqRetorno,
										_boletoCedente.id,
										idBoleto,
										idBoletoItem,
										linhaRegistro.numeroDocumento.valor,
										linhaRegistro.nossoNumeroSemDigito.valor,
										linhaRegistro.digitoNossoNumero.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataVenctoTitulo.valor),
										Global.decodificaCampoMonetario(linhaRegistro.valorTitulo.valor),
										linhaRegistro.identificacaoOcorrencia.valor,
										linhaRegistro.codRejeicao.valor,
										Global.converteDdMmYyParaDateTime(linhaRegistro.dataOcorrencia.valor),
										strObsOcorrencia,
										linhaTextoRegistroArquivo,
										ref strMsgErro))
				{
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao inserir o registro no BD durante o tratamento de vala comum da ocorrência " + linhaRegistro.identificacaoOcorrencia.valor + "!!" + strMsgErro;
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return false;
			}
		}
		#endregion

		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FBoletoArqRetorno ]

		#region [ FBoletoArqRetorno_Load ]
		private void FBoletoArqRetorno_Load(object sender, EventArgs e)
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

		#region [ FBoletoArqRetorno_Shown ]
		private void FBoletoArqRetorno_Shown(object sender, EventArgs e)
		{
			try
			{
				#region[ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
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

		#region [ FBoletoArqRetorno_FormClosing ]
		private void FBoletoArqRetorno_FormClosing(object sender, FormClosingEventArgs e)
		{
			FMain.fMain.Location = this.Location;
			FMain.fMain.Visible = true;
			this.Visible = false;
		}
		#endregion

		#endregion

		#region [ btnSelecionaArqRetorno ]

		#region [ btnSelecionaArqRetorno_Click ]
		private void btnSelecionaArqRetorno_Click(object sender, EventArgs e)
		{
			trataBotaoSelecionaArqRetorno();
		}
		#endregion

		#endregion

		#region [ btnCarregaArqRetorno ]

		#region [ btnCarregaArqRetorno_Click ]
		private void btnCarregaArqRetorno_Click(object sender, EventArgs e)
		{
			trataBotaoCarregaArqRetorno();
		}
		#endregion

		#endregion

		#region [ txtArqRetorno ]

		#region [ txtArqRetorno_Enter ]
		private void txtArqRetorno_Enter(object sender, EventArgs e)
		{
			Global.textBoxSelecionaConteudo(sender);
		}
		#endregion

		#region [ txtArqRetorno_DoubleClick ]
		private void txtArqRetorno_DoubleClick(object sender, EventArgs e)
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

		#endregion
	}
}
