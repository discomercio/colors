#region[ using ]
using System;
using System.Data.SqlClient;
using System.Data;
using System.Threading;
using System.Configuration;
#endregion

namespace FinanceiroService
{
	static class BD
	{
		#region [ Atributos ]
		public static SqlConnection cnConexao;
		private static SqlTransaction _sqlTransacao;
		private static bool _transacaoEmAndamento;
		#endregion

		#region[ Constantes ]

		#region [ Parâmetros de conexão ]
		public static readonly string strServidor = ConfigurationManager.ConnectionStrings["ServidorBanco"].ConnectionString;
		public static readonly string strNomeBancoDados = ConfigurationManager.ConnectionStrings["NomeBanco"].ConnectionString;
		public static readonly string strNomeUsuario = ConfigurationManager.ConnectionStrings["LoginBanco"].ConnectionString;
		private static readonly string strSenhaUsuarioCriptografada = ConfigurationManager.ConnectionStrings["SenhaBanco"].ConnectionString;
		public static readonly string strDescricaoAmbiente = Global.GetConfigurationValue("AmbienteExecucao");
		#endregion

		public const int MAX_TAMANHO_VARCHAR = 8000;
		public const int MAX_TENTATIVAS_INSERT_BD = 3;
		public const int MAX_TENTATIVAS_UPDATE_BD = 2;
		public const int MAX_TENTATIVAS_DELETE_BD = 2;
		public const int intCommandTimeoutEmSegundos = 5 * 60;
		public const char CARACTER_CURINGA_TODOS = '%';
		#endregion

		#region[ Métodos ]

		#region[ montaStringConexaoBd ]
		private static String montaStringConexaoBd()
		{
			String strStringConexaoBd;
			strStringConexaoBd = "Data Source=" + strServidor + ";" +
								 "Initial Catalog=" + strNomeBancoDados + ";" +
								 "User Id=" + strNomeUsuario + ";" +
								 "Password=" + Criptografia.Descriptografa(strSenhaUsuarioCriptografada) + ";";
			return strStringConexaoBd;
		}
		#endregion

		#region [ abreConexao ]
		public static void abreConexao()
		{
			BD.cnConexao = abreNovaConexao();
		}
		#endregion

		#region [ abreNovaConexao ]
		public static SqlConnection abreNovaConexao()
		{
			SqlConnection cn;
			String strConnection;

			strConnection = montaStringConexaoBd();
			cn = new SqlConnection(strConnection);
			cn.Open();

			return cn;
		}
		#endregion

		#region [ fechaConexao ]
		public static void fechaConexao()
		{
			try
			{
				fechaConexao(ref cnConexao);
			}
			catch (Exception)
			{
				// Nop
			}
		}

		public static void fechaConexao(ref SqlConnection cn)
		{
			try
			{
				if (cn == null) return;
				if (cn.State != ConnectionState.Closed) cn.Close();
			}
			catch (Exception)
			{
				// Nop
			}
		}
		#endregion

		#region [ criaSqlCommand ]
		public static SqlCommand criaSqlCommand()
		{
			SqlCommand cmCommand;
			cmCommand = criaSqlCommand(ref cnConexao);
			if (_transacaoEmAndamento) cmCommand.Transaction = _sqlTransacao;
			return cmCommand;
		}

		public static SqlCommand criaSqlCommand(ref SqlConnection cn)
		{
			SqlCommand cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			cmCommand.CommandTimeout = 0;
			cmCommand.CommandType = CommandType.Text;
			return cmCommand;
		}
		#endregion

		#region [ criaSqlDataAdapter ]
		public static SqlDataAdapter criaSqlDataAdapter()
		{
			SqlDataAdapter daDataAdapter = new SqlDataAdapter();
			return daDataAdapter;
		}
		#endregion

		#region [ iniciaTransacao ]
		public static void iniciaTransacao()
		{
			_transacaoEmAndamento = true;
			_sqlTransacao = cnConexao.BeginTransaction();
		}
		#endregion

		#region [ commitTransacao ]
		public static void commitTransacao()
		{
			_transacaoEmAndamento = false;
			_sqlTransacao.Commit();
		}
		#endregion

		#region [ rollbackTransacao ]
		public static void rollbackTransacao()
		{
			_transacaoEmAndamento = false;
			_sqlTransacao.Rollback();
		}
		#endregion

		#region [ executaNonQuery ]
		public static int executaNonQuery(ref SqlCommand cmComando)
		{
			if (_transacaoEmAndamento)
			{
				if (cmComando.Transaction != _sqlTransacao) cmComando.Transaction = _sqlTransacao;
			}
			return cmComando.ExecuteNonQuery();
		}
		#endregion

		#region [ executeScalar ]
		public static object executeScalar(ref SqlCommand cmComando)
		{
			if (_transacaoEmAndamento)
			{
				if (cmComando.Transaction != _sqlTransacao) cmComando.Transaction = _sqlTransacao;
			}
			return cmComando.ExecuteScalar();
		}
		#endregion

		#region [ geraNsuUsandoTabelaFinControle ]
		/// <summary>
		/// Gera o NSU para a chave informada
		/// </summary>
		/// <param name="idNsu">
		/// Identificação da chave para gerar o NSU, normalmente é o próprio nome da tabela para a qual se deseja gerar o NSU para se usar como ID
		/// </param>
		/// <param name="nsu">
		/// Retorna o NSU gerado
		/// </param>
		/// <param name="strMsgErro">
		/// Retorna a mensagem de erro em caso de exception
		/// </param>
		/// <returns>
		/// true: sucesso ao gerar o NSU
		/// false: falha ao gerar o NSU
		/// </returns>
		public static bool geraNsuUsandoTabelaFinControle(String idNsu, out int nsu, out String strMsgErro)
		{
			#region [ Declarações ]
			const int MAX_TENTATIVAS = 10;
			int intQtdeTentativas = 0;
			bool blnSucesso = false;
			int intRetorno;
			int intNsuUltimo;
			int intNsuNovo;
			SqlCommand cmCommand;
			String strSql;
			#endregion

			strMsgErro = "";
			nsu = 0;
			try
			{
				cmCommand = criaSqlCommand();

				#region [ Verifica se registro existe, senão cria agora ]
				strSql = "SELECT" +
							" Count(*) AS qtde" +
						" FROM t_FIN_CONTROLE" +
						" WHERE" +
							" (id='" + idNsu + "')";
				cmCommand.CommandText = strSql;
				intRetorno = (int)cmCommand.ExecuteScalar();

				#region [ Não está cadastrado, então cadastra agora ]
				if (intRetorno == 0)
				{
					strSql = "INSERT INTO t_FIN_CONTROLE (" +
								"id, " +
								"nsu, " +
								"dt_hr_ult_atualizacao" +
							") VALUES (" +
								"'" + idNsu + "'," +
								"0," +
								"getdate()" +
							")";
					cmCommand.CommandText = strSql;
					intRetorno = BD.executaNonQuery(ref cmCommand);
					if (intRetorno != 1)
					{
						strMsgErro = "Falha ao criar o registro para geração de NSU!!";
						return false;
					}
				}
				#endregion
				#endregion

				#region [ Bloqueia registro p/ evitar acesso concorrente ]
				if (Global.Parametros.Geral.TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO)
				{
					strSql = "UPDATE t_FIN_CONTROLE SET" +
								" dummy = ~dummy" +
							" WHERE" +
								" (id='" + idNsu + "')";
					cmCommand.CommandText = strSql;
					BD.executaNonQuery(ref cmCommand);
				}
				#endregion

				#region [ Laço de tentativas para gerar o NSU (devido a acesso concorrente ]
				do
				{
					intQtdeTentativas++;

					// Obtém o último NSU usado
					strSql = "SELECT" +
								" nsu" +
							" FROM t_FIN_CONTROLE" +
							" WHERE" +
								" id = '" + idNsu + "'";
					cmCommand.CommandText = strSql;
					intNsuUltimo = (int)cmCommand.ExecuteScalar();

					// Incrementa 1
					intNsuNovo = intNsuUltimo + 1;

					// Tenta atualizar o banco de dados
					strSql = "UPDATE t_FIN_CONTROLE SET" +
								" nsu = " + intNsuNovo.ToString() + ", " +
								" dt_hr_ult_atualizacao = getdate()" +
							" WHERE" +
								" (id = '" + idNsu + "')" +
								" AND (nsu = " + intNsuUltimo.ToString() + ")";
					cmCommand.CommandText = strSql;
					intRetorno = BD.executaNonQuery(ref cmCommand);
					if (intRetorno == 1)
					{
						blnSucesso = true;
						nsu = intNsuNovo;
					}
					else
					{
						Thread.Sleep(500);
					}
				} while ((!blnSucesso) && (intQtdeTentativas < MAX_TENTATIVAS));
				#endregion

				// Ok
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar gerar o NSU!!";
					return false;
				}
			}
			catch (Exception ex)
			{
				strMsgErro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ gera_uid ]
		public static string gera_uid()
		{
			#region [ Declarações ]
			string strUID = "";
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			try
			{
				#region [ Prepara objetos de acesso ao BD ]
				cmCommand = BD.criaSqlCommand();
				daDataAdapter = BD.criaSqlDataAdapter();
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT Convert(varchar(36), NEWID()) AS uid";
				cmCommand.CommandText = strSql;
				daDataAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count > 0)
				{
					strUID = readToString(dtbResultado.Rows[0]["uid"]);
				}

				return strUID;
			}
			catch (Exception)
			{
				return "";
			}
		}
		#endregion

		#region [ getVersaoModulo ]
		public static VersaoModulo getVersaoModulo(string modulo, out string strMsgErro)
		{
			#region [ Declarações ]
			VersaoModulo versaoModulo = new VersaoModulo();
			String strSql;
			SqlCommand cmCommand;
			SqlDataReader drVersao;
			#endregion

			strMsgErro = "";
			try
			{
				cmCommand = BD.criaSqlCommand();

				strSql = "SELECT " +
							"*" +
						" FROM t_VERSAO" +
						" WHERE" +
							" (modulo = '" + modulo + "')";
				cmCommand.CommandText = strSql;
				drVersao = cmCommand.ExecuteReader();
				try
				{
					if (drVersao.Read())
					{
						versaoModulo.modulo = readToString(drVersao["modulo"]);
						versaoModulo.versao = readToString(drVersao["versao"]);
						versaoModulo.mensagem = readToString(drVersao["mensagem"]);
						versaoModulo.cor_fundo_padrao = readToString(drVersao["cor_fundo_padrao"]);
						versaoModulo.identificador_ambiente = readToString(drVersao["identificador_ambiente"]);
						return versaoModulo;
					}
					else
					{
						strMsgErro = "Módulo '" + modulo + "' não cadastrado no controle de versões do sistema!!";
						return null;
					}
				}
				finally
				{
					drVersao.Close();
				}
			}
			catch (Exception ex)
			{
				strMsgErro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ isConexaoOk ]
		public static bool isConexaoOk()
		{
			#region [ Declarações ]
			string sParam;
			#endregion

			try
			{
				if (BD.cnConexao == null) return false;

				sParam = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.PARAMETRO_DUMMY_TESTE_CONEXAO_BD);
				if ((sParam ?? "").Length > 0) return true;
				return false;
			}
			catch (Exception)
			{
				return false;
			}
		}
		#endregion

		#region [ obtemDataHoraServidor ]
		public static DateTime obtemDataHoraServidor()
		{
			#region [ Declarações ]
			DateTime dataHoraResposta = DateTime.MinValue;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			try
			{
				#region [ Prepara objetos de acesso ao BD ]
				cmCommand = BD.criaSqlCommand();
				daDataAdapter = BD.criaSqlDataAdapter();
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT getdate() AS data_hora";
				cmCommand.CommandText = strSql;
				daDataAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count > 0)
				{
					dataHoraResposta = readToDateTime(dtbResultado.Rows[0]["data_hora"]);
				}

				return dataHoraResposta;
			}
			catch (Exception)
			{
				return DateTime.MinValue;
			}
		}
		#endregion

		#region [ readToString ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo texto
		/// </param>
		/// <returns>
		/// Retorna o texto armazenado no campo. Caso o conteúdo seja DBNull, retorna uma String vazia.
		/// </returns>
		public static String readToString(object campo)
		{
			return !Convert.IsDBNull(campo) ? campo.ToString() : "";
		}
		#endregion

		#region [ readToDateTime ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo data
		/// </param>
		/// <returns>
		/// Retorna a data armazenada no campo. Caso o conteúdo seja DBNull, retorna DateTime.MinValue
		/// </returns>
		public static DateTime readToDateTime(object campo)
		{
			return !Convert.IsDBNull(campo) ? (DateTime)campo : DateTime.MinValue;
		}
		#endregion

		#region [ readToSingle ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo real
		/// </param>
		/// <returns>
		/// Retorna o número real armazenado no campo
		/// </returns>
		public static Single readToSingle(object campo)
		{
			return (Single)campo;
		}
		#endregion

		#region [ readToByte ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo byte
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static byte readToByte(object campo)
		{
			return !Convert.IsDBNull(campo) ? (byte)campo : (byte)0;
		}
		#endregion

		#region [ readToShort ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo short
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static short readToShort(object campo)
		{
			return (short)campo;
		}
		#endregion

		#region [ readToInt ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo int
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static int readToInt(object campo)
		{
			if (campo.GetType().Name.Equals("Int16"))
			{
				return (int)(Int16)campo;
			}
			else
			{
				return (int)campo;
			}
		}
		#endregion

		#region [ readToInt16 ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo System.Int16
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static System.Int16 readToInt16(object campo)
		{
			return (System.Int16)campo;
		}
		#endregion

		#region [ readToChar ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja do tipo char
		/// </param>
		/// <returns>
		/// Retorna o caracter armazenado no campo. Caso o conteúdo seja DBNull, retorna um caracter nulo.
		/// </returns>
		public static char readToChar(object campo)
		{
			String s;
			char c = '\0';

			if (!Convert.IsDBNull(campo))
			{
				s = campo.ToString();
				if (s.Length > 0) c = s[0];
			}

			return c;
		}
		#endregion

		#region [ readToDecimal ]
		/// <summary>
		/// O parâmetro informado deve ser uma coluna de um DataRow, ou seja, um campo lido do DataRow
		/// </summary>
		/// <param name="campo">
		/// Coluna de um DataRow, ou seja, um campo lido do DataRow cujo conteúdo seja compatível com o tipo decimal
		/// </param>
		/// <returns>
		/// Retorna o número armazenado no campo
		/// </returns>
		public static decimal readToDecimal(object campo)
		{
			return (decimal)campo;
		}
		#endregion

		#endregion
	}

	#region [ Classe VersaoModulo ]
	public class VersaoModulo
	{
		private string _modulo;
		public string modulo
		{
			get { return _modulo; }
			set { _modulo = value; }
		}

		private string _versao;
		public string versao
		{
			get { return _versao; }
			set { _versao = value; }
		}

		private string _mensagem;
		public string mensagem
		{
			get { return _mensagem; }
			set { _mensagem = value; }
		}

		private string _cor_fundo_padrao;
		public string cor_fundo_padrao
		{
			get { return _cor_fundo_padrao; }
			set { _cor_fundo_padrao = value; }
		}

		private string _identificador_ambiente;
		public string identificador_ambiente
		{
			get { return _identificador_ambiente; }
			set { _identificador_ambiente = value; }
		}
	}
	#endregion
}
