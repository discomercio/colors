using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	public class BancoDados
	{
		#region [ Atributos ]
		public SqlConnection cnConexao;
		private SqlTransaction _sqlTransacao = null;
		private bool _transacaoEmAndamento = false;
		public bool IsConnectionOpened { get; set; } = false;
		#endregion

		#region [ Constantes ]
		public const int MAX_TAMANHO_VARCHAR = 8000;
		public const int MAX_TENTATIVAS_INSERT_BD = 3;
		public const int MAX_TENTATIVAS_UPDATE_BD = 2;
		public const int MAX_TENTATIVAS_DELETE_BD = 2;
		public const int intCommandTimeoutEmSegundos = 5 * 60;
		public const char CARACTER_CURINGA_TODOS = '%';
		#endregion

		#region [ Parâmetros de conexão ]
		public readonly string NomeAmbiente;
		public readonly string EnderecoServidor;
		public readonly string NomeBancoDados;
		public readonly string NomeUsuarioBD;
		public readonly string SenhaUsuarioCriptografada;
		#endregion

		#region [ Construtor ]
		public BancoDados(string nomeAmbiente, string enderecoServidor, string nomeBancoDados, string nomeUsuarioBD, string senhaUsuarioCriptografada)
		{
			NomeAmbiente = nomeAmbiente;
			EnderecoServidor = enderecoServidor;
			NomeBancoDados = nomeBancoDados;
			NomeUsuarioBD = nomeUsuarioBD;
			SenhaUsuarioCriptografada = senhaUsuarioCriptografada;
		}
		#endregion

		#region [ Métodos ]

		#region [ montaStringConexaoBD ]
		private string montaStringConexaoBD()
		{
			string stringConexaoBD;
			stringConexaoBD = "Data Source=" + EnderecoServidor + ";" +
								 "Initial Catalog=" + NomeBancoDados + ";" +
								 "User Id=" + NomeUsuarioBD + ";" +
								 "Password=" + Criptografia.Descriptografa(SenhaUsuarioCriptografada) + ";";
			return stringConexaoBD;
		}
		#endregion

		#region [ abreConexao ]
		public bool abreConexao(out string msgErroCompleto, out string msgErroResumido)
		{
			msgErroCompleto = "";
			msgErroResumido = "";
			try
			{
				this.cnConexao = getNovaConexao();
				return true;
			}
			catch (Exception ex)
			{
				msgErroCompleto = ex.ToString();
				msgErroResumido = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ getNovaConexao ]
		public SqlConnection getNovaConexao()
		{
			SqlConnection cn;
			string sConexao;

			sConexao = montaStringConexaoBD();
			cn = new SqlConnection(sConexao);
			cn.Open();
			
			IsConnectionOpened = true;

			return cn;
		}
		#endregion

		#region [ fechaConexao ]
		public void fechaConexao(ref SqlConnection cn)
		{
			try
			{
				IsConnectionOpened = false;
				if (cn == null) return;
				if (cn.State != ConnectionState.Closed) cn.Close();
			}
			catch (Exception)
			{
				// nop
			}
		}

		public void fechaConexao()
		{
			try
			{
				fechaConexao(ref this.cnConexao);
			}
			catch (Exception)
			{
				// nop
			}
		}
		#endregion

		#region [ criaSqlCommand ]
		public SqlCommand criaSqlCommand(ref SqlConnection cn)
		{
			SqlCommand cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			cmCommand.CommandTimeout = 0;
			cmCommand.CommandType = CommandType.Text;
			return cmCommand;
		}

		public SqlCommand criaSqlCommand()
		{
			SqlCommand cmCommand;
			cmCommand = criaSqlCommand(ref this.cnConexao);
			if (this._transacaoEmAndamento) cmCommand.Transaction = this._sqlTransacao;
			return cmCommand;
		}
		#endregion

		#region [ criaSqlDataAdapter ]
		public SqlDataAdapter criaSqlDataAdapter()
		{
			SqlDataAdapter daDataAdapter = new SqlDataAdapter();
			return daDataAdapter;
		}
		#endregion

		#region [ iniciaTransacao ]
		public void iniciaTransacao()
		{
			this._transacaoEmAndamento = true;
			this._sqlTransacao = this.cnConexao.BeginTransaction();
		}
		#endregion

		#region [ commitTransacao ]
		public void commitTransacao()
		{
			this._transacaoEmAndamento = false;
			this._sqlTransacao.Commit();
		}
		#endregion

		#region [ rollbackTransacao ]
		public void rollbackTransacao()
		{
			this._transacaoEmAndamento = false;
			this._sqlTransacao.Rollback();
		}
		#endregion

		#region [ executaNonQuery ]
		public int executaNonQuery(ref SqlCommand cmComando)
		{
			if (this._transacaoEmAndamento)
			{
				if (cmComando.Transaction != this._sqlTransacao) cmComando.Transaction = this._sqlTransacao;
			}
			return cmComando.ExecuteNonQuery();
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
		public bool geraNsuUsandoTabelaFinControle(string idNsu, out int nsu, out string strMsgErro)
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
					intRetorno = executaNonQuery(ref cmCommand);
					if (intRetorno != 1)
					{
						strMsgErro = "Falha ao criar o registro para geração de NSU!!";
						return false;
					}
				}
				#endregion
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
					intRetorno = executaNonQuery(ref cmCommand);
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
		public string gera_uid()
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
				cmCommand = criaSqlCommand();
				daDataAdapter = criaSqlDataAdapter();
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
		public VersaoModulo getVersaoModulo(string modulo, out string msgErro)
		{
			#region [ Declarações ]
			VersaoModulo versaoModulo = new VersaoModulo();
			String strSql;
			SqlCommand cmCommand;
			SqlDataReader drVersao;
			#endregion

			msgErro = "";
			try
			{
				cmCommand = criaSqlCommand();

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
						msgErro = "Módulo '" + modulo + "' não cadastrado no controle de versões do sistema!!";
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
				msgErro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ isConexaoOk ]
		public bool isConexaoOk()
		{
			#region [ Declarações ]
			DateTime dtHrServidor = DateTime.MinValue;
			#endregion

			try
			{
				if (this.cnConexao == null) return false;

				dtHrServidor = obtemDataHoraServidor();
				if (dtHrServidor != DateTime.MinValue) return true;
				return false;
			}
			catch (Exception)
			{
				return false;
			}
		}
		#endregion

		#region [ obtemDataHoraServidor ]
		public DateTime obtemDataHoraServidor()
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
				cmCommand = criaSqlCommand();
				daDataAdapter = criaSqlDataAdapter();
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
		public String readToString(object campo)
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
		public DateTime readToDateTime(object campo)
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
		public Single readToSingle(object campo)
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
		public byte readToByte(object campo)
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
		public short readToShort(object campo)
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
		public int readToInt(object campo)
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
		public System.Int16 readToInt16(object campo)
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
		public char readToChar(object campo)
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
		public decimal readToDecimal(object campo)
		{
			return (decimal)campo;
		}
		#endregion

		#endregion
	}
}
