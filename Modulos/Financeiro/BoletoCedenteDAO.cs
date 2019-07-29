#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	class BoletoCedenteDAO
	{
		#region [ geraIndiceArqRemessaNoDia ]
		public static bool geraIndiceArqRemessaNoDia(short id_boleto_cedente, ref int nsu, ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intNsuUltimo;
			int intNsuNovo;
			int intRetorno;
			DateTime dt_indice_arq_remessa_no_dia;
			DateTime dt_hoje;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			String strSql;
			String strDescricaoLog;
			String strMsgErroAux = "";
			FinLog finLog = new FinLog();
			#endregion

			strMsgErro = "";
			nsu = 0;
			try
			{
				#region [ Prepara acesso ao BD ]
				cmCommand = BD.criaSqlCommand();
				daDataAdapter = BD.criaSqlDataAdapter();
				#endregion

				strSql = "SELECT" +
							" indice_arq_remessa_no_dia," +
							" dt_indice_arq_remessa_no_dia," +
							" getdate() AS dt_hoje" +
						" FROM t_FIN_BOLETO_CEDENTE" +
						" WHERE" +
							" (id = " + id_boleto_cedente.ToString() + ")";
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);

				if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Registro id=" + id_boleto_cedente.ToString() + " não localizado na tabela t_FIN_BOLETO_CEDENTE!!");

				rowResultado = dtbResultado.Rows[0];

				intNsuUltimo = (int)Global.converteInteiro(rowResultado["indice_arq_remessa_no_dia"].ToString());
				dt_indice_arq_remessa_no_dia = BD.readToDateTime(rowResultado["dt_indice_arq_remessa_no_dia"]);
				dt_hoje = BD.readToDateTime(rowResultado["dt_hoje"]);

				if (dt_indice_arq_remessa_no_dia.Date != dt_hoje.Date)
				{
					intNsuNovo = 1;
				}
				else
				{
					// Incrementa 1
					intNsuNovo = intNsuUltimo + 1;
				}

				// Tenta atualizar o banco de dados
				strSql = "UPDATE t_FIN_BOLETO_CEDENTE SET" +
							" indice_arq_remessa_no_dia = " + intNsuNovo.ToString() + "," +
							" dt_indice_arq_remessa_no_dia = " + Global.sqlMontaGetdateSomenteData() +
						" WHERE" +
							" (id = " + id_boleto_cedente.ToString() + ")" +
							" AND (indice_arq_remessa_no_dia = " + intNsuUltimo.ToString() + ")";
				cmCommand.CommandText = strSql;
				intRetorno = BD.executaNonQuery(ref cmCommand);
				if (intRetorno == 1)
				{
					blnSucesso = true;
					nsu = intNsuNovo;
				}
				else
				{
					throw new FinanceiroException("Falha ao incrementar a sequência diária utilizada para nomear o arquivo de remessa no registro id=" + id_boleto_cedente.ToString() + " na tabela t_FIN_BOLETO_CEDENTE!!");
				}

				// Ok
				if (blnSucesso)
				{
					#region [ Grava o log ]
					strDescricaoLog = "Gerado número de sequência diária para nomear arquivo de remessa: Cedente=" + id_boleto_cedente.ToString() + ", Nº sequencial=" + nsu.ToString();
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_GERA_INDICE_DIARIO_ARQ_REMESSA;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO_CEDENTE;
					finLog.id_registro_origem = id_boleto_cedente;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroAux);
					#endregion

					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar gerar a sequência utilizada para nomear o arquivo de remessa!!";
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

		#region [ geraNumSequencialRemessa ]
		/// <summary>
		/// Gera o "número sequencial de remessa", que é o NSU enviado no arquivo de remessa.
		/// </summary>
		/// <param name="id_boleto_cedente">
		/// Identificação do registro de t_FIN_BOLETO_CEDENTE que especifica qual é o cedente utilizado nos boletos.
		/// </param>
		/// <param name="nsu">
		/// Retorna o "número sequencial de remessa" gerado.
		/// </param>
		/// <param name="strMsgErro">
		/// Em caso de erro, retorna a descrição do erro.
		/// </param>
		/// <returns>
		/// true: sucesso na geração do NSU
		/// false: falha na geração do NSU
		/// </returns>
		public static bool geraNumSequencialRemessa(short id_boleto_cedente, ref int nsu, ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intNsuUltimo;
			int intNsuNovo;
			int intRetorno;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			String strSql;
			String strDescricaoLog;
			String strMsgErroAux = "";
			FinLog finLog = new FinLog();
			#endregion

			strMsgErro = "";
			nsu = 0;
			try
			{
				#region [ Prepara acesso ao BD ]
				cmCommand = BD.criaSqlCommand();
				daDataAdapter = BD.criaSqlDataAdapter();
				#endregion

				strSql = "SELECT" +
							" nsu_arq_remessa" +
						" FROM t_FIN_BOLETO_CEDENTE" +
						" WHERE" +
							" (id = " + id_boleto_cedente.ToString() + ")";
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);

				if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Registro id=" + id_boleto_cedente.ToString() + " não localizado na tabela t_FIN_BOLETO_CEDENTE!!");

				rowResultado = dtbResultado.Rows[0];

				intNsuUltimo = (int)Global.converteInteiro(rowResultado["nsu_arq_remessa"].ToString());

				// Incrementa 1
				intNsuNovo = intNsuUltimo + 1;

				// Tenta atualizar o banco de dados
				strSql = "UPDATE t_FIN_BOLETO_CEDENTE SET" +
							" nsu_arq_remessa = " + intNsuNovo.ToString() +
						" WHERE" +
							" (id = " + id_boleto_cedente.ToString() + ")" +
							" AND (nsu_arq_remessa = " + intNsuUltimo.ToString() + ")";
				cmCommand.CommandText = strSql;
				intRetorno = BD.executaNonQuery(ref cmCommand);
				if (intRetorno == 1)
				{
					blnSucesso = true;
					nsu = intNsuNovo;
				}
				else
				{
					throw new FinanceiroException("Falha ao incrementar o número sequencial de remessa no registro id=" + id_boleto_cedente.ToString() + " na tabela t_FIN_BOLETO_CEDENTE!!");
				}
				
				// Ok
				if (blnSucesso)
				{
					#region [ Grava o log ]
					strDescricaoLog = "Gerado número sequencial de remessa: Cedente=" + id_boleto_cedente.ToString() + ", NSU=" + nsu.ToString();
					finLog.usuario = Global.Usuario.usuario;
					finLog.operacao = Global.Cte.FIN.LogOperacao.BOLETO_GERA_NSU_ARQ_REMESSA;
					finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.MANUAL;
					finLog.fin_modulo = Global.Cte.FIN.Modulo.BOLETO;
					finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_BOLETO_CEDENTE;
					finLog.id_registro_origem = id_boleto_cedente;
					finLog.descricao = strDescricaoLog;
					FinLogDAO.insere(Global.Usuario.usuario, finLog, ref strMsgErroAux);
					#endregion

					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar gerar o número sequencial de remessa!!";
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

		#region [ getBoletoCedente ]
		/// <summary>
		/// Retorna um objeto BoletoCedente contendo os dados lidos do BD
		/// </summary>
		/// <param name="id">
		/// Identificação do registro
		/// </param>
		/// <returns>
		/// Retorna um objeto BoletoCedente contendo os dados lidos do BD
		/// </returns>
		public static BoletoCedente getBoletoCedente(int id)
		{
			#region [ Declarações ]
			String strSql;
			BoletoCedente boletoCedente = new BoletoCedente();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (id == 0) throw new FinanceiroException("O identificador do registro não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados do cedente ]
			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_BOLETO_CEDENTE" +
					" WHERE" +
						" (id = " + id.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Registro id=" + id.ToString() + " não localizado na tabela t_FIN_BOLETO_CEDENTE!!");

			rowResultado = dtbResultado.Rows[0];

			boletoCedente.id = (System.Int16)rowResultado["id"];
			boletoCedente.id_conta_corrente = (byte)rowResultado["id_conta_corrente"];
			boletoCedente.st_ativo = (byte)rowResultado["st_ativo"];
			boletoCedente.nsu_arq_remessa = (int)rowResultado["nsu_arq_remessa"];
			boletoCedente.codigo_empresa = !Convert.IsDBNull(rowResultado["codigo_empresa"]) ? rowResultado["codigo_empresa"].ToString() : "";
			boletoCedente.nome_empresa = !Convert.IsDBNull(rowResultado["nome_empresa"]) ? rowResultado["nome_empresa"].ToString() : "";
			boletoCedente.num_banco = !Convert.IsDBNull(rowResultado["num_banco"]) ? rowResultado["num_banco"].ToString() : "";
			boletoCedente.nome_banco = !Convert.IsDBNull(rowResultado["nome_banco"]) ? rowResultado["nome_banco"].ToString() : "";
			boletoCedente.agencia = !Convert.IsDBNull(rowResultado["agencia"]) ? rowResultado["agencia"].ToString() : "";
			boletoCedente.digito_agencia = !Convert.IsDBNull(rowResultado["digito_agencia"]) ? rowResultado["digito_agencia"].ToString() : "";
			boletoCedente.conta = !Convert.IsDBNull(rowResultado["conta"]) ? rowResultado["conta"].ToString() : "";
			boletoCedente.digito_conta = !Convert.IsDBNull(rowResultado["digito_conta"]) ? rowResultado["digito_conta"].ToString() : "";
			boletoCedente.carteira = !Convert.IsDBNull(rowResultado["carteira"]) ? rowResultado["carteira"].ToString() : "";
			boletoCedente.juros_mora = (Single)rowResultado["juros_mora"];
			boletoCedente.perc_multa = (Single)rowResultado["perc_multa"];
			boletoCedente.qtde_dias_protestar_apos_padrao = (byte)rowResultado["qtde_dias_protestar_apos_padrao"];
			boletoCedente.segunda_mensagem_padrao = !Convert.IsDBNull(rowResultado["segunda_mensagem_padrao"]) ? rowResultado["segunda_mensagem_padrao"].ToString() : "";
			boletoCedente.mensagem_1_padrao = !Convert.IsDBNull(rowResultado["mensagem_1_padrao"]) ? rowResultado["mensagem_1_padrao"].ToString() : "";
			boletoCedente.mensagem_2_padrao = !Convert.IsDBNull(rowResultado["mensagem_2_padrao"]) ? rowResultado["mensagem_2_padrao"].ToString() : "";
			boletoCedente.mensagem_3_padrao = !Convert.IsDBNull(rowResultado["mensagem_3_padrao"]) ? rowResultado["mensagem_3_padrao"].ToString() : "";
			boletoCedente.mensagem_4_padrao = !Convert.IsDBNull(rowResultado["mensagem_4_padrao"]) ? rowResultado["mensagem_4_padrao"].ToString() : "";
			boletoCedente.dt_cadastro = (DateTime)rowResultado["dt_cadastro"];
			boletoCedente.usuario_cadastro = !Convert.IsDBNull(rowResultado["usuario_cadastro"]) ? rowResultado["usuario_cadastro"].ToString() : "";
			boletoCedente.dt_ult_atualizacao = (DateTime)rowResultado["dt_ult_atualizacao"];
			boletoCedente.usuario_ult_atualizacao = !Convert.IsDBNull(rowResultado["usuario_ult_atualizacao"]) ? rowResultado["usuario_ult_atualizacao"].ToString() : "";
			boletoCedente.endereco = !Convert.IsDBNull(rowResultado["endereco"]) ? rowResultado["endereco"].ToString() : "";
			boletoCedente.endereco_numero = !Convert.IsDBNull(rowResultado["endereco_numero"]) ? rowResultado["endereco_numero"].ToString() : "";
			boletoCedente.endereco_complemento = !Convert.IsDBNull(rowResultado["endereco_complemento"]) ? rowResultado["endereco_complemento"].ToString() : "";
			boletoCedente.bairro = !Convert.IsDBNull(rowResultado["bairro"]) ? rowResultado["bairro"].ToString() : "";
			boletoCedente.cidade = !Convert.IsDBNull(rowResultado["cidade"]) ? rowResultado["cidade"].ToString() : "";
			boletoCedente.uf = !Convert.IsDBNull(rowResultado["uf"]) ? rowResultado["uf"].ToString() : "";
			boletoCedente.cep = !Convert.IsDBNull(rowResultado["cep"]) ? rowResultado["cep"].ToString() : "";
			boletoCedente.st_boleto_cedente_padrao = (byte)rowResultado["st_boleto_cedente_padrao"];
			boletoCedente.apelido = !Convert.IsDBNull(rowResultado["apelido"]) ? rowResultado["apelido"].ToString() : "";
			boletoCedente.loja_default_boleto_plano_contas = !Convert.IsDBNull(rowResultado["loja_default_boleto_plano_contas"]) ? rowResultado["loja_default_boleto_plano_contas"].ToString() : "";
			boletoCedente.st_participante_serasa_reciprocidade = (byte)rowResultado["st_participante_serasa_reciprocidade"];
			#endregion

			return boletoCedente;
		}
		#endregion

		#region [ getBoletoCedenteByCodigoEmpresa ]
		/// <summary>
		/// Retorna um objeto BoletoCedente contendo os dados lidos do BD
		/// </summary>
		/// <param name="codigoEmpresa">
		/// Código da empresa fornecido pelo banco
		/// </param>
		/// <returns>
		/// Retorna um objeto BoletoCedente contendo os dados lidos do BD
		/// </returns>
		public static BoletoCedente getBoletoCedenteByCodigoEmpresa(String codigoEmpresa)
		{
			#region [ Declarações ]
			String strSql;
			BoletoCedente boletoCedente = new BoletoCedente();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (codigoEmpresa == null) throw new FinanceiroException("O identificador do registro não foi fornecido!!");
			if (codigoEmpresa.Trim().Length == 0) throw new FinanceiroException("O identificador do registro não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados do cedente ]
			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_BOLETO_CEDENTE" +
					" WHERE" +
						" (CONVERT(int, codigo_empresa) = " + Global.digitos(codigoEmpresa) + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) throw new FinanceiroException("Registro com codigo_empresa=" + codigoEmpresa + " não localizado na tabela t_FIN_BOLETO_CEDENTE!!");

			rowResultado = dtbResultado.Rows[0];

			boletoCedente.id = (System.Int16)rowResultado["id"];
			boletoCedente.id_conta_corrente = (byte)rowResultado["id_conta_corrente"];
			boletoCedente.st_ativo = (byte)rowResultado["st_ativo"];
			boletoCedente.nsu_arq_remessa = (int)rowResultado["nsu_arq_remessa"];
			boletoCedente.codigo_empresa = !Convert.IsDBNull(rowResultado["codigo_empresa"]) ? rowResultado["codigo_empresa"].ToString() : "";
			boletoCedente.nome_empresa = !Convert.IsDBNull(rowResultado["nome_empresa"]) ? rowResultado["nome_empresa"].ToString() : "";
			boletoCedente.num_banco = !Convert.IsDBNull(rowResultado["num_banco"]) ? rowResultado["num_banco"].ToString() : "";
			boletoCedente.nome_banco = !Convert.IsDBNull(rowResultado["nome_banco"]) ? rowResultado["nome_banco"].ToString() : "";
			boletoCedente.agencia = !Convert.IsDBNull(rowResultado["agencia"]) ? rowResultado["agencia"].ToString() : "";
			boletoCedente.digito_agencia = !Convert.IsDBNull(rowResultado["digito_agencia"]) ? rowResultado["digito_agencia"].ToString() : "";
			boletoCedente.conta = !Convert.IsDBNull(rowResultado["conta"]) ? rowResultado["conta"].ToString() : "";
			boletoCedente.digito_conta = !Convert.IsDBNull(rowResultado["digito_conta"]) ? rowResultado["digito_conta"].ToString() : "";
			boletoCedente.carteira = !Convert.IsDBNull(rowResultado["carteira"]) ? rowResultado["carteira"].ToString() : "";
			boletoCedente.juros_mora = (Single)rowResultado["juros_mora"];
			boletoCedente.perc_multa = (Single)rowResultado["perc_multa"];
			boletoCedente.qtde_dias_protestar_apos_padrao = (byte)rowResultado["qtde_dias_protestar_apos_padrao"];
			boletoCedente.segunda_mensagem_padrao = !Convert.IsDBNull(rowResultado["segunda_mensagem_padrao"]) ? rowResultado["segunda_mensagem_padrao"].ToString() : "";
			boletoCedente.mensagem_1_padrao = !Convert.IsDBNull(rowResultado["mensagem_1_padrao"]) ? rowResultado["mensagem_1_padrao"].ToString() : "";
			boletoCedente.mensagem_2_padrao = !Convert.IsDBNull(rowResultado["mensagem_2_padrao"]) ? rowResultado["mensagem_2_padrao"].ToString() : "";
			boletoCedente.mensagem_3_padrao = !Convert.IsDBNull(rowResultado["mensagem_3_padrao"]) ? rowResultado["mensagem_3_padrao"].ToString() : "";
			boletoCedente.mensagem_4_padrao = !Convert.IsDBNull(rowResultado["mensagem_4_padrao"]) ? rowResultado["mensagem_4_padrao"].ToString() : "";
			boletoCedente.dt_cadastro = (DateTime)rowResultado["dt_cadastro"];
			boletoCedente.usuario_cadastro = !Convert.IsDBNull(rowResultado["usuario_cadastro"]) ? rowResultado["usuario_cadastro"].ToString() : "";
			boletoCedente.dt_ult_atualizacao = (DateTime)rowResultado["dt_ult_atualizacao"];
			boletoCedente.usuario_ult_atualizacao = !Convert.IsDBNull(rowResultado["usuario_ult_atualizacao"]) ? rowResultado["usuario_ult_atualizacao"].ToString() : "";
			boletoCedente.endereco = !Convert.IsDBNull(rowResultado["endereco"]) ? rowResultado["endereco"].ToString() : "";
			boletoCedente.endereco_numero = !Convert.IsDBNull(rowResultado["endereco_numero"]) ? rowResultado["endereco_numero"].ToString() : "";
			boletoCedente.endereco_complemento = !Convert.IsDBNull(rowResultado["endereco_complemento"]) ? rowResultado["endereco_complemento"].ToString() : "";
			boletoCedente.bairro = !Convert.IsDBNull(rowResultado["bairro"]) ? rowResultado["bairro"].ToString() : "";
			boletoCedente.cidade = !Convert.IsDBNull(rowResultado["cidade"]) ? rowResultado["cidade"].ToString() : "";
			boletoCedente.uf = !Convert.IsDBNull(rowResultado["uf"]) ? rowResultado["uf"].ToString() : "";
			boletoCedente.cep = !Convert.IsDBNull(rowResultado["cep"]) ? rowResultado["cep"].ToString() : "";
			boletoCedente.st_boleto_cedente_padrao = (byte)rowResultado["st_boleto_cedente_padrao"];
			boletoCedente.apelido = !Convert.IsDBNull(rowResultado["apelido"]) ? rowResultado["apelido"].ToString() : "";
			boletoCedente.loja_default_boleto_plano_contas = !Convert.IsDBNull(rowResultado["loja_default_boleto_plano_contas"]) ? rowResultado["loja_default_boleto_plano_contas"].ToString() : "";
			boletoCedente.st_participante_serasa_reciprocidade = (byte)rowResultado["st_participante_serasa_reciprocidade"];
			#endregion

			return boletoCedente;
		}
		#endregion

		#region [ getLojasBoletoCedente ]
		/// <summary>
		/// Retorna a lista das lojas que estão definidas p/ usarem o referido cedente.
		/// </summary>
		/// <param name="id_boleto_cedente">Código do cedente usado em t_FIN_BOLETO_CEDENTE.id</param>
		/// <returns>Retorna a lista das lojas que estão definidas p/ usarem o referido cedente.</returns>
		public static List<String> getLojasBoletoCedente(int id_boleto_cedente)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			List<String> listaLoja = new List<String>();
			BoletoCedente boletoCedente;
			#endregion

			boletoCedente = getBoletoCedente(id_boleto_cedente);

			#region [ Prepara objetos de acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			if (boletoCedente.st_boleto_cedente_padrao == 0)
			{
				#region [ Não é o cedente padrão ]
				strSql = "SELECT" +
							" loja" +
						" FROM t_FIN_BOLETO_CEDENTE_X_LOJA" +
						" WHERE" +
							" (id_boleto_cedente = " + id_boleto_cedente.ToString() + ")" +
							" AND (excluido_status = 0)" +
						" ORDER BY" +
							" CONVERT(smallint, loja)";
				cmCommand.CommandText = strSql;
				daDataAdapter.Fill(dtbResultado);

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					listaLoja.Add(BD.readToString(dtbResultado.Rows[i]["loja"]));
				}
				#endregion
			}
			else
			{
				#region [ É o cedente padrão ]
			//  OBTÉM TODAS AS LOJAS QUE NÃO ESTEJAM ALOCADAS P/ OS OUTROS CEDENTES
				strSql = "SELECT" +
							" loja" +
						" FROM t_LOJA" +
						" WHERE" +
							" (" +
								"CONVERT(smallint, loja) NOT IN " +
									"(" +
										"SELECT" +
											" CONVERT(smallint, loja)" +
										" FROM t_FIN_BOLETO_CEDENTE_X_LOJA" +
										" WHERE" +
											" (id_boleto_cedente <> " + id_boleto_cedente.ToString() + ")" +
											" AND (excluido_status  = 0)" +
									")" +
							")" +
						" ORDER BY" +
							" CONVERT(smallint, loja)";
				cmCommand.CommandText = strSql;
				daDataAdapter.Fill(dtbResultado);

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					listaLoja.Add(BD.readToString(dtbResultado.Rows[i]["loja"]));
				}
				#endregion
			}

			return listaLoja;
		}
		#endregion

		#region [ getNFeEmitentesBoletoCedente ]
		/// <summary>
		/// Retorna a lista de empresas emitentes de NFe que estão definidas para emitirem boletos através do referido cedente.
		/// </summary>
		/// <param name="id_boleto_cedente">Código do cedente usado em t_FIN_BOLETO_CEDENTE.id</param>
		/// <returns>Retorna a lista de empresas emitentes de NFe que estão definidas para emitirem boletos através do referido cedente.</returns>
		public static List<int> getNFeEmitentesBoletoCedente(int id_boleto_cedente)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			List<int> listaNFeEmitente = new List<int>();
			BoletoCedente boletoCedente;
			#endregion

			boletoCedente = getBoletoCedente(id_boleto_cedente);

			#region [ Prepara objetos de acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Obtém o Id das empresas emitentes de NFe vinculadas a este cedente ]
			strSql = "SELECT" +
						" id" +
					" FROM t_NFe_EMITENTE" +
					" WHERE" +
						" (id_boleto_cedente = " + id_boleto_cedente.ToString() + ")" +
					" ORDER BY" +
						" ordem";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				listaNFeEmitente.Add(BD.readToInt(dtbResultado.Rows[i]["id"]));
			}
			#endregion

			return listaNFeEmitente;
		}
		#endregion
	}
}
