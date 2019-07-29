#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
#endregion

namespace Financeiro
{
	class SerasaDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmSerasaClienteInsert;
		private static SqlCommand cmSerasaTituloMovimentoInsert;
		#endregion

		#region [ inicializaConstrutorEstatico ]
		public static void inicializaConstrutorEstatico()
		{
			// NOP
			// 1) The static constructor for a class executes before any instance of the class is created.
			// 2) The static constructor for a class executes before any of the static members for the class are referenced.
			// 3) The static constructor for a class executes after the static field initializers (if any) for the class.
			// 4) The static constructor for a class executes at most one time during a single program instantiation
			// 5) A static constructor does not take access modifiers or have parameters.
			// 6) A static constructor is called automatically to initialize the class before the first instance is created or any static members are referenced.
			// 7) A static constructor cannot be called directly.
			// 8) The user has no control on when the static constructor is executed in the program.
			// 9) A typical use of static constructors is when the class is using a log file and the constructor is used to write entries to this file.
		}
		#endregion

		#region [ Construtor estático ]
		static SerasaDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmSerasaClienteInsert ]
			strSql = "INSERT INTO t_SERASA_CLIENTE (" +
						"id, " +
						"id_cliente, " +
						"cnpj, " +
						"dt_cliente_desde" +
					") VALUES (" +
						"@id, " +
						"@id_cliente, " +
						"@cnpj, " +
						"@dt_cliente_desde" +
					")";
			cmSerasaClienteInsert = BD.criaSqlCommand();
			cmSerasaClienteInsert.CommandText = strSql;
			cmSerasaClienteInsert.Parameters.Add("@id", SqlDbType.Int);
			cmSerasaClienteInsert.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
			cmSerasaClienteInsert.Parameters.Add("@cnpj", SqlDbType.VarChar, 14);
			cmSerasaClienteInsert.Parameters.Add("@dt_cliente_desde", SqlDbType.VarChar, 10);
			cmSerasaClienteInsert.Prepare();
			#endregion

			#region [ cmSerasaTituloMovimentoInsert ]
			strSql = "INSERT INTO t_SERASA_TITULO_MOVIMENTO (" +
						"id, " +
						"id_boleto_arq_retorno, " +
						"id_boleto_item, " +
						"id_serasa_cliente, " +
						"cnpj, " +
						"identificacao_ocorrencia_boleto, " +
						"numero_documento, " +
						"nosso_numero, " +
						"digito_nosso_numero, " +
						"dt_emissao, " +
						"vl_titulo, " +
						"dt_vencto, " +
						"dt_pagto, " +
						"vl_pago" +
					") VALUES (" +
						"@id, " +
						"@id_boleto_arq_retorno, " +
						"@id_boleto_item, " +
						"@id_serasa_cliente, " +
						"@cnpj, " +
						"@identificacao_ocorrencia_boleto, " +
						"@numero_documento, " +
						"@nosso_numero, " +
						"@digito_nosso_numero, " +
						"@dt_emissao, " +
						"@vl_titulo, " +
						"@dt_vencto, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_pagto") + ", " +
						"@vl_pago" +
					")";
			cmSerasaTituloMovimentoInsert = BD.criaSqlCommand();
			cmSerasaTituloMovimentoInsert.CommandText = strSql;
			cmSerasaTituloMovimentoInsert.Parameters.Add("@id", SqlDbType.Int);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@id_boleto_arq_retorno", SqlDbType.Int);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@id_boleto_item", SqlDbType.Int);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@id_serasa_cliente", SqlDbType.Int);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@cnpj", SqlDbType.VarChar, 14);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@identificacao_ocorrencia_boleto", SqlDbType.VarChar, 2);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@numero_documento", SqlDbType.VarChar, 10);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@nosso_numero", SqlDbType.VarChar, 11);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@digito_nosso_numero", SqlDbType.VarChar, 1);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@dt_emissao", SqlDbType.VarChar, 10);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@vl_titulo", SqlDbType.Money);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 10);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@dt_pagto", SqlDbType.VarChar, 10);
			cmSerasaTituloMovimentoInsert.Parameters.Add("@vl_pago", SqlDbType.Money);
			cmSerasaTituloMovimentoInsert.Prepare();
			#endregion
		}
		#endregion

		#region [ clienteInsere ]
		public static bool clienteInsere(String usuario,
										SerasaCliente cliente,
										out String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			bool blnGerouNsu;
			int intQtdeTentativas = 0;
			int intNsu = 0;
			int intRetorno;
			String strDescricaoLog;
			String strOperacao = "Gravação de registro de cliente em t_SERASA_CLIENTE";
			StringBuilder sbLog = new StringBuilder("");
			FinLog finLog = new FinLog();
			#endregion

			try
			{
				#region [ Laço de tentativas de inserção no banco de dados ]
				do
				{
					intQtdeTentativas++;

					strMsgErro = "";
					blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_SERASA_CLIENTE, ref intNsu, ref strMsgErro);

					#region [ Se gerou o NSU, tenta gravar o registro ]
					if (blnGerouNsu)
					{
						#region [ Preenche o valor dos parâmetros ]
						cmSerasaClienteInsert.Parameters["@id"].Value = intNsu;
						cmSerasaClienteInsert.Parameters["@id_cliente"].Value = cliente.id_cliente;
						cmSerasaClienteInsert.Parameters["@cnpj"].Value = Global.digitos(cliente.cnpj);
						cmSerasaClienteInsert.Parameters["@dt_cliente_desde"].Value = Global.formataDataYyyyMmDdComSeparador(cliente.dt_cliente_desde);
						#endregion

						#region [ Monta texto para o log em arquivo ]
						// Se houver conteúdo de alguma tentativa anterior, descarta
						sbLog = new StringBuilder("");
						foreach (SqlParameter item in cmSerasaClienteInsert.Parameters)
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
						}
						#endregion

						#region [ Tenta inserir o registro ]
						try
						{
							intRetorno = BD.executaNonQuery(ref cmSerasaClienteInsert);
						}
						catch (Exception ex)
						{
							intRetorno = 0;
							Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
						}
						#endregion

						#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
						if (intRetorno == 1)
						{
							cliente.id = intNsu;
							blnSucesso = true;
						}
						else
						{
							Thread.Sleep(100);
						}
						#endregion
					}
					#endregion

				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
				#endregion

				#region [ Grava log ]
				if (blnSucesso)
				{
					if (sbLog.Length > 0)
					{
						strDescricaoLog = "Inserção do registro em t_SERASA_CLIENTE.id=" + intNsu.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.SERASA_CLIENTE_INSERE;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.SERASA_RECIPROCIDADE;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_SERASA_CLIENTE;
						finLog.id_registro_origem = intNsu;
						finLog.id_cliente = cliente.id_cliente;
						finLog.cnpj_cpf = cliente.cnpj;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar gravar no banco de dados o registro do cliente em t_SERASA_CLIENTE após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ tituloMovimentoInsere ]
		public static bool tituloMovimentoInsere(String usuario,
												SerasaTituloMovimento tituloMovimento,
												out String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			bool blnGerouNsu;
			int intQtdeTentativas = 0;
			int intNsu = 0;
			int intRetorno;
			String strDescricaoLog;
			String strOperacao = "Gravação de registro em t_SERASA_TITULO_MOVIMENTO";
			StringBuilder sbLog = new StringBuilder("");
			FinLog finLog = new FinLog();
			SerasaCliente serasaCliente;
			#endregion

			try
			{
				#region [ Laço de tentativas de inserção no banco de dados ]
				do
				{
					intQtdeTentativas++;

					strMsgErro = "";
					blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_SERASA_TITULO_MOVIMENTO, ref intNsu, ref strMsgErro);

					#region [ Se gerou o NSU, tenta gravar o registro ]
					if (blnGerouNsu)
					{
						#region [ Preenche o valor dos parâmetros ]
						cmSerasaTituloMovimentoInsert.Parameters["@id"].Value = intNsu;
						cmSerasaTituloMovimentoInsert.Parameters["@id_boleto_arq_retorno"].Value = tituloMovimento.id_boleto_arq_retorno;
						cmSerasaTituloMovimentoInsert.Parameters["@id_boleto_item"].Value = tituloMovimento.id_boleto_item;
						cmSerasaTituloMovimentoInsert.Parameters["@id_serasa_cliente"].Value = tituloMovimento.id_serasa_cliente;
						cmSerasaTituloMovimentoInsert.Parameters["@cnpj"].Value = tituloMovimento.cnpj;
						cmSerasaTituloMovimentoInsert.Parameters["@identificacao_ocorrencia_boleto"].Value = tituloMovimento.identificacao_ocorrencia_boleto;
						cmSerasaTituloMovimentoInsert.Parameters["@numero_documento"].Value = tituloMovimento.numero_documento;
						cmSerasaTituloMovimentoInsert.Parameters["@nosso_numero"].Value = tituloMovimento.nosso_numero;
						cmSerasaTituloMovimentoInsert.Parameters["@digito_nosso_numero"].Value = tituloMovimento.digito_nosso_numero;
						cmSerasaTituloMovimentoInsert.Parameters["@dt_emissao"].Value = Global.formataDataYyyyMmDdComSeparador(tituloMovimento.dt_emissao);
						cmSerasaTituloMovimentoInsert.Parameters["@vl_titulo"].Value = tituloMovimento.vl_titulo;
						cmSerasaTituloMovimentoInsert.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(tituloMovimento.dt_vencto);
						cmSerasaTituloMovimentoInsert.Parameters["@dt_pagto"].Value = Global.formataDataYyyyMmDdComSeparador(tituloMovimento.dt_pagto);
						cmSerasaTituloMovimentoInsert.Parameters["@vl_pago"].Value = tituloMovimento.vl_pago;
						#endregion

						#region [ Monta texto para o log em arquivo ]
						// Se houver conteúdo de alguma tentativa anterior, descarta
						sbLog = new StringBuilder("");
						foreach (SqlParameter item in cmSerasaTituloMovimentoInsert.Parameters)
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
						}
						#endregion

						#region [ Tenta inserir o registro ]
						try
						{
							intRetorno = BD.executaNonQuery(ref cmSerasaTituloMovimentoInsert);
						}
						catch (Exception ex)
						{
							intRetorno = 0;
							Global.gravaLogAtividade(strOperacao + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: " + sbLog.ToString() + "\n" + ex.ToString());
						}
						#endregion

						#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
						if (intRetorno == 1)
						{
							tituloMovimento.id = intNsu;
							blnSucesso = true;
						}
						else
						{
							Thread.Sleep(100);
						}
						#endregion
					}
					#endregion

				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
				#endregion

				#region [ Grava log ]
				if (blnSucesso)
				{
					if (sbLog.Length > 0)
					{
						serasaCliente = getSerasaClienteById(tituloMovimento.id_serasa_cliente);
						strDescricaoLog = "Inserção do registro em t_SERASA_TITULO_MOVIMENTO.id=" + intNsu.ToString() + ": " + sbLog.ToString();
						Global.gravaLogAtividade(strDescricaoLog);
						finLog.usuario = usuario;
						finLog.operacao = Global.Cte.FIN.LogOperacao.SERASA_TITULO_MOVIMENTO_INSERE;
						finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
						finLog.fin_modulo = Global.Cte.FIN.Modulo.SERASA_RECIPROCIDADE;
						finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_SERASA_TITULO_MOVIMENTO;
						finLog.id_registro_origem = intNsu;
						finLog.id_cliente = serasaCliente.id_cliente;
						finLog.cnpj_cpf = serasaCliente.cnpj;
						finLog.descricao = strDescricaoLog;
						FinLogDAO.insere(usuario, finLog, ref strMsgErro);
					}
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar gravar no banco de dados o registro em t_SERASA_TITULO_MOVIMENTO " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ clienteCarregaFromDataRow ]
		private static SerasaCliente clienteCarregaFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			SerasaCliente cliente = new SerasaCliente();
			#endregion

			cliente.id = BD.readToInt(rowDados["id"]);
			cliente.id_cliente = BD.readToString(rowDados["id_cliente"]);
			cliente.cnpj = BD.readToString(rowDados["cnpj"]);
			cliente.raiz_cnpj = BD.readToString(rowDados["raiz_cnpj"]);
			cliente.dt_cliente_desde = BD.readToDateTime(rowDados["dt_cliente_desde"]);
			cliente.st_enviado_serasa = BD.readToByte(rowDados["st_enviado_serasa"]);
			cliente.dt_enviado_serasa = BD.readToDateTime(rowDados["dt_enviado_serasa"]);
			cliente.id_serasa_arq_remessa_normal = BD.readToInt(rowDados["id_serasa_arq_remessa_normal"]);

			return cliente;
		}
		#endregion

		#region [ tituloMovimentoCarregaFromDataRow ]
		private static SerasaTituloMovimento tituloMovimentoCarregaFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			SerasaTituloMovimento tituloMovimento = new SerasaTituloMovimento();
			#endregion

			tituloMovimento.id = BD.readToInt(rowDados["id"]);
			tituloMovimento.id_boleto_arq_retorno = BD.readToInt(rowDados["id_boleto_arq_retorno"]);
			tituloMovimento.id_boleto_item = BD.readToInt(rowDados["id_boleto_item"]);
			tituloMovimento.id_serasa_cliente = BD.readToInt(rowDados["id_serasa_cliente"]);
			tituloMovimento.cnpj = BD.readToString(rowDados["cnpj"]);
			tituloMovimento.raiz_cnpj = BD.readToString(rowDados["raiz_cnpj"]);
			tituloMovimento.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			tituloMovimento.dt_hr_cadastro = BD.readToDateTime(rowDados["dt_hr_cadastro"]);
			tituloMovimento.identificacao_ocorrencia_boleto = BD.readToString(rowDados["identificacao_ocorrencia_boleto"]);
			tituloMovimento.st_envio_serasa_cancelado = BD.readToByte(rowDados["st_envio_serasa_cancelado"]);
			tituloMovimento.dt_envio_serasa_cancelado = BD.readToDateTime(rowDados["dt_envio_serasa_cancelado"]);
			tituloMovimento.dt_hr_envio_serasa_cancelado = BD.readToDateTime(rowDados["dt_hr_envio_serasa_cancelado"]);
			tituloMovimento.usuario_envio_serasa_cancelado = BD.readToString(rowDados["usuario_envio_serasa_cancelado"]);
			tituloMovimento.st_enviado_serasa = BD.readToByte(rowDados["st_enviado_serasa"]);
			tituloMovimento.id_serasa_arq_remessa_normal = BD.readToInt(rowDados["id_serasa_arq_remessa_normal"]);
			tituloMovimento.st_retorno_serasa = BD.readToByte(rowDados["st_retorno_serasa"]);
			tituloMovimento.id_serasa_arq_retorno_normal = BD.readToInt(rowDados["id_serasa_arq_retorno_normal"]);
			tituloMovimento.st_processado_serasa_sucesso = BD.readToByte(rowDados["st_processado_serasa_sucesso"]);
			tituloMovimento.st_editado_manual = BD.readToByte(rowDados["st_editado_manual"]);
			tituloMovimento.dt_editado_manual = BD.readToDateTime(rowDados["dt_editado_manual"]);
			tituloMovimento.dt_hr_editado_manual = BD.readToDateTime(rowDados["dt_hr_editado_manual"]);
			tituloMovimento.usuario_editado_manual = BD.readToString(rowDados["usuario_editado_manual"]);
			tituloMovimento.qtde_vezes_editado_manual = BD.readToInt(rowDados["qtde_vezes_editado_manual"]);
			tituloMovimento.numero_documento = BD.readToString(rowDados["numero_documento"]);
			tituloMovimento.nosso_numero = BD.readToString(rowDados["nosso_numero"]);
			tituloMovimento.digito_nosso_numero = BD.readToString(rowDados["digito_nosso_numero"]);
			tituloMovimento.dt_emissao = BD.readToDateTime(rowDados["dt_emissao"]);
			tituloMovimento.vl_titulo = BD.readToDecimal(rowDados["vl_titulo"]);
			tituloMovimento.dt_vencto = BD.readToDateTime(rowDados["dt_vencto"]);
			tituloMovimento.dt_pagto = BD.readToDateTime(rowDados["dt_pagto"]);
			tituloMovimento.vl_pago = BD.readToDecimal(rowDados["vl_pago"]);
			tituloMovimento.retorno_codigos_erro = BD.readToString(rowDados["retorno_codigos_erro"]);

			return tituloMovimento;
		}
		#endregion

		#region [ getSerasaClienteById ]
		/// <summary>
		/// Localiza e retorna o registro de t_SERASA_CLIENTE a partir do campo "id" que é a chave primária de t_SERASA_CLIENTE
		/// </summary>
		/// <param name="id">Identificação do cliente definida em t_SERASA_CLIENTE.id</param>
		/// <returns>Objeto SerasaCliente com os dados armazenados em t_SERASA_CLIENTE</returns>
		public static SerasaCliente getSerasaClienteById(int id)
		{
			#region [ Declarações ]
			String strSql;
			SerasaCliente cliente = new SerasaCliente();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (id == 0) throw new FinanceiroException("O ID do registro do cliente não foi fornecido!!");
			if (id < 0) throw new FinanceiroException("O ID do registro do cliente não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados do cliente ]
			strSql = "SELECT " +
						"*" +
					" FROM t_SERASA_CLIENTE" +
					" WHERE" +
						" (id = " + id.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			#region [ Cliente não encontrado ]
			if (dtbResultado.Rows.Count == 0) return null;
			#endregion

			rowResultado = dtbResultado.Rows[0];

			cliente = clienteCarregaFromDataRow(rowResultado);
			#endregion

			return cliente;
		}
		#endregion

		#region [ getSerasaClienteByIdCliente ]
		/// <summary>
		/// Localiza e retorna o registro de t_SERASA_CLIENTE a partir do campo "id_cliente" que é a chave estrangeira para t_CLIENTE.id
		/// </summary>
		/// <param name="id_cliente">Identificação do cliente conforme definido em t_CLIENTE.id</param>
		/// <returns>Objeto SerasaCliente com os dados armazenados em t_SERASA_CLIENTE</returns>
		public static SerasaCliente getSerasaClienteByIdCliente(string id_cliente)
		{
			#region [ Declarações ]
			String strSql;
			SerasaCliente cliente = new SerasaCliente();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (id_cliente == null) throw new FinanceiroException("A identificação do cliente não foi fornecida!!");
			if (id_cliente.Length == 0) throw new FinanceiroException("A identificação do cliente não foi informada!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados do cliente ]
			strSql = "SELECT " +
						"*" +
					" FROM t_SERASA_CLIENTE" +
					" WHERE" +
						" (id_cliente = '" + id_cliente + "')";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			#region [ Cliente não encontrado ]
			if (dtbResultado.Rows.Count == 0) return null;
			#endregion

			rowResultado = dtbResultado.Rows[0];

			cliente = clienteCarregaFromDataRow(rowResultado);
			#endregion

			return cliente;
		}
		#endregion

		#region [ getSerasaClienteByRaizCnpj ]
		/// <summary>
		/// Localiza e retorna o registro de t_SERASA_CLIENTE a partir da raiz do CNPJ
		/// </summary>
		/// <param name="Cnpj_ou_RaizCnpj">Pode ser informado o CNPJ completo ou apenas a raiz do CNPJ do cliente</param>
		/// <returns>Objeto SerasaCliente com os dados armazenados em t_SERASA_CLIENTE</returns>
		public static SerasaCliente getSerasaClienteByRaizCnpj(String Cnpj_ou_RaizCnpj)
		{
			#region [ Declarações ]
			String strSql;
			String strRaizCnpj;
			SerasaCliente cliente = new SerasaCliente();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (Cnpj_ou_RaizCnpj == null) throw new FinanceiroException("O CNPJ do cliente não foi fornecido!!");
			if (Global.digitos(Cnpj_ou_RaizCnpj).Length == 0) throw new FinanceiroException("O CNPJ do cliente não foi informado!!");
			strRaizCnpj = Texto.leftStr(Global.digitos(Cnpj_ou_RaizCnpj), Global.Cte.Etc.TAMANHO_RAIZ_CNPJ);
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Dados do cliente ]
			strSql = "SELECT " +
						"*" +
					" FROM t_SERASA_CLIENTE" +
					" WHERE" +
						" (raiz_cnpj = '" + Global.digitos(strRaizCnpj) + "')";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			#region [ Cliente não encontrado ]
			if (dtbResultado.Rows.Count == 0) return null;
			#endregion

			rowResultado = dtbResultado.Rows[0];

			cliente = clienteCarregaFromDataRow(rowResultado);
			#endregion

			return cliente;
		}
		#endregion

		#region [ getDataClienteDesde ]
		public static DateTime getDataClienteDesde(String cnpj)
		{
			#region [ Declarações ]
			String strSql;
			DateTime dt_cliente_desde;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (cnpj == null) throw new FinanceiroException("O CNPJ do cliente não foi fornecido!!");
			if (Global.digitos(cnpj).Length == 0) throw new FinanceiroException("O CNPJ do cliente não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Pesquisa o pedido mais antigo do cliente ]
			strSql = "SELECT" +
						" Min(data) AS dt_cliente_desde" +
					" FROM t_PEDIDO tP" +
						" INNER JOIN t_CLIENTE tC ON (tP.id_cliente = tC.id)" +
					" WHERE" +
						" (tC.cnpj_cpf = '" + Global.digitos(cnpj) + "')";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			#region [ Nenhum pedido encontrado ]
			if (dtbResultado.Rows.Count == 0) return DateTime.MinValue;
			#endregion

			rowResultado = dtbResultado.Rows[0];

			dt_cliente_desde = BD.readToDateTime(rowResultado["dt_cliente_desde"]);
			#endregion

			return dt_cliente_desde;
		}
		#endregion

		#region [ getTituloMovimentoById ]
		public static SerasaTituloMovimento getTituloMovimentoById(int id)
		{
			#region [ Declarações ]
			String strSql;
			SerasaTituloMovimento tituloMovimento;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (id == 0) throw new FinanceiroException("O ID do registro do título não foi fornecido!!");
			if (id < 0) throw new FinanceiroException("O ID do registro do título não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Executa a consulta ]
			strSql = "SELECT " +
						"*" +
					" FROM t_SERASA_TITULO_MOVIMENTO" +
					" WHERE" +
						" (id = " + id.ToString() + ")" +
					" ORDER BY" +
						" id";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			#region [ Nenhum registro encontrado ]
			if (dtbResultado.Rows.Count == 0) return null;
			#endregion

			rowResultado = dtbResultado.Rows[0];
			tituloMovimento = tituloMovimentoCarregaFromDataRow(rowResultado);
			#endregion

			return tituloMovimento;
		}
		#endregion

		#region [ getTituloMovimentoByIdBoletoItem ]
		public static List<SerasaTituloMovimento> getTituloMovimentoByIdBoletoItem(int id_boleto_item)
		{
			#region [ Declarações ]
			String strSql;
			SerasaTituloMovimento tituloMovimento;
			List<SerasaTituloMovimento> listaTituloMovimento = new List<SerasaTituloMovimento>();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (id_boleto_item == 0) throw new FinanceiroException("O identificador do boleto (t_FIN_BOLETO_ITEM.id) não foi fornecido!!");
			if (id_boleto_item < 0) throw new FinanceiroException("O identificador do boleto (t_FIN_BOLETO_ITEM.id) não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Executa consulta ]
			strSql = "SELECT " +
						"*" +
					" FROM t_SERASA_TITULO_MOVIMENTO" +
					" WHERE" +
						" (id_boleto_item = " + id_boleto_item.ToString() + ")" +
					" ORDER BY" +
						" id";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) return listaTituloMovimento;

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];
				tituloMovimento = tituloMovimentoCarregaFromDataRow(rowResultado);
				listaTituloMovimento.Add(tituloMovimento);
			}
			#endregion

			return listaTituloMovimento;
		}
		#endregion

		#region [ getTituloMovimentoByNumeroDocumento ]
		public static List<SerasaTituloMovimento> getTituloMovimentoByNumeroDocumento(String numeroDocumento)
		{
			#region [ Declarações ]
			String strSql;
			SerasaTituloMovimento tituloMovimento;
			List<SerasaTituloMovimento> listaTituloMovimento = new List<SerasaTituloMovimento>();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (numeroDocumento == null) throw new FinanceiroException("O número do documento do título não foi fornecido!!");
			if (numeroDocumento.Length == 0) throw new FinanceiroException("O número do documento do título não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Executa a consulta ]
			strSql = "SELECT " +
						"*" +
					" FROM t_SERASA_TITULO_MOVIMENTO" +
					" WHERE" +
						" (numero_documento = '" + numeroDocumento + "')" +
					" ORDER BY" +
						" id";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) return listaTituloMovimento;

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];
				tituloMovimento = tituloMovimentoCarregaFromDataRow(rowResultado);
				listaTituloMovimento.Add(tituloMovimento);
			}
			#endregion

			return listaTituloMovimento;
		}
		#endregion

		#region [ getTituloMovimentoByNossoNumero ]
		public static List<SerasaTituloMovimento> getTituloMovimentoByNossoNumero(String nossoNumero)
		{
			#region [ Declarações ]
			String strSql;
			SerasaTituloMovimento tituloMovimento;
			List<SerasaTituloMovimento> listaTituloMovimento = new List<SerasaTituloMovimento>();
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (nossoNumero == null) throw new FinanceiroException("O nosso número do título não foi fornecido!!");
			if (nossoNumero.Length == 0) throw new FinanceiroException("O nosso número do título não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Executa a consulta ]
			strSql = "SELECT " +
						"*" +
					" FROM t_SERASA_TITULO_MOVIMENTO" +
					" WHERE" +
						" (nosso_numero = '" + nossoNumero + "')" +
					" ORDER BY" +
						" id";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);

			if (dtbResultado.Rows.Count == 0) return listaTituloMovimento;

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];
				tituloMovimento = tituloMovimentoCarregaFromDataRow(rowResultado);
				listaTituloMovimento.Add(tituloMovimento);
			}
			#endregion

			return listaTituloMovimento;
		}
		#endregion
	}
}
