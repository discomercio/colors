using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ART3WebAPI.Models.Entities;
using System.Data.SqlClient;
using System.Data;
using ART3WebAPI.Models.Domains;
using System.Runtime.InteropServices;
using System.Text;

namespace ART3WebAPI.Models.Repository
{
	public class ClienteDAO
	{
		#region [ getClienteByCpfCnpj ]
		public static Cliente getClienteByCpfCnpj(string cpfCnpj)
		{
			#region [ Declarações ]
			string strSql;
			string id_cliente;
			Cliente cliente = null;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if ((cpfCnpj ?? "").Trim().Length == 0) throw new Exception("CPF/CNPJ do cliente não foi informado!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			daDataAdapter = new SqlDataAdapter();
			#endregion

			try
			{
				#region [ Monta Select ]
				strSql = "SELECT " +
							"id" +
						" FROM t_CLIENTE" +
						" WHERE" +
							" (cnpj_cpf = '" + Global.digitos(cpfCnpj) + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) return null;

				rowResultado = dtbResultado.Rows[0];
				id_cliente = BD.readToString(rowResultado["id"]);
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			cliente = getClienteById(id_cliente);
			return cliente;
		}
		#endregion

		#region [ getClienteById ]
		public static Cliente getClienteById(string id)
		{
			#region [ Declarações ]
			String strSql;
			Cliente cliente = new Cliente();
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if ((id ?? "").Trim().Length == 0) throw new Exception("ID do cliente não foi informado!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			cmCommand = new SqlCommand();
			cmCommand.Connection = cn;
			daDataAdapter = new SqlDataAdapter();
			#endregion

			try // finally: BD.fechaConexao(ref cn);
			{
				#region [ Monta Select ]
				strSql = "SELECT " +
							"*" +
						" FROM t_CLIENTE" +
						" WHERE" +
							" (id = '" + id + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) throw new Exception("Cliente com ID " + id + " não foi encontrado!!");

				#region [ Carrega os dados ]
				rowResultado = dtbResultado.Rows[0];
				cliente.id = BD.readToString(rowResultado["id"]);
				cliente.cnpj_cpf = BD.readToString(rowResultado["cnpj_cpf"]);
				cliente.tipo = BD.readToString(rowResultado["tipo"]);
				cliente.ie = BD.readToString(rowResultado["ie"]);
				cliente.rg = BD.readToString(rowResultado["rg"]);
				cliente.nome = BD.readToString(rowResultado["nome"]);
				cliente.sexo = BD.readToString(rowResultado["sexo"]);
				cliente.endereco = BD.readToString(rowResultado["endereco"]);
				cliente.bairro = BD.readToString(rowResultado["bairro"]);
				cliente.cidade = BD.readToString(rowResultado["cidade"]);
				cliente.uf = BD.readToString(rowResultado["uf"]);
				cliente.cep = BD.readToString(rowResultado["cep"]);
				cliente.ddd_res = BD.readToString(rowResultado["ddd_res"]);
				cliente.tel_res = BD.readToString(rowResultado["tel_res"]);
				cliente.ddd_com = BD.readToString(rowResultado["ddd_com"]);
				cliente.tel_com = BD.readToString(rowResultado["tel_com"]);
				cliente.ramal_com = BD.readToString(rowResultado["ramal_com"]);
				cliente.contato = BD.readToString(rowResultado["contato"]);
				cliente.dt_nasc = BD.readToDateTime(rowResultado["dt_nasc"]);
				cliente.filiacao = BD.readToString(rowResultado["filiacao"]);
				cliente.obs_crediticias = BD.readToString(rowResultado["obs_crediticias"]);
				cliente.midia = BD.readToString(rowResultado["midia"]);
				cliente.email = BD.readToString(rowResultado["email"]);
				cliente.email_opcoes = BD.readToString(rowResultado["email_opcoes"]);
				cliente.dt_cadastro = BD.readToDateTime(rowResultado["dt_cadastro"]);
				cliente.dt_ult_atualizacao = BD.readToDateTime(rowResultado["dt_ult_atualizacao"]);
				cliente.SocMaj_Nome = BD.readToString(rowResultado["SocMaj_Nome"]);
				cliente.SocMaj_CPF = BD.readToString(rowResultado["SocMaj_CPF"]);
				cliente.SocMaj_banco = BD.readToString(rowResultado["SocMaj_banco"]);
				cliente.SocMaj_agencia = BD.readToString(rowResultado["SocMaj_agencia"]);
				cliente.SocMaj_conta = BD.readToString(rowResultado["SocMaj_conta"]);
				cliente.SocMaj_ddd = BD.readToString(rowResultado["SocMaj_ddd"]);
				cliente.SocMaj_telefone = BD.readToString(rowResultado["SocMaj_telefone"]);
				cliente.SocMaj_contato = BD.readToString(rowResultado["SocMaj_contato"]);
				cliente.usuario_cadastro = BD.readToString(rowResultado["usuario_cadastro"]);
				cliente.usuario_ult_atualizacao = BD.readToString(rowResultado["usuario_ult_atualizacao"]);
				cliente.indicador = BD.readToString(rowResultado["indicador"]);
				cliente.endereco_numero = BD.readToString(rowResultado["endereco_numero"]);
				cliente.endereco_complemento = BD.readToString(rowResultado["endereco_complemento"]);
				cliente.nome_iniciais_em_maiusculas = BD.readToString(rowResultado["nome_iniciais_em_maiusculas"]);
				cliente.spc_negativado_status = BD.readToByte(rowResultado["spc_negativado_status"]);
				cliente.spc_negativado_data_negativacao = BD.readToDateTime(rowResultado["spc_negativado_data_negativacao"]);
				cliente.spc_negativado_data = BD.readToDateTime(rowResultado["spc_negativado_data"]);
				cliente.spc_negativado_data_hora = BD.readToDateTime(rowResultado["spc_negativado_data_hora"]);
				cliente.spc_negativado_usuario = BD.readToString(rowResultado["spc_negativado_usuario"]);
				cliente.email_anterior = BD.readToString(rowResultado["email_anterior"]);
				cliente.email_atualizacao_data = BD.readToDateTime(rowResultado["email_atualizacao_data"]);
				cliente.email_atualizacao_data_hora = BD.readToDateTime(rowResultado["email_atualizacao_data_hora"]);
				cliente.email_atualizacao_usuario = BD.readToString(rowResultado["email_atualizacao_usuario"]);
				cliente.contribuinte_icms_status = BD.readToByte(rowResultado["contribuinte_icms_status"]);
				cliente.contribuinte_icms_data = BD.readToDateTime(rowResultado["contribuinte_icms_data"]);
				cliente.contribuinte_icms_data_hora = BD.readToDateTime(rowResultado["contribuinte_icms_data_hora"]);
				cliente.contribuinte_icms_usuario = BD.readToString(rowResultado["contribuinte_icms_usuario"]);
				cliente.produtor_rural_status = BD.readToByte(rowResultado["produtor_rural_status"]);
				cliente.produtor_rural_data = BD.readToDateTime(rowResultado["produtor_rural_data"]);
				cliente.produtor_rural_data_hora = BD.readToDateTime(rowResultado["produtor_rural_data_hora"]);
				cliente.produtor_rural_usuario = BD.readToString(rowResultado["produtor_rural_usuario"]);
				cliente.email_xml = BD.readToString(rowResultado["email_xml"]);
				cliente.ddd_cel = BD.readToString(rowResultado["ddd_cel"]);
				cliente.tel_cel = BD.readToString(rowResultado["tel_cel"]);
				cliente.ddd_com_2 = BD.readToString(rowResultado["ddd_com_2"]);
				cliente.tel_com_2 = BD.readToString(rowResultado["tel_com_2"]);
				cliente.ramal_com_2 = BD.readToString(rowResultado["ramal_com_2"]);
				cliente.sistema_responsavel_cadastro = BD.readToInt(rowResultado["sistema_responsavel_cadastro"]);
				cliente.sistema_responsavel_atualizacao = BD.readToInt(rowResultado["sistema_responsavel_atualizacao"]);
				#endregion
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}

			return cliente;
		}
		#endregion

		#region [ insere ]
		public static bool insere(Cliente cliente, string loja, string usuario, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClienteDAO.insere()";
			string strSql;
			string msg;
			string msg_erro_aux;
			string id_cliente;
			StringBuilder sbLog;
			bool blnSucesso = false;
			int intRetorno;
			Log log;
			SqlConnection cn;
			SqlCommand cmInsert;
			SqlCommand cmUpdateTabelaControleNsuAcquireXLock;
			SqlParameter p;
			SqlTransaction trx;
			#endregion

			msg_erro = "";

			#region [ Consistências ]
			if ((cliente.cnpj_cpf ?? "").Trim().Length == 0)
			{
				msg_erro = "Não foi informado o CNPJ/CPF do cliente";
				return false;
			}

			if (!Global.isCnpjCpfOk(cliente.cnpj_cpf))
			{
				msg_erro = "CNPJ/CPF do cliente é inválido (" + cliente.cnpj_cpf + ")";
				return false;
			}
			#endregion

			#region [ Prepara acesso ao BD ]
			cn = new SqlConnection(BD.getConnectionString());
			cn.Open();
			#endregion

			try // try-catch-finally: BD.fechaConexao(ref cn);
			{
				trx = cn.BeginTransaction();
				try // try-finally: commit/rollback
				{
					#region [ Bloqueia registro p/ evitar acesso concorrente ]
					if (Global.Parametros.Geral.TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO)
					{
						// BLOQUEIA REGISTRO PARA EVITAR ACESSO CONCORRENTE (REALIZA O FLIP EM UM CAMPO BIT APENAS P/ ADQUIRIR O LOCK EXCLUSIVO)
						// OBS: TODOS OS MÓDULOS DO SISTEMA QUE REALIZEM ESTA OPERAÇÃO DE CADASTRAMENTO DEVEM SINCRONIZAR O ACESSO OBTENDO O LOCK EXCLUSIVO DO REGISTRO DE CONTROLE DESIGNADO
						strSql = "UPDATE t_CONTROLE SET" +
									" dummy = ~dummy" +
								" WHERE" +
									" (id_nsu = @id_nsu)";
						cmUpdateTabelaControleNsuAcquireXLock = new SqlCommand();
						cmUpdateTabelaControleNsuAcquireXLock.Connection = cn;
						cmUpdateTabelaControleNsuAcquireXLock.Transaction = trx;
						cmUpdateTabelaControleNsuAcquireXLock.CommandText = strSql;
						cmUpdateTabelaControleNsuAcquireXLock.Parameters.Add("@id_nsu", SqlDbType.VarChar, 80);
						cmUpdateTabelaControleNsuAcquireXLock.Prepare();
						cmUpdateTabelaControleNsuAcquireXLock.Parameters["@id_nsu"].Value = Global.Cte.ID_T_CONTROLE.ID_XLOCK_SYNC_CLIENTE;
						cmUpdateTabelaControleNsuAcquireXLock.ExecuteNonQuery();
					}
					#endregion

					if (!GeralDAO.geraNsuUsandoTabelaControle(ref cn, ref trx, Global.Cte.ID_T_CONTROLE.NSU_CADASTRO_CLIENTES, out id_cliente, out msg_erro))
					{
						if (msg_erro.Length > 0) msg_erro = "\n" + msg_erro;
						msg_erro = "Falha ao tentar gerar o identificador do registro para o cadastro de novo cliente!" +
									msg_erro;
						return false;
					}

					cliente.id = id_cliente;

					#region [ cmInsert ]
					strSql = "INSERT INTO t_CLIENTE (" +
								"id, " +
								"dt_cadastro, " +
								"usuario_cadastro, " +
								"dt_ult_atualizacao, " +
								"usuario_ult_atualizacao, " +
								"cnpj_cpf, " +
								"tipo, " +
								"ie, " +
								"produtor_rural_status, " +
								"nome, " +
								"sexo, " +
								"endereco, " +
								"endereco_numero, " +
								"endereco_complemento, " +
								"bairro, " +
								"cidade, " +
								"uf, " +
								"cep, " +
								"ddd_res, " +
								"tel_res, " +
								"ddd_cel, " +
								"tel_cel, " +
								"ddd_com, " +
								"tel_com, " +
								"ddd_com_2, " +
								"tel_com_2, " +
								"dt_nasc, " +
								"email, " +
								"sistema_responsavel_cadastro, " +
								"sistema_responsavel_atualizacao" +
							") VALUES (" +
								"@id, " +
								Global.sqlMontaGetdateSomenteData() + ", " +
								"@usuario_cadastro, " +
								"getdate(), " +
								"@usuario_ult_atualizacao, " +
								"@cnpj_cpf, " +
								"@tipo, " +
								"@ie, " +
								"@produtor_rural_status, " +
								"@nome, " +
								"@sexo, " +
								"@endereco, " +
								"@endereco_numero, " +
								"@endereco_complemento, " +
								"@bairro, " +
								"@cidade, " +
								"@uf, " +
								"@cep, " +
								"@ddd_res, " +
								"@tel_res, " +
								"@ddd_cel, " +
								"@tel_cel, " +
								"@ddd_com, " +
								"@tel_com, " +
								"@ddd_com_2, " +
								"@tel_com_2, " +
								Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_nasc") + ", " +
								"@email, " +
								"@sistema_responsavel_cadastro, " +
								"@sistema_responsavel_atualizacao" +
							")";
					cmInsert = new SqlCommand();
					cmInsert.Connection = cn;
					cmInsert.Transaction = trx;
					cmInsert.CommandText = strSql;
					cmInsert.Parameters.Add("@id", SqlDbType.VarChar, 12);
					cmInsert.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
					cmInsert.Parameters.Add("@tipo", SqlDbType.VarChar, 2);
					cmInsert.Parameters.Add("@ie", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@produtor_rural_status", SqlDbType.TinyInt);
					cmInsert.Parameters.Add("@nome", SqlDbType.VarChar, 60);
					cmInsert.Parameters.Add("@sexo", SqlDbType.VarChar, 1);
					cmInsert.Parameters.Add("@endereco", SqlDbType.VarChar, 80);
					cmInsert.Parameters.Add("@endereco_numero", SqlDbType.VarChar, 20);
					cmInsert.Parameters.Add("@endereco_complemento", SqlDbType.VarChar, 60);
					cmInsert.Parameters.Add("@bairro", SqlDbType.VarChar, 72);
					cmInsert.Parameters.Add("@cidade", SqlDbType.VarChar, 60);
					cmInsert.Parameters.Add("@uf", SqlDbType.VarChar, 2);
					cmInsert.Parameters.Add("@cep", SqlDbType.VarChar, 8);
					cmInsert.Parameters.Add("@ddd_res", SqlDbType.VarChar, 4);
					cmInsert.Parameters.Add("@tel_res", SqlDbType.VarChar, 11);
					cmInsert.Parameters.Add("@ddd_cel", SqlDbType.VarChar, 2);
					cmInsert.Parameters.Add("@tel_cel", SqlDbType.VarChar, 9);
					cmInsert.Parameters.Add("@ddd_com", SqlDbType.VarChar, 4);
					cmInsert.Parameters.Add("@tel_com", SqlDbType.VarChar, 11);
					cmInsert.Parameters.Add("@ddd_com_2", SqlDbType.VarChar, 2);
					cmInsert.Parameters.Add("@tel_com_2", SqlDbType.VarChar, 9);
					cmInsert.Parameters.Add("@dt_nasc", SqlDbType.VarChar, 10);
					cmInsert.Parameters.Add("@email", SqlDbType.VarChar, 60);
					cmInsert.Parameters.Add("@sistema_responsavel_cadastro", SqlDbType.Int);
					cmInsert.Parameters.Add("@sistema_responsavel_atualizacao", SqlDbType.Int);
					cmInsert.Prepare();
					#endregion

					#region [ Preenche o valor dos parâmetros ]
					cmInsert.Parameters["@id"].Value = cliente.id;
					p = cmInsert.Parameters["@usuario_cadastro"]; p.Value = Texto.leftStr(usuario, p.Size);
					p = cmInsert.Parameters["@usuario_ult_atualizacao"]; p.Value = Texto.leftStr(usuario, p.Size);
					cmInsert.Parameters["@cnpj_cpf"].Value = Global.digitos(cliente.cnpj_cpf);
					cmInsert.Parameters["@tipo"].Value = (cliente.tipo ?? "");
					p = cmInsert.Parameters["@ie"]; p.Value = Texto.leftStr((cliente.ie ?? ""), p.Size);
					cmInsert.Parameters["@produtor_rural_status"].Value = cliente.produtor_rural_status;
					p = cmInsert.Parameters["@nome"]; p.Value = Texto.leftStr((cliente.nome ?? ""), p.Size);
					p = cmInsert.Parameters["@sexo"]; p.Value = Texto.leftStr((cliente.sexo ?? ""), p.Size);
					p = cmInsert.Parameters["@endereco"]; p.Value = Texto.leftStr((cliente.endereco ?? ""), p.Size);
					p = cmInsert.Parameters["@endereco_numero"]; p.Value = Texto.leftStr((cliente.endereco_numero ?? ""), p.Size);
					p = cmInsert.Parameters["@endereco_complemento"]; p.Value = Texto.leftStr((cliente.endereco_complemento ?? ""), p.Size);
					p = cmInsert.Parameters["@bairro"]; p.Value = Texto.leftStr((cliente.bairro ?? ""), p.Size);
					p = cmInsert.Parameters["@cidade"]; p.Value = Texto.leftStr((cliente.cidade ?? ""), p.Size);
					p = cmInsert.Parameters["@uf"]; p.Value = Texto.leftStr((cliente.uf ?? ""), p.Size);
					p = cmInsert.Parameters["@cep"]; p.Value = Texto.leftStr(Global.digitos(cliente.cep ?? ""), p.Size);
					p = cmInsert.Parameters["@ddd_res"]; p.Value = Texto.leftStr((cliente.ddd_res ?? ""), p.Size);
					p = cmInsert.Parameters["@tel_res"]; p.Value = Texto.leftStr((cliente.tel_res ?? ""), p.Size);
					p = cmInsert.Parameters["@ddd_cel"]; p.Value = Texto.leftStr((cliente.ddd_cel ?? ""), p.Size);
					p = cmInsert.Parameters["@tel_cel"]; p.Value = Texto.leftStr((cliente.tel_cel ?? ""), p.Size);
					p = cmInsert.Parameters["@ddd_com"]; p.Value = Texto.leftStr((cliente.ddd_com ?? ""), p.Size);
					p = cmInsert.Parameters["@tel_com"]; p.Value = Texto.leftStr((cliente.tel_com ?? ""), p.Size);
					p = cmInsert.Parameters["@ddd_com_2"]; p.Value = Texto.leftStr((cliente.ddd_com_2 ?? ""), p.Size);
					p = cmInsert.Parameters["@tel_com_2"]; p.Value = Texto.leftStr((cliente.tel_com_2 ?? ""), p.Size);
					cmInsert.Parameters["@dt_nasc"].Value = Global.formataDataYyyyMmDdComSeparador(cliente.dt_nasc);
					p = cmInsert.Parameters["@email"]; p.Value = Texto.leftStr((cliente.email ?? ""), p.Size);
					cmInsert.Parameters["@sistema_responsavel_cadastro"].Value = Global.Cte.SistemaResponsavelCadastro.COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP_WEBAPI;
					cmInsert.Parameters["@sistema_responsavel_atualizacao"].Value = Global.Cte.SistemaResponsavelCadastro.COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP_WEBAPI;
					#endregion

					#region [ Monta texto para o log em arquivo ]
					// Se houver conteúdo de alguma tentativa anterior, descarta
					sbLog = new StringBuilder("");
					foreach (SqlParameter item in cmInsert.Parameters)
					{
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
					}
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = cmInsert.ExecuteNonQuery();
					}
					catch
					{
						intRetorno = 0;
					}

					if (intRetorno > 0) blnSucesso = true;
					#endregion
				}
				finally
				{
					if (blnSucesso)
					{
						trx.Commit();
					}
					else
					{
						trx.Rollback();
					}
				}

				#region [ Grava o log ]
				if (blnSucesso)
				{
					log = new Log();
					log.usuario = usuario;
					log.loja = loja;
					log.operacao = Global.Cte.LogOperacao.OP_LOG_CLIENTE_INCLUSAO;
					log.complemento = sbLog.ToString();
					LogDAO.insere(usuario, log, out msg_erro_aux);
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Cliente cadastrado com sucesso - Detalhes:\n" + sbLog.ToString());
					return true;
				}
				else
				{
					msg_erro = "Falha ao tentar inserir o registro do novo cliente!";
					msg = NOME_DESTA_ROTINA + " - " + msg_erro + "\nDetalhes:\n" + Global.serializaObjectToXml(cliente);
					Global.gravaLogAtividade(msg);
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				msg_erro = "Falha ao tentar cadastrar novo cliente: " + ex.Message;
				msg = NOME_DESTA_ROTINA + " - Exception: " + ex.ToString() + "\nDetalhes:\n" + Global.serializaObjectToXml(cliente);
				Global.gravaLogAtividade(msg);
				return false;
			}
			finally
			{
				BD.fechaConexao(ref cn);
			}
		}
		#endregion
	}
}