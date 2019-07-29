using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace FinanceiroService
{
	class EmailCtrlDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsertEmailCtrl;
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
		static EmailCtrlDAO()
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

			#region [ cmInsertEmailCtrl ]
			strSql = "INSERT INTO t_PAGTO_GW_EMAIL_CTRL (" +
						"id, " +
						"id_emailsndsvc_mensagem, " +
						"pedido, " +
						"id_cliente, " +
						"cnpj_cpf_cliente, " +
						"tipo_destinatario, " +
						"modulo, " +
						"tipo_msg, " +
						"codigo_msg, " +
						"rotina, " +
						"remetente, " +
						"destinatario" +
					") VALUES (" +
						"@id, " +
						"@id_emailsndsvc_mensagem, " +
						"@pedido, " +
						"@id_cliente, " +
						"@cnpj_cpf_cliente, " +
						"@tipo_destinatario, " +
						"@modulo, " +
						"@tipo_msg, " +
						"@codigo_msg, " +
						"@rotina, " +
						"@remetente, " +
						"@destinatario" +
					")";
			cmInsertEmailCtrl = BD.criaSqlCommand();
			cmInsertEmailCtrl.CommandText = strSql;
			cmInsertEmailCtrl.Parameters.Add("@id", SqlDbType.Int);
			cmInsertEmailCtrl.Parameters.Add("@id_emailsndsvc_mensagem", SqlDbType.Int);
			cmInsertEmailCtrl.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertEmailCtrl.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
			cmInsertEmailCtrl.Parameters.Add("@cnpj_cpf_cliente", SqlDbType.VarChar, 14);
			cmInsertEmailCtrl.Parameters.Add("@tipo_destinatario", SqlDbType.VarChar, 1);
			cmInsertEmailCtrl.Parameters.Add("@modulo", SqlDbType.VarChar, 2);
			cmInsertEmailCtrl.Parameters.Add("@tipo_msg", SqlDbType.VarChar, 1);
			cmInsertEmailCtrl.Parameters.Add("@codigo_msg", SqlDbType.VarChar, 6);
			cmInsertEmailCtrl.Parameters.Add("@rotina", SqlDbType.VarChar, 160);
			cmInsertEmailCtrl.Parameters.Add("@remetente", SqlDbType.VarChar, 160);
			cmInsertEmailCtrl.Parameters.Add("@destinatario", SqlDbType.VarChar, 160);
			cmInsertEmailCtrl.Prepare();
			#endregion
		}
		#endregion

		#region [ getLastEmailCtrlByPedidoFilteredByCodigoMsg ]
		public static EmailCtrl getLastEmailCtrlByPedidoFilteredByCodigoMsg(string numeroPedido, Global.Cte.EmailCtrl.CodigoMsg codigoMsg, out string msg_erro)
		{
			#region [ Declarações ]
			String strSql;
			EmailCtrl emailCtrl;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Prepara acesso ao BD ]
				cmCommand = BD.criaSqlCommand();
				daDataAdapter = BD.criaSqlDataAdapter();
				#endregion

				#region [ Inicialização ]
				numeroPedido = numeroPedido.Trim();
				#endregion

				#region [ Monta Select ]
				strSql = "SELECT " +
							"*" +
						" FROM t_PAGTO_GW_EMAIL_CTRL" +
						" WHERE" +
							" (pedido = '" + numeroPedido + "')";

				if (codigoMsg != null)
				{
					strSql += " AND (codigo_msg = '" + codigoMsg + "')";
				}

				strSql += " ORDER BY" +
							" data_hora DESC";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daDataAdapter.SelectCommand = cmCommand;
				daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count > 0)
				{
					rowResultado = dtbResultado.Rows[0];
					emailCtrl = new EmailCtrl();
					emailCtrl.id = BD.readToInt(rowResultado["id"]);
					emailCtrl.id_emailsndsvc_mensagem = BD.readToInt(rowResultado["id_emailsndsvc_mensagem"]);
					emailCtrl.data = BD.readToDateTime(rowResultado["data"]);
					emailCtrl.data_hora = BD.readToDateTime(rowResultado["data_hora"]);
					emailCtrl.pedido = BD.readToString(rowResultado["pedido"]);
					emailCtrl.id_cliente = BD.readToString(rowResultado["id_cliente"]);
					emailCtrl.cnpj_cpf_cliente = BD.readToString(rowResultado["cnpj_cpf_cliente"]);
					emailCtrl.tipo_destinatario = BD.readToString(rowResultado["tipo_destinatario"]);
					emailCtrl.modulo = BD.readToString(rowResultado["modulo"]);
					emailCtrl.tipo_msg = BD.readToString(rowResultado["tipo_msg"]);
					emailCtrl.codigo_msg = BD.readToString(rowResultado["codigo_msg"]);
					emailCtrl.rotina = BD.readToString(rowResultado["rotina"]);
					emailCtrl.remetente = BD.readToString(rowResultado["remetente"]);
					emailCtrl.destinatario = BD.readToString(rowResultado["destinatario"]);
					return emailCtrl;
				}

				return null;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getLastEmailCtrlByPedido ]
		public static EmailCtrl getLastEmailCtrlByPedido(string numeroPedido, out string msg_erro)
		{
			msg_erro = "";
			return getLastEmailCtrlByPedidoFilteredByCodigoMsg(numeroPedido, null, out msg_erro);
		}
		#endregion

		#region [ insere ]
		public static bool insere(EmailCtrl emailCtrl, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "EmailCtrlDAO.insere()";
			bool blnGerouNsu;
			int idEmailCtrl = 0;
			int intRetorno;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Gera NSU ]
				if (emailCtrl.id == 0)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_EMAIL_CTRL, out idEmailCtrl, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro de controle de envio de email!!\n" + msg_erro;
						return false;
					}
					emailCtrl.id = idEmailCtrl;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					idEmailCtrl = emailCtrl.id;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmInsertEmailCtrl.Parameters["@id"].Value = emailCtrl.id;
				cmInsertEmailCtrl.Parameters["@id_emailsndsvc_mensagem"].Value = emailCtrl.id_emailsndsvc_mensagem;
				cmInsertEmailCtrl.Parameters["@pedido"].Value = emailCtrl.pedido;
				cmInsertEmailCtrl.Parameters["@id_cliente"].Value = emailCtrl.id_cliente;
				cmInsertEmailCtrl.Parameters["@cnpj_cpf_cliente"].Value = emailCtrl.cnpj_cpf_cliente;
				cmInsertEmailCtrl.Parameters["@tipo_destinatario"].Value = emailCtrl.tipo_destinatario;
				cmInsertEmailCtrl.Parameters["@modulo"].Value = emailCtrl.modulo;
				cmInsertEmailCtrl.Parameters["@tipo_msg"].Value = emailCtrl.tipo_msg;
				cmInsertEmailCtrl.Parameters["@codigo_msg"].Value = emailCtrl.codigo_msg;
				cmInsertEmailCtrl.Parameters["@rotina"].Value = emailCtrl.rotina;
				cmInsertEmailCtrl.Parameters["@remetente"].Value = emailCtrl.remetente;
				cmInsertEmailCtrl.Parameters["@destinatario"].Value = emailCtrl.destinatario;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertEmailCtrl);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					msg_erro = ex.Message;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Exception:\n" + ex.ToString());
				}
				#endregion

				#region [ Gravou o registro? ]
				if (intRetorno == 0)
				{
					msg_erro = "Falha ao tentar gravar o registro de controle de envio de email!!\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_PAGTO_GW_EMAIL_CTRL.id=" + emailCtrl.id.ToString() + ")");

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion
	}
}
