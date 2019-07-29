using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace EmailSenderService
{
	class MensagemDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmAtualizaStatusDaMensagem;
		private static SqlCommand cmAtualizaStatusProcessamentoDaMensagem;
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

		#region [ Construtor Estático ]
		static MensagemDAO()
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

			#region [ cmAtualizaStatusDaMensagem ]
			strSql = "UPDATE T_EMAILSNDSVC_MENSAGEM " +
							"SET qtde_tentativas_realizadas = @qtde_tentativas_realizadas, " +
							"st_enviado_sucesso = @st_enviado_sucesso, " +
							"dt_hr_enviado_sucesso = @dt_hr_enviado_sucesso, " +
							"st_falhou_em_definitivo = @st_falhou_em_definitivo, " +
							"dt_hr_falhou_em_definitivo = @dt_hr_falhou_em_definitivo, " +
							"resultado_ult_tentativa_envio = @resultado_ult_tentativa_envio, " +
							"dt_hr_ult_tentativa_envio = @dt_hr_ult_tentativa_envio, " +
							"msg_erro_ult_tentativa_envio = @msg_erro_ult_tentativa_envio " +
							"WHERE id = @id ";

			cmAtualizaStatusDaMensagem = BD.criaSqlCommand();
			cmAtualizaStatusDaMensagem.CommandText = strSql;
			cmAtualizaStatusDaMensagem.Parameters.Add("@qtde_tentativas_realizadas", SqlDbType.Int);
			cmAtualizaStatusDaMensagem.Parameters.Add("@st_enviado_sucesso", SqlDbType.TinyInt);
			cmAtualizaStatusDaMensagem.Parameters.Add("@dt_hr_enviado_sucesso", SqlDbType.DateTime);
			cmAtualizaStatusDaMensagem.Parameters.Add("@st_falhou_em_definitivo", SqlDbType.TinyInt);
			cmAtualizaStatusDaMensagem.Parameters.Add("@dt_hr_falhou_em_definitivo", SqlDbType.DateTime);
			cmAtualizaStatusDaMensagem.Parameters.Add("@resultado_ult_tentativa_envio", SqlDbType.VarChar, 2);
			cmAtualizaStatusDaMensagem.Parameters.Add("@dt_hr_ult_tentativa_envio", SqlDbType.DateTime);
			cmAtualizaStatusDaMensagem.Parameters.Add("@msg_erro_ult_tentativa_envio", SqlDbType.VarChar, 1024);
			cmAtualizaStatusDaMensagem.Parameters.Add("@id", SqlDbType.Int);
			cmAtualizaStatusDaMensagem.Prepare();
			#endregion

			#region [ cmAtualizaStatusProcessamentoDaMensagem ]
			strSql = "UPDATE T_EMAILSNDSVC_MENSAGEM " +
							"SET st_processamento_mensagem = @st_processamento_mensagem " +
							"WHERE id = @id ";

			cmAtualizaStatusProcessamentoDaMensagem = BD.criaSqlCommand();
			cmAtualizaStatusProcessamentoDaMensagem.CommandText = strSql;
			cmAtualizaStatusProcessamentoDaMensagem.Parameters.Add("@st_processamento_mensagem", SqlDbType.TinyInt);
			cmAtualizaStatusProcessamentoDaMensagem.Parameters.Add("@id", SqlDbType.Int);
			cmAtualizaStatusProcessamentoDaMensagem.Prepare();
			#endregion
		}
		#endregion

		#region [ obtemMensagensNovasSemDataHoraAgendamento ]
		public static DataSet obtemMensagensNovasSemDataHoraAgendamento(HashSet<Int32> remetentes)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataTable dtbMensagem = new DataTable("dtbMensagem");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			StringBuilder sbMontaSQL = new StringBuilder();
			for (int i = 0; i < remetentes.Count; i++)
			{
				sbMontaSQL.Append("@p" + i);
				if (i < remetentes.Count - 1)
				{
					sbMontaSQL.Append(", ");
				}
			}

			strSql = "SELECT id, id_remetente, dt_hr_cadastro " +
							"FROM T_EMAILSNDSVC_MENSAGEM " +
							"WHERE dt_hr_agendamento_envio IS NULL " +
							"AND resultado_ult_tentativa_envio IS NULL " +
							"AND id_remetente IN ( " +
							sbMontaSQL.ToString() + " ) " +
							"AND st_envio_cancelado = 0 " +
							"AND st_processamento_mensagem = 0 " +
							"ORDER BY dt_hr_cadastro ";

			cmCommand.CommandText = strSql;

			for (int i = 0; i < remetentes.Count; i++)
			{
				cmCommand.Parameters.Add("@p" + i, SqlDbType.Int);
				cmCommand.Parameters["@p" + i].Value = remetentes.ElementAt(i);
			}

			daDataAdapter.Fill(dtbMensagem);
			dsResultado.Tables.Add(dtbMensagem);

			return dsResultado;
		}
		#endregion

		#region [ obtemMensagensNovasComDataHoraAgendamento ]
		public static DataSet obtemMensagensNovasComDataHoraAgendamento(HashSet<Int32> remetentes)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataTable dtbMensagem = new DataTable("dtbMensagem");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			StringBuilder sbMontaSQL = new StringBuilder();
			for (int i = 0; i < remetentes.Count; i++)
			{
				sbMontaSQL.Append("@p" + i);
				if (i < remetentes.Count - 1)
				{
					sbMontaSQL.Append(", ");
				}
			}

			strSql = "SELECT id, id_remetente, dt_hr_agendamento_envio " +
							"FROM T_EMAILSNDSVC_MENSAGEM " +
							"WHERE dt_hr_agendamento_envio IS NOT NULL " +
							"AND resultado_ult_tentativa_envio IS NULL " +
							"AND id_remetente IN ( " +
							sbMontaSQL.ToString() + " ) " +
							"AND st_envio_cancelado = 0 " +
							"AND st_processamento_mensagem = 0 " +
							"ORDER BY dt_hr_agendamento_envio ";

			cmCommand.CommandText = strSql;

			for (int i = 0; i < remetentes.Count; i++)
			{
				cmCommand.Parameters.Add("@p" + i, SqlDbType.Int);
				cmCommand.Parameters["@p" + i].Value = remetentes.ElementAt(i);
			}

			daDataAdapter.Fill(dtbMensagem);
			dsResultado.Tables.Add(dtbMensagem);

			return dsResultado;
		}
		#endregion

		#region [ obtemMensagensQueFalharam ]
		public static DataSet obtemMensagensQueFalharam(HashSet<Int32> remetentes)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataTable dtbMensagem = new DataTable("dtbMensagem");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			StringBuilder sbMontaSQL = new StringBuilder();
			for (int i = 0; i < remetentes.Count; i++)
			{
				sbMontaSQL.Append("@p" + i);
				if (i < remetentes.Count - 1)
				{
					sbMontaSQL.Append(", ");
				}
			}

			strSql = "SELECT id, id_remetente, qtde_tentativas_realizadas, dt_hr_ult_tentativa_envio, dt_hr_agendamento_envio, dt_hr_cadastro " +
							"FROM T_EMAILSNDSVC_MENSAGEM " +
							"WHERE resultado_ult_tentativa_envio = 'F' " +
							"AND id_remetente IN ( " +
							sbMontaSQL.ToString() + " ) " +
							"AND st_envio_cancelado = 0 " +
							"AND st_processamento_mensagem = 0 " +
							"ORDER BY id_remetente, dt_hr_agendamento_envio, dt_hr_cadastro ";

			cmCommand.CommandText = strSql;

			for (int i = 0; i < remetentes.Count; i++)
			{
				cmCommand.Parameters.Add("@p" + i, SqlDbType.Int);
				cmCommand.Parameters["@p" + i].Value = remetentes.ElementAt(i);
			}

			daDataAdapter.Fill(dtbMensagem);
			dsResultado.Tables.Add(dtbMensagem);

			return dsResultado;
		}
		#endregion

		#region [ obtemMensagensParaEnvio ]
		public static DataSet obtemMensagensParaEnvio(HashSet<Int32> idsMensagens, HashSet<Int32> idsRemetentes)
		{
			#region [ Declarações ]
			String strSql;
			StringBuilder sbMontaSQL;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataTable dtbMensagem = new DataTable("dtbMensagem");
			DataTable dtbRemetente = new DataTable("dtbRemetente");
			DataRelation drlRemetenteMensagem;
			#endregion

			//LHGX
			//por alguma razão não encontrada, existem situações em que o serviço entra nesta rotina sem existirem remetentes selecionados
			//a condição abaixo evita que um erro seja gerado e lançado no Event Viewer
			if (idsRemetentes.Count <= 0)
			{
				return dsResultado;
			}

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			//Mensagem
			sbMontaSQL = new StringBuilder();
			for (int i = 0; i < idsMensagens.Count; i++)
			{
				sbMontaSQL.Append("@p" + i);
				if (i < idsMensagens.Count - 1)
				{
					sbMontaSQL.Append(", ");
				}
			}

			strSql = "SELECT * " +
							"FROM T_EMAILSNDSVC_MENSAGEM " +
							"WHERE id IN ( " +
							sbMontaSQL.ToString() + " ) " +
							"AND st_envio_cancelado = 0 " +
							"AND st_processamento_mensagem = 0 " +
							"ORDER BY id ";

			cmCommand.CommandText = strSql;

			for (int i = 0; i < idsMensagens.Count; i++)
			{
				cmCommand.Parameters.Add("@p" + i, SqlDbType.Int);
				cmCommand.Parameters["@p" + i].Value = idsMensagens.ElementAt(i);
			}

			daDataAdapter.Fill(dtbMensagem);
			dsResultado.Tables.Add(dtbMensagem);

			//Remetente
			sbMontaSQL = new StringBuilder();
			for (int i = 0; i < idsRemetentes.Count; i++)
			{
				sbMontaSQL.Append("@x" + i);
				if (i < idsRemetentes.Count - 1)
				{
					sbMontaSQL.Append(", ");
				}
			}

			strSql = "SELECT * " +
							"FROM T_EMAILSNDSVC_REMETENTE " +
							"WHERE id IN ( " +
							sbMontaSQL.ToString() + " ) " +
							"ORDER BY id ";

			cmCommand.CommandText = strSql;

			for (int i = 0; i < idsRemetentes.Count; i++)
			{
				cmCommand.Parameters.Add("@x" + i, SqlDbType.Int);
				cmCommand.Parameters["@x" + i].Value = idsRemetentes.ElementAt(i);
			}

			daDataAdapter.Fill(dtbRemetente);
			dsResultado.Tables.Add(dtbRemetente);

			//Table Relation
			drlRemetenteMensagem = new DataRelation("dtbRemetente_dtbMensagem",
							dsResultado.Tables["dtbRemetente"].Columns["id"], dsResultado.Tables["dtbMensagem"].Columns["id_remetente"]);
			dsResultado.Relations.Add(drlRemetenteMensagem);

			return dsResultado;
		}
		#endregion

		#region [ listaEmailsDestinatariosFalhas ]
		public static String listaEmailsDestinatariosFalhas()
		{
			#region [ Declarações ]
			String strSql;
			String strListaDestinatarios = "";
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbDestinatario = new DataTable("dtbDestinatario");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT email_destinatario_falha " +
						"FROM T_EMAILSNDSVC_DESTINATARIO_FALHA " +
						"WHERE st_recebimento_habilitado = 1 ";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbDestinatario);

			foreach (DataRow row in dtbDestinatario.Rows)
			{
				if (strListaDestinatarios.Trim() != "") strListaDestinatarios = strListaDestinatarios + ";";
				strListaDestinatarios = strListaDestinatarios + BD.readToString(row["email_destinatario_falha"]);
			}

			return strListaDestinatarios;
		}
		#endregion

		#region [ atualizaStatusDaMensagem ]
		public static bool atualizaStatusDaMensagem(int qtde_tentativas_realizadas,
													int st_enviado_sucesso,
													DateTime dt_hr_enviado_sucesso,
													int st_falhou_em_definitivo,
													DateTime dt_hr_falhou_em_definitivo,
													String resultado_ult_tentativa_envio,
													DateTime dt_hr_ult_tentativa_envio,
													String msg_erro_ult_tentativa_envio,
													int id)
		{
			#region [Declarações]												
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaStatusDaMensagem.Parameters["@qtde_tentativas_realizadas"].Value = qtde_tentativas_realizadas;
			cmAtualizaStatusDaMensagem.Parameters["@st_enviado_sucesso"].Value = st_enviado_sucesso;

			if (dt_hr_enviado_sucesso == DateTime.MinValue)
			{
				cmAtualizaStatusDaMensagem.Parameters["@dt_hr_enviado_sucesso"].Value = DBNull.Value;
			}
			else
			{
				cmAtualizaStatusDaMensagem.Parameters["@dt_hr_enviado_sucesso"].Value = dt_hr_enviado_sucesso;
			}

			cmAtualizaStatusDaMensagem.Parameters["@st_falhou_em_definitivo"].Value = st_falhou_em_definitivo;

			if (dt_hr_falhou_em_definitivo == DateTime.MinValue)
			{
				cmAtualizaStatusDaMensagem.Parameters["@dt_hr_falhou_em_definitivo"].Value = DBNull.Value;
			}
			else
			{
				cmAtualizaStatusDaMensagem.Parameters["@dt_hr_falhou_em_definitivo"].Value = dt_hr_falhou_em_definitivo;
			}

			cmAtualizaStatusDaMensagem.Parameters["@resultado_ult_tentativa_envio"].Value = resultado_ult_tentativa_envio;
			cmAtualizaStatusDaMensagem.Parameters["@dt_hr_ult_tentativa_envio"].Value = dt_hr_ult_tentativa_envio;
			cmAtualizaStatusDaMensagem.Parameters["@msg_erro_ult_tentativa_envio"].Value = msg_erro_ult_tentativa_envio == null ? "" : Texto.leftStr(msg_erro_ult_tentativa_envio, 1024);
			cmAtualizaStatusDaMensagem.Parameters["@id"].Value = id;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaStatusDaMensagem);
			}
			catch (Exception)
			{
				intRetorno = 0;
			}

			if (intRetorno == 1)
			{
				blnSucesso = true;
			}
			else
			{
				blnSucesso = false;
			}
			#endregion

			return blnSucesso;
		}
		#endregion

		#region [ atualizaStatusProcessamentoMensagem ]
		public static bool atualizaStatusProcessamentoMensagem(int st_processamento_mensagem, int id)
		{
			#region [Declarações]
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaStatusProcessamentoDaMensagem.Parameters["@st_processamento_mensagem"].Value = st_processamento_mensagem;
			cmAtualizaStatusProcessamentoDaMensagem.Parameters["@id"].Value = id;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaStatusProcessamentoDaMensagem);
			}
			catch (Exception)
			{
				intRetorno = 0;
			}

			if (intRetorno == 1)
			{
				blnSucesso = true;
			}
			else
			{
				blnSucesso = false;
			}
			#endregion

			return blnSucesso;
		}
		#endregion
	}
}
