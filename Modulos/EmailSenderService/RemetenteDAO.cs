using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace EmailSenderService
{
	class RemetenteDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmAtualizaStatusDoRemetente;
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
		static RemetenteDAO()
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

			#region [ cmAtualizaStatusDoRemetente ]
			strSql = "UPDATE T_EMAILSNDSVC_REMETENTE " +
							"SET resultado_ult_tentativa_envio = @resultado_ult_tentativa_envio, " +
							"dt_hr_ult_tentativa_envio = @dt_hr_ult_tentativa_envio, " +
							"ult_id_mensagem = @ult_id_mensagem " +
							"WHERE id = @id ";

			cmAtualizaStatusDoRemetente = BD.criaSqlCommand();
			cmAtualizaStatusDoRemetente.CommandText = strSql;
			cmAtualizaStatusDoRemetente.Parameters.Add("@resultado_ult_tentativa_envio", SqlDbType.VarChar, 2);
			cmAtualizaStatusDoRemetente.Parameters.Add("@dt_hr_ult_tentativa_envio", SqlDbType.DateTime);
			cmAtualizaStatusDoRemetente.Parameters.Add("@ult_id_mensagem", SqlDbType.Int);
			cmAtualizaStatusDoRemetente.Parameters.Add("@id", SqlDbType.Int);
			cmAtualizaStatusDoRemetente.Prepare();
			#endregion
		}
		#endregion

		#region [ obtemRemetentesQueNuncaMandaramEmail ]
		public static DataSet obtemRemetentesQueNuncaMandaramEmail()
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataTable dtbRemetente = new DataTable("dtbRemetente");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT id " +
						"FROM T_EMAILSNDSVC_REMETENTE " +
						"WHERE dt_hr_ult_tentativa_envio IS NULL " +
						"AND st_envio_mensagem_habilitado = 1 ";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbRemetente);
			dsResultado.Tables.Add(dtbRemetente);

			return dsResultado;
		}
		#endregion

		#region [ obtemRemetentesQueJaEnviaramEmail ]
		public static DataSet obtemRemetentesQueJaEnviaramEmail()
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataTable dtbRemetente = new DataTable("dtbRemetente");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT id, dt_hr_ult_tentativa_envio " +
							"FROM T_EMAILSNDSVC_REMETENTE " +
							"WHERE dt_hr_ult_tentativa_envio IS NOT NULL " +
							"AND st_envio_mensagem_habilitado = 1 ";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbRemetente);
			dsResultado.Tables.Add(dtbRemetente);

			return dsResultado;
		}
		#endregion

		#region [ emailRemetenteQueEnviaFalhas ]
		public static String emailRemetenteQueEnviaFalhas()
		{
			#region [ Declarações ]
			String strSql;
			String strEmailRemetente = "";
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbRemetenteFalhas = new DataTable("dtbRemetenteFalhas");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT email_remetente " +
						"FROM T_EMAILSNDSVC_REMETENTE " +
						"WHERE st_envia_falha = 1 " +
						"AND st_envio_mensagem_habilitado = 1 ";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbRemetenteFalhas);

			strEmailRemetente = BD.readToString(dtbRemetenteFalhas.Rows[0]["email_remetente"]);
			return strEmailRemetente.Trim();
		}
		#endregion

		#region [ atualizaStatusDoRemetente ]
		public static bool atualizaStatusDoRemetente(String resultado_ult_tentativa_envio,
														DateTime dt_hr_ult_tentativa_envio,
														int ult_id_mensagem,
														int id)
		{
			#region [Declarações]
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmAtualizaStatusDoRemetente.Parameters["@resultado_ult_tentativa_envio"].Value = resultado_ult_tentativa_envio;
			cmAtualizaStatusDoRemetente.Parameters["@dt_hr_ult_tentativa_envio"].Value = dt_hr_ult_tentativa_envio;
			cmAtualizaStatusDoRemetente.Parameters["@ult_id_mensagem"].Value = ult_id_mensagem;
			cmAtualizaStatusDoRemetente.Parameters["@id"].Value = id;

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmAtualizaStatusDoRemetente);
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
