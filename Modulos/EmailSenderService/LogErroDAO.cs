using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace EmailSenderService
{
	class LogErroDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsereLogErro;
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
		static LogErroDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ inicializaObjetosEstaticos ]
		static void inicializaObjetosEstaticos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmInsereLog ]
			strSql = "INSERT INTO T_EMAILSNDSVC_LOG_ERRO ( " +
						"id, " +
						"id_mensagem, " +
						"dt_cadastro, " +
						"dt_hr_cadastro, " +
						"msg_erro " +
						") " +
						"VALUES ( " +
						"@id, " +
						"@id_mensagem, " +
						"@dt_cadastro, " +
						"@dt_hr_cadastro, " +
						"@msg_erro " +
						") ";

			cmInsereLogErro = BD.criaSqlCommand();
			cmInsereLogErro.CommandText = strSql;
			cmInsereLogErro.Parameters.Add("@id", SqlDbType.Int);
			cmInsereLogErro.Parameters.Add("@id_mensagem", SqlDbType.Int);
			cmInsereLogErro.Parameters.Add("@dt_cadastro", SqlDbType.DateTime);
			cmInsereLogErro.Parameters.Add("@dt_hr_cadastro", SqlDbType.DateTime);
			cmInsereLogErro.Parameters.Add("@msg_erro", SqlDbType.VarChar, 1024);
			cmInsereLogErro.Prepare();
			#endregion
		}
		#endregion

		#region [ insereLogErro ]
		public static bool insereLogErro(int id,
						int id_mensagem,
						DateTime dt_cadastro,
						DateTime dt_hr_cadastro,
						String msg_erro)
		{
			#region [Declarações]
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			cmInsereLogErro.Parameters["@id"].Value = id;
			cmInsereLogErro.Parameters["@id_mensagem"].Value = id_mensagem;
			cmInsereLogErro.Parameters["@dt_cadastro"].Value = dt_cadastro;
			cmInsereLogErro.Parameters["@dt_hr_cadastro"].Value = dt_hr_cadastro;
			cmInsereLogErro.Parameters["@msg_erro"].Value = msg_erro == null ? "" : Texto.leftStr(msg_erro, 1024);

			#region [ Tenta alterar o registro ]
			try
			{
				intRetorno = BD.executaNonQuery(ref cmInsereLogErro);
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
