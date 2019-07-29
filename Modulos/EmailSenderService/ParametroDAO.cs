using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace EmailSenderService
{
	class ParametroDAO
	{
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
		static ParametroDAO()
		{
			// NOP
		}
		#endregion

		#region [ obtemParamDeEnvioDeMensagem ]
		public static DataSet obtemParamDeEnvioDeMensagem()
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DataTable dtbParametro = new DataTable("DtbParametro");
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT id, campo_inteiro, campo_texto " +
						"FROM t_PARAMETRO " +
						"WHERE id IN ('EmailSndSvc_IntervaloMinEmSegundosEntreMsgs', " +
						"'EmailSndSvc_IntervaloEmSegundosAposCicloOcioso', " +
						"'EmailSndSvc_IntervaloMinEmSegundos_Tentativa_1_2', " +
						"'EmailSndSvc_IntervaloMinEmSegundos_Tentativa_2_3', " +
						"'EmailSndSvc_IntervaloMinEmSegundos_Tentativa_Demais', " +
						"'EmailSndSvc_QtdeMaxTentativas', " +
						"'EmailSndSvc_PeriodoSuspensao', " +
						"'EmailSndSvc_FlagHabilitacao') ";

			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbParametro);
			dsResultado.Tables.Add(dtbParametro);

			return dsResultado;
		}
		#endregion
	}
}
