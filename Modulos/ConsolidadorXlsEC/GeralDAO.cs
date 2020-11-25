#region [ using ]
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
#endregion

namespace ConsolidadorXlsEC
{
	public class GeralDAO
	{
		#region [ Atributos ]
		private BancoDados _bd;
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

		#region [ Construtor ]
		public GeralDAO(ref BancoDados bd)
		{
			_bd = bd;
		}
		#endregion

		#region [ Métodos ]

		#region [ getCampoDataTabelaParametro ]
		public DateTime getCampoDataTabelaParametro(String nomeParametro)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado;
			DateTime dtHrResultado = DateTime.MinValue;
			SqlCommand cmCommand;
			#endregion

			strSql = "SELECT " +
						Global.sqlMontaDateTimeParaYyyyMmDdHhMmSsComSeparador("campo_data") +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand = _bd.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				strResultado = objResultado.ToString();
				if ((strResultado != null) && (strResultado.Length > 0)) dtHrResultado = Global.converteYyyyMmDdHhMmSsParaDateTime(strResultado);
			}
			return dtHrResultado;
		}
		#endregion

		#region [ getCampoInteiroTabelaParametro ]
		public int getCampoInteiroTabelaParametro(String nomeParametro)
		{
			return getCampoInteiroTabelaParametro(nomeParametro, 0);
		}

		public int getCampoInteiroTabelaParametro(String nomeParametro, int valorDefault)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			int intResultado;
			SqlCommand cmCommand;
			#endregion

			intResultado = valorDefault;

			strSql = "SELECT " +
						"campo_inteiro" +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand = _bd.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				intResultado = BD.readToInt(objResultado);
			}
			return intResultado;
		}
		#endregion

		#region [ getCampoTextoTabelaParametro ]
		public String getCampoTextoTabelaParametro(String nomeParametro)
		{
			return getCampoTextoTabelaParametro(nomeParametro, "");
		}

		public String getCampoTextoTabelaParametro(String nomeParametro, String valorDefault)
		{
			#region [ Declarações ]
			String strSql;
			Object objResultado;
			String strResultado;
			SqlCommand cmCommand;
			#endregion

			strResultado = valorDefault;

			strSql = "SELECT " +
						"campo_texto" +
					" FROM t_PARAMETRO" +
					" WHERE" +
						" (id = '" + nomeParametro + "')";
			cmCommand = _bd.criaSqlCommand();
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (objResultado != null)
			{
				strResultado = BD.readToString(objResultado);
			}
			return strResultado;
		}
		#endregion

		#endregion
	}
}
