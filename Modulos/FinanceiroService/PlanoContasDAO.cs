#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
#endregion

namespace FinanceiroService
{
	class PlanoContasDAO
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

		#region [ Construtor estático ]
		static PlanoContasDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
		}
		#endregion

		#region [ getPlanoContasContaById ]
		public static PlanoContasConta getPlanoContasContaById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PlanoContasDAO.getPlanoContasContaById()";
			string strSql = "";
			string msg_erro_aux;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			PlanoContasConta conta = new PlanoContasConta();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT * FROM t_FIN_PLANO_CONTAS_CONTA WHERE (id = " + id.ToString() + ")";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];

				conta.id = BD.readToInt(row["id"]);
				conta.natureza = BD.readToChar(row["natureza"]);
				conta.id_plano_contas_grupo = BD.readToShort(row["id_plano_contas_grupo"]);
				conta.st_ativo = BD.readToByte(row["st_ativo"]);
				conta.st_sistema = BD.readToByte(row["st_sistema"]);
				conta.descricao = BD.readToString(row["descricao"]);
				conta.dt_cadastro = BD.readToDateTime(row["dt_cadastro"]);
				conta.usuario_cadastro = BD.readToString(row["usuario_cadastro"]);
				conta.dt_ult_atualizacao = BD.readToDateTime(row["dt_ult_atualizacao"]);
				conta.usuario_ult_atualizacao = BD.readToString(row["usuario_ult_atualizacao"]);
				
				return conta;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "t_FIN_PLANO_CONTAS_CONTA.id=" + id.ToString();
				svcLog.complemento_2 = Global.serializaObjectToXml(conta);
				svcLog.complemento_3 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion
	}
}
