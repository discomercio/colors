using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

namespace EmailSenderService
{
	static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		static void Main()
		{
			string NOME_DESTA_ROTINA = "Main()";
			string strMsg;
			ServiceBase[] ServicesToRun;

			try
			{
				#region [ Log ]
				if (!Directory.Exists(Global.Cte.LogAtividade.PathLogAtividade)) Directory.CreateDirectory(Global.Cte.LogAtividade.PathLogAtividade);

				Global.gravaLogAtividade(new String('=', 80));
				Global.gravaLogAtividade(Global.Cte.Aplicativo.M_ID);
				Global.gravaLogAtividade(new String('=', 80));

				Global.gravaEventLog("Iniciando serviço '" + Global.Cte.Aplicativo.ID_SISTEMA_EMAILSENDER + "' (" + Global.Cte.Aplicativo.VERSAO + ")", EventLogEntryType.Information);
				#endregion

				ServicesToRun = new ServiceBase[]
				{ 
				//new EmailSenderService() 
				// Singleton
				EmailSenderService.getInstance()

				};
				ServiceBase.Run(ServicesToRun);
			}
			catch (Exception ex)
			{
				strMsg = ex.ToString();
				Global.gravaEventLog(NOME_DESTA_ROTINA + "\r\n" + strMsg, EventLogEntryType.Error);
			}
		}
	}
}
