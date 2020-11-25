#region[ using ]
using System;
using System.Collections.Generic;
using System.Text;
using System.EnterpriseServices;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
#endregion

[assembly: ApplicationName("ComPlusWrapper_DllInscE32")]
[assembly: Description("Wrapper para a DllInscE32.dll")]
namespace ComPlusWrapper_DllInscE32
{
    [JustInTimeActivation(true)]
    [Transaction(TransactionOption.Required)]
	[GuidAttribute("1584E6B9-3E0F-46D6-A6EC-568B2A6094D4")]
	public class ComPlusWrapper_DllInscE32 : ServicedComponent
    {
        #region [ Declarações ]

        #region [ Declaração da chamada à API DllInscE32.dll do Sintegra ]
        [DllImport(@"DllInscE32.dll", EntryPoint = "ConsisteInscricaoEstadual")]
        private static extern Int32 DllImport_ConsisteInscricaoEstadual(string Insc, string UF);
        #endregion

        #endregion

        #region[ Rotinas Públicas ]

        #region [ Construtor ]
        public ComPlusWrapper_DllInscE32()
        {
            // NOP
            /*
             * ATENÇÃO: lembre-se que o código executado no construtor é executado
             * ======== sempre que um novo objeto instanciar este serviço de componente.
             * No caso de inicializar variáveis globais, é fundamental ter em mente
             * que as variáveis do tipo 'static' são únicas, independentemente da instância.
             * Ou seja, uma instância interfere na outra no caso das variáveis globais do
             * tipo 'static', por isso é necessário muito cuidado com elas.
             */
        }
        #endregion

        #region[ Versao ]
        public string Versao()
        {
            return Global.Cte.Versao.strVersao;
        }
        #endregion

        #region[ DataHora ]
        public string DataHora()
        {
            return DateTime.Now.ToString(Global.Cte.DataHora.FmtDdMmYyyyHhMmSsComSeparador);
        }
        #endregion

        #region [ ConsisteInscricaoEstadual ]
        public Int32 ConsisteInscricaoEstadual(String Insc, String UF)
        {
            #region [ Declarações ]
            Int32 intRetorno;
            #endregion

            try
            {
                Global.rwlDllInscE32.AcquireWriterLock(20 * 1000);
                try
                {
                    intRetorno = DllImport_ConsisteInscricaoEstadual(Insc, UF);
                    return intRetorno;
                }
                finally
                {
                    Global.rwlDllInscE32.ReleaseWriterLock();
                }
            }
            catch (Exception ex)
            {
                Global.GravaEventLog(Global.Cte.Versao.strNomeSistema, ex.ToString(), EventLogEntryType.Error);
                throw new Exception(ex.ToString());
            }
        }
        #endregion

        #region [ isInscricaoEstadualOk ]
        public bool isInscricaoEstadualOk(String Insc, String UF)
        {
            #region [ Declarações ]
            Int32 intRetorno;
            #endregion

            try
            {
                Global.rwlDllInscE32.AcquireWriterLock(20 * 1000);
                try
                {
                    intRetorno = DllImport_ConsisteInscricaoEstadual(Insc, UF);
                    if (intRetorno == 0)
                        return true;
                    else
                        return false;
                }
                finally
                {
                    Global.rwlDllInscE32.ReleaseWriterLock();
                }
            }
            catch (Exception ex)
            {
                Global.GravaEventLog(Global.Cte.Versao.strNomeSistema, ex.ToString(), EventLogEntryType.Error);
                throw new Exception(ex.ToString());
            }
        }
        #endregion

        #endregion
    }
}
