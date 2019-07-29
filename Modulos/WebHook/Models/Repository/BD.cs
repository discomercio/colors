#region [ using ]
using System.Configuration;
using System.Data.SqlClient; 
#endregion

namespace WebHook.Models.Repository
{
	public class BD
	{
        #region [ Atributos ]
        private static string BancoServidor = ConfigurationManager.ConnectionStrings["ServidorBancoDados"].ConnectionString;
        private static string BancoNome = ConfigurationManager.ConnectionStrings["NomeBancoDados"].ConnectionString;
        private static string BancoLogin = ConfigurationManager.ConnectionStrings["LoginBancoDados"].ConnectionString;
        private static string BancoSenha = Domains.Criptografia.Descriptografa(ConfigurationManager.ConnectionStrings["SenhaBancoDados"].ConnectionString);
        #endregion

        #region [ getConnectionString ]
        public static string getConnectionString()
        {
            return $"{BancoServidor};{BancoNome};{BancoLogin};Password={BancoSenha}";
        } 
        #endregion

        #region [ Grava (t_BRASPAG_WEBHOOK) ]
        public static int Grava(BraspagParameters parameters, string empresa)
        {
            int intLinhasAfetadas = 0;

            try
            {
                using (SqlConnection conn = new SqlConnection(BD.getConnectionString()))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand("insert into t_BRASPAG_WEBHOOK (Empresa, NumPedido, Status, CODPAGAMENTO) values (@empresa, @pedido, @status, @codpagamento)", conn))
                    {
                        cmd.Parameters.AddWithValue("@empresa", empresa);
                        cmd.Parameters.AddWithValue("@pedido", parameters.NumPedido);
                        cmd.Parameters.AddWithValue("@status", parameters.Status);
                        cmd.Parameters.AddWithValue("@codpagamento", parameters.CODPAGAMENTO);

                        intLinhasAfetadas = cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception(ex.ToString());
            }

            return intLinhasAfetadas;
        } 
        #endregion

        #region [ DataHora ]
        public static string DataHora()
        {
            string strRetorno;

            using (SqlConnection conn = new SqlConnection(getConnectionString()))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand("SELECT getdate()", conn))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            strRetorno = reader.GetDateTime(0).ToString();
                        }
                        else
                        {
                            strRetorno = "";
                        }
                    }
                }
            }

            return strRetorno;
        } 
        #endregion
    }
}