#region [ using ]
using System.Configuration;
using System.Data.SqlClient;
#endregion

namespace ExternalServiceAPI.Models.Repository
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