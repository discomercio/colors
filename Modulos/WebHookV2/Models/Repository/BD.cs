using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using WebHookV2.Models.Domains;
using WebHookV2.Models.Entities;

namespace WebHookV2.Models.Repository
{
	public class BD
	{
		#region [ Atributos ]
		private static readonly string BancoServidor = ConfigurationManager.ConnectionStrings["ServidorBanco"].ConnectionString;
		private static readonly string BancoNome = ConfigurationManager.ConnectionStrings["NomeBanco"].ConnectionString;
		private static readonly string BancoLogin = ConfigurationManager.ConnectionStrings["LoginBanco"].ConnectionString;
		private static readonly string BancoSenha = Domains.Criptografia.Descriptografa(ConfigurationManager.ConnectionStrings["SenhaBanco"].ConnectionString);
		#endregion

		#region [ getConnectionString ]
		public static string getConnectionString()
		{
			return $"{BancoServidor};{BancoNome};{BancoLogin};Password={BancoSenha}";
		}
		#endregion

		#region [ insereBraspagWebHookV2 ]
		public int insereBraspagWebHookV2(string empresa, BraspagPostNotificacao braspagPost)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BD.insereBraspagWebHookV2()";
			int intLinhasAfetadas = 0;
			string sMsg;
			#endregion

			try
			{
				using (SqlConnection conn = new SqlConnection(BD.getConnectionString()))
				{
					conn.Open();

					using (SqlCommand cmd = new SqlCommand("INSERT INTO t_BRASPAG_WEBHOOK_V2 (Empresa, RecurrentPaymentId, PaymentId, ChangeType) VALUES (@Empresa, @RecurrentPaymentId, @PaymentId, @ChangeType)", conn))
					{
						cmd.Parameters.AddWithValue("@Empresa ", (empresa ?? ""));
						cmd.Parameters.AddWithValue("@RecurrentPaymentId", (braspagPost.RecurrentPaymentId ?? ""));
						cmd.Parameters.AddWithValue("@PaymentId", (braspagPost.PaymentId ?? ""));
						cmd.Parameters.AddWithValue("@ChangeType", braspagPost.ChangeType);

						intLinhasAfetadas = cmd.ExecuteNonQuery();

						if (intLinhasAfetadas == 0)
						{
							sMsg = NOME_DESTA_ROTINA + ": Falha desconhecida ao gravar dados!\n" + braspagPost.FormataDados(showOnePropertyPerLine: false, inlinePropertySeparator: ", ");
							Global.gravaLogAtividade(sMsg);
						}
						else
						{
							sMsg = NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados!\n\t" + braspagPost.FormataDados(showOnePropertyPerLine: false, inlinePropertySeparator: ", ");
							Global.gravaLogAtividade(sMsg);
						}
					}
				}
			}
			catch (System.Exception ex)
			{
				sMsg = NOME_DESTA_ROTINA + ": Exception ao tentar gravar dados!\n" + braspagPost.FormataDados(showOnePropertyPerLine: false, inlinePropertySeparator: ", ") + "\n" + ex.ToString();
				Global.gravaLogAtividade(sMsg);
				throw new System.Exception(ex.ToString());
			}

			return intLinhasAfetadas;
		}
		#endregion
	}
}