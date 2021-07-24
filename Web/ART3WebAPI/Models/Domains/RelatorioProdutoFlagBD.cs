using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace ART3WebAPI.Models.Domains
{
    public class RelatorioProdutoFlagBD
    {
        public static Task RelatorioProdutoFlagPostAsync(string paginaId, string usuario, string codFabricante, string codProduto, short flag)
        {
            #region [ Declarações ]
            string strSql;
            int intRetorno = 0;
            SqlConnection cn;
            SqlCommand cmdUpdate;
            SqlCommand cmdInsere;
            #endregion

            #region [ Action ]
            return Task.Run(() =>
                {
					try
					{
						using (cn = new SqlConnection(Repository.BD.getConnectionString()))
						{
							cn.Open();
							try // Finally: cn.Close()
							{
								SqlParameter[] parameters = {
									new SqlParameter("@flag", SqlDbType.TinyInt),
									new SqlParameter("@usuario", SqlDbType.VarChar, 20),
									new SqlParameter("@codFabricante", SqlDbType.VarChar, 4),
									new SqlParameter("@codProduto", SqlDbType.VarChar, 8),
									new SqlParameter("@paginaId", SqlDbType.Int)
										};

								parameters[0].Value = flag;
								parameters[1].Value = usuario;
								parameters[2].Value = codFabricante;
								parameters[3].Value = codProduto;
								parameters[4].Value = int.Parse(paginaId);

								strSql = "UPDATE t_RELATORIO_PRODUTO_FLAG SET Flag = @flag, DataHoraUltimaAtualizacao = getdate() WHERE Usuario = @usuario AND Fabricante = @codFabricante AND Produto = @codProduto AND IdPagina = @paginaId";
								cmdUpdate = new SqlCommand(strSql, cn);
								cmdUpdate.Parameters.AddRange(parameters);
								intRetorno = cmdUpdate.ExecuteNonQuery();

								if (intRetorno == 0)
								{
									cmdUpdate.Parameters.Clear();
									strSql = "INSERT INTO t_RELATORIO_PRODUTO_FLAG (IdPagina, Usuario, Fabricante, Produto, DataHoraCadastro, Flag) VALUES (@paginaId, @usuario, @codFabricante, @codProduto, getdate(), @flag)";
									cmdInsere = new SqlCommand(strSql, cn);
									cmdInsere.Parameters.AddRange(parameters);
									intRetorno = cmdInsere.ExecuteNonQuery();
								}

								if (intRetorno == 0)
								{
									throw new Exception("Não foi possível alterar ou inserir flag");
								}
							}
							finally
							{
								cn.Close();
							}
						}
					}
					catch (Exception)
					{
						throw;
					}
                });
            #endregion
        }
    }
}