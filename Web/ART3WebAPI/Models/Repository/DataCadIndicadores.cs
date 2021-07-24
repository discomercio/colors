using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models.Entities;

namespace ART3WebAPI.Models.Repository
{
    public class DataCadIndicadores
    {
        public Indicador[] GetIndicador(string loja)
        {
            List<Indicador> listaInd = new List<Indicador>();
            SqlConnection cn = new SqlConnection(BD.getConnectionString());
            string s_where;

            if ((loja != "") && (loja != "vazio"))
            {
                s_where = "loja=" + loja + " AND";
            }
            else
            {
                s_where = "";
            }

			cn.Open();
			try // Finally: cn.Close()
			{
                StringBuilder sqlString = new StringBuilder();
                sqlString.AppendLine("SELECT razao_social_nome,uf,email,vendedor FROM t_ORCAMENTISTA_E_INDICADOR WHERE " + s_where + " status='A' ORDER BY razao_social_nome");
                
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandText = sqlString.ToString();
                IDataReader reader = cmd.ExecuteReader();
                
                try
                {
                    int idxNome = reader.GetOrdinal("razao_social_nome");
                    int idxEmail = reader.GetOrdinal("email");
                    int idxUF = reader.GetOrdinal("uf");
                    int idxVendedor = reader.GetOrdinal("vendedor");

                    while (reader.Read())
                    {
                        Indicador _novo = new Indicador();
                        _novo.Nome = reader.GetString(idxNome);
                        _novo.Email = reader.IsDBNull(idxEmail) ? "" : reader.GetString(idxEmail);
                        _novo.Uf = reader.IsDBNull(idxUF) ? "" : reader.GetString(idxUF);
                        _novo.Vendedor = reader.IsDBNull(idxVendedor) ? "" : reader.GetString(idxVendedor);
                        listaInd.Add(_novo);
                    }
                }
                finally
                {
                    reader.Close();
                }
            }
            finally
            {
                cn.Close();
            }

            return listaInd.ToArray();
        }
    }
}