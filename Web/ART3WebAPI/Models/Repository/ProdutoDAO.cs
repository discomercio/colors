using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using System.Data;
using ART3WebAPI.Models.Domains;
using ART3WebAPI.Models.Entities;

namespace ART3WebAPI.Models.Repository
{
	public class ProdutoDAO
	{
		#region [ produtoLoadFromDataRow ]
		public static Produto produtoLoadFromDataRow(DataRow rowDados)
		{
			#region [ Declarações ]
			Produto produto = new Produto();
			#endregion

			produto.fabricante = BD.readToString(rowDados["fabricante"]);
			produto.produto = BD.readToString(rowDados["produto"]);
			produto.descricao = BD.readToString(rowDados["descricao"]);
			produto.ean = BD.readToString(rowDados["ean"]);
			produto.grupo = BD.readToString(rowDados["grupo"]);
			produto.preco_fabricante = BD.readToDecimal(rowDados["preco_fabricante"]);
			produto.estoque_critico = BD.readToInt(rowDados["estoque_critico"]);
			produto.peso = BD.readToSingle(rowDados["peso"]);
			produto.qtde_volumes = BD.readToInt(rowDados["qtde_volumes"]);
			produto.dt_cadastro = BD.readToDateTime(rowDados["dt_cadastro"]);
			produto.dt_ult_atualizacao = BD.readToDateTime(rowDados["dt_ult_atualizacao"]);
			produto.excluido_status = BD.readToInt(rowDados["excluido_status"]);
			produto.vl_custo2 = BD.readToDecimal(rowDados["vl_custo2"]);
			produto.descricao_html = BD.readToString(rowDados["descricao_html"]);
			produto.cubagem = BD.readToSingle(rowDados["cubagem"]);
			produto.ncm = BD.readToString(rowDados["ncm"]);
			produto.cst = BD.readToString(rowDados["cst"]);
			produto.perc_MVA_ST = BD.readToSingle(rowDados["perc_MVA_ST"]);
			produto.deposito_zona_id = BD.readToInt(rowDados["deposito_zona_id"]);
			produto.deposito_zona_usuario_ult_atualiz = BD.readToString(rowDados["deposito_zona_usuario_ult_atualiz"]);
			produto.deposito_zona_dt_hr_ult_atualiz = BD.readToDateTime(rowDados["deposito_zona_dt_hr_ult_atualiz"]);
			produto.farol_qtde_comprada = BD.readToInt(rowDados["farol_qtde_comprada"]);
			produto.farol_qtde_comprada_usuario_ult_atualiz = BD.readToString(rowDados["farol_qtde_comprada_usuario_ult_atualiz"]);
			produto.farol_qtde_comprada_dt_hr_ult_atualiz = BD.readToDateTime(rowDados["farol_qtde_comprada_dt_hr_ult_atualiz"]);
			produto.descontinuado = BD.readToString(rowDados["descontinuado"]);
			produto.potencia_BTU = BD.readToInt(rowDados["potencia_BTU"]);
			produto.ciclo = BD.readToString(rowDados["ciclo"]);
			produto.posicao_mercado = BD.readToString(rowDados["posicao_mercado"]);

			return produto;
		}
		#endregion

		#region [ getProduto ]
		public static Produto getProduto(string codFabricante, string codProduto, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			Produto produto;
			#endregion

			msg_erro = "";
			try
			{
				if (Global.digitos(codFabricante ?? "").Length == 0)
				{
					msg_erro = "Código do fabricante é inválido!";
					return null;
				}

				if (Global.digitos(codProduto ?? "").Length == 0)
				{
					msg_erro = "Código do produto é inválido!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_PRODUTO" +
							" WHERE" +
								" (fabricante = '" + codFabricante.Trim() + "')" +
								" AND (produto = '" + codProduto.Trim() + "')";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Nenhum produto encontrado com o código '" + codProduto.Trim() + "' (fabricante: " + codFabricante.Trim() + ")";
						return null;
					}

					produto = produtoLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return produto;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion

		#region [ getProdutoBySku ]
		public static Produto getProdutoBySku(string codProduto, out string msg_erro)
		{
			#region [ Declarações ]
			string strSql;
			SqlConnection cn;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			Produto produto;
			#endregion

			msg_erro = "";
			try
			{
				if (Global.digitos(codProduto ?? "").Length == 0)
				{
					msg_erro = "Código do produto é inválido!";
					return null;
				}

				#region [ Prepara acesso ao BD ]
				cn = new SqlConnection(BD.getConnectionString());
				cn.Open();
				cmCommand = new SqlCommand();
				cmCommand.Connection = cn;
				daDataAdapter = new SqlDataAdapter();
				#endregion

				try // finally: BD.fechaConexao(ref cn);
				{
					#region [ Monta Select ]
					strSql = "SELECT " +
								"*" +
							" FROM t_PRODUTO" +
							" WHERE" +
								" (produto = '" + codProduto.Trim() + "')";
					#endregion

					#region [ Executa a consulta ]
					cmCommand.CommandText = strSql;
					daDataAdapter.SelectCommand = cmCommand;
					daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
					daDataAdapter.Fill(dtbResultado);
					#endregion

					if (dtbResultado.Rows.Count == 0)
					{
						msg_erro = "Nenhum produto encontrado com o código '" + codProduto.Trim() + "'";
						return null;
					}

					produto = produtoLoadFromDataRow(dtbResultado.Rows[0]);
				}
				finally
				{
					BD.fechaConexao(ref cn);
				}

				return produto;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return null;
			}
		}
		#endregion
	}
}