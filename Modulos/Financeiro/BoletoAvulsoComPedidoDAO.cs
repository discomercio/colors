#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	class BoletoAvulsoComPedidoDAO
	{
		#region [ obtemBoletoPlanoContasDestino ]
		/// <summary>
		/// Para cada pedido que participa do rateio, localiza a loja ao qual o pedido pertence.
		/// Para cada loja, obtém o plano de contas para o qual deverá ser lançado o lançamento 
		/// do fluxo de caixa gerado em decorrência do boleto.
		/// Gera uma exceção no caso de não encontrar nenhum plano de contas ou se houver mais
		/// do que 1 plano de contas.
		/// </summary>
		/// <param name="listaPedidoRateio">Lista com os detalhes dos pedidos que participam do rateio</param>
		/// <returns>Retorna um objeto do tipo BoletoPlanoContasDestino com os dados do plano de contas</returns>
		public static BoletoPlanoContasDestino obtemBoletoPlanoContasDestino(List<String> listaPedidoRateio)
		{
			#region [ Declarações ]
			BoletoPlanoContasDestino boletoPlanoContasDestino = new BoletoPlanoContasDestino();
			String strSql;
			String strPedidos = "";
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Prepara objetos de acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Lista dos pedidos que participam do rateio ]
			for (int i = 0; i < listaPedidoRateio.Count; i++)
			{
				if (listaPedidoRateio[i].ToString().Trim().Length > 0)
				{
					if (strPedidos.Length > 0) strPedidos += ", ";
					strPedidos += "'" + listaPedidoRateio[i].ToString().Trim() + "'";
				}
			}

			if (strPedidos.Length == 0)
			{
				throw new FinanceiroException("Não há nenhum pedido na lista de rateio do boleto!!");
			}
			#endregion

			#region [ Consulta SQL ]
			strSql = "SELECT DISTINCT" +
						" id_plano_contas_empresa," +
						" id_plano_contas_grupo," +
						" id_plano_contas_conta," +
						" natureza" +
					" FROM t_PEDIDO tP" +
						" INNER JOIN t_LOJA tL ON (tP.loja=tL.loja)" +
					" WHERE" +
						" (tP.pedido IN (" + strPedidos + "))";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Consistência ]
			if (dtbResultado.Rows.Count == 0)
			{
				throw new FinanceiroException("Não há informações do plano de contas (pedido(s)=" + strPedidos + ")");
			}
			else if (dtbResultado.Rows.Count > 1)
			{
				throw new FinanceiroException("Há mais de 1 plano de contas (pedido(s)=" + strPedidos + ")");
			}
			rowResultado = dtbResultado.Rows[0];
			if (BD.readToInt(rowResultado["id_plano_contas_conta"]) == 0)
			{
				throw new FinanceiroException("A informação do plano de contas não foi preenchida adequadamente no cadastro de lojas (pedido(s)=" + strPedidos + ")!!");
			}
			#endregion

			#region [ Carrega os dados ]
			boletoPlanoContasDestino.id_plano_contas_empresa = BD.readToByte(rowResultado["id_plano_contas_empresa"]);
			boletoPlanoContasDestino.id_plano_contas_grupo = BD.readToShort(rowResultado["id_plano_contas_grupo"]);
			boletoPlanoContasDestino.id_plano_contas_conta = BD.readToInt(rowResultado["id_plano_contas_conta"]);
			boletoPlanoContasDestino.natureza = BD.readToChar(rowResultado["natureza"]);
			#endregion

			return boletoPlanoContasDestino;
		}
		#endregion
	}
}
