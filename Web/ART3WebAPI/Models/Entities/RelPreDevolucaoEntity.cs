using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	public class RelPreDevolucaoEntity
	{
		public int id_devolucao { get; set; }
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_hr_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public string loja { get; set; }
		public string pedido { get; set; }
		public DateTime data_pedido { get; set; }
		public string vendedor { get; set; }
		public string indicador { get; set; }
		public string cliente_id { get; set; }
		public string cliente_nome { get; set; }
		public string cliente_cpf { get; set; }
		public string transportadora_id { get; set; }
		public string cod_procedimento { get; set; }
		public string cod_devolucao_motivo { get; set; }
		public string descricao_devolucao_motivo { get; set; }
		public string cod_credito_transacao { get; set; }
		public decimal vl_pedido { get; set; }
		public decimal vl_devolucao { get; set; }
		public byte status { get; set; }
		public DateTime status_data_hora { get; set; }
		public string status_usuario { get; set; }
	}
}