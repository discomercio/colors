using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	public class Produto
	{
		public string fabricante { get; set; }
		public string produto { get; set; }
		public string descricao { get; set; }
		public string ean { get; set; }
		public string grupo { get; set; }
		public decimal preco_fabricante { get; set; }
		public int estoque_critico { get; set; }
		public double peso { get; set; }
		public int qtde_volumes { get; set; }
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public int excluido_status { get; set; }
		public decimal vl_custo2 { get; set; }
		public string descricao_html { get; set; }
		public double cubagem { get; set; }
		public string ncm { get; set; }
		public string cst { get; set; }
		public double perc_MVA_ST { get; set; }
		public int deposito_zona_id { get; set; }
		public string deposito_zona_usuario_ult_atualiz { get; set; }
		public DateTime deposito_zona_dt_hr_ult_atualiz { get; set; }
		public int farol_qtde_comprada { get; set; }
		public string farol_qtde_comprada_usuario_ult_atualiz { get; set; }
		public DateTime farol_qtde_comprada_dt_hr_ult_atualiz { get; set; }
		public string descontinuado { get; set; }
		public int potencia_BTU { get; set; }
		public string ciclo { get; set; }
		public string posicao_mercado { get; set; }
	}
}