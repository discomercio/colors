using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADM2
{
	public class PedidoFrete
	{
		public int id { get; set; }
		public string pedido { get; set; }
		public string codigo_tipo_frete { get; set; }
		public string descricao_tipo_frete { get; set; } = ""; // Campo usado quando se pesquisa fretes cadastrados no pedido
		public decimal vl_frete { get; set; }
		public decimal vl_frete_original_EDI { get; set; }
		public decimal vl_NF { get; set; }
		public string transportadora_id { get; set; }
		public string transportadora_cnpj { get; set; }
		public string emissor_cnpj { get; set; }
		public int id_nfe_emitente { get; set; }
		public int serie_NF { get; set; }
		public int numero_NF { get; set; }
		public int tipo_preenchimento { get; set; }
		public int id_editrp_arq_input_linha_processada_n1 { get; set; }
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_hr_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public DateTime dt_hr_ult_atualizacao { get; set; }
		public string usuario_ult_atualizacao { get; set; }
	}
}
