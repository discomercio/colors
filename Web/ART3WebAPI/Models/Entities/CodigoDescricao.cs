using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	#region [ CodigoDescricao ]
	public class CodigoDescricao
	{
		public string grupo { get; set; }
		public string codigo { get; set; }
		public int ordenacao { get; set; }
		public byte st_inativo { get; set; }
		public string descricao { get; set; }
		public DateTime dt_hr_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public DateTime dt_hr_ult_atualizacao { get; set; }
		public string usuario_ult_atualizacao { get; set; }
		public byte st_possui_sub_codigo { get; set; }
		public byte st_eh_sub_codigo { get; set; }
		public string grupo_pai { get; set; }
		public string codigo_pai { get; set; }
		public string lojas_habilitadas { get; set; }
		public byte parametro_1_campo_flag { get; set; }
		public byte parametro_2_campo_flag { get; set; }
		public byte parametro_3_campo_flag { get; set; }
		public byte parametro_4_campo_flag { get; set; }
		public byte parametro_5_campo_flag { get; set; }
		public int parametro_campo_inteiro { get; set; }
		public decimal parametro_campo_monetario { get; set; }
		public double parametro_campo_real { get; set; }
		public DateTime parametro_campo_data { get; set; }
		public string parametro_campo_texto { get; set; }
		public string parametro_2_campo_texto { get; set; }
        public string parametro_3_campo_texto { get; set; }
		public string parametro_4_campo_texto { get; set; }
		public string descricao_parametro { get; set; }
	}
	#endregion
}