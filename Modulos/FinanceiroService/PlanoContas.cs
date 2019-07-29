using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinanceiroService
{
	class PlanoContasConta
	{
		public int id { get; set; }
		public char natureza { get; set; }
		public short id_plano_contas_grupo { get; set; }
		public byte st_ativo { get; set; }
		public byte st_sistema { get; set; }
		public string descricao { get; set; }
		public DateTime dt_cadastro { get; set; }
		public string usuario_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public string usuario_ult_atualizacao { get; set; }
	}
}
