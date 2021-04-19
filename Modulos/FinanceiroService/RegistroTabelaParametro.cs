using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinanceiroService
{
	class RegistroTabelaParametro
	{
		public string id { get; set; }
		public int campo_inteiro { get; set; }
		public decimal campo_monetario { get; set; }
		public double campo_real { get; set; }
		public DateTime campo_data { get; set; }
		public string campo_texto { get; set; }
		public string campo_2_texto { get; set; }
		public DateTime dt_hr_ult_atualizacao { get; set; }
		public string usuario_ult_atualizacao { get; set; }
		public string obs { get; set; }
	}
}
