using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	class ProdutoConferePreco
	{
		public string sku { get; set; } = string.Empty;
		public string skuFormatado { get; set; } = string.Empty;
		public string priceCsv { get; set; } = string.Empty;
		public decimal vlPriceCsv { get; set; } = 0m;
		public bool isCadastradoMagento { get; set; } = false;
		public string product_id { get; set; } = string.Empty;
		public string name { get; set; } = string.Empty;
		public string priceMagento { get; set; } = string.Empty;
		public decimal vlPriceMagento { get; set; } = 0m;
		public string campoOrdenacao { get; set; } = string.Empty;
		public decimal vlDiferenca { get; set; } = 0m;
		public double? percDiferenca { get; set; } = null;
	}
}
