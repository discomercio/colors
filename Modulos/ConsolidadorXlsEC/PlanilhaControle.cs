using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	class PlanilhaControleColumn
	{
		public int ColIndex { get; set; } = 0;
		public string ColTitle { get; set; } = string.Empty;
		public string ColTitleEsperado { get; set; } = string.Empty;
	}

	class PlanilhaControleHeader
	{
		public PlanilhaControleColumn Sku { get; set; } = new PlanilhaControleColumn();
		public PlanilhaControleColumn Asterisco { get; set; } = new PlanilhaControleColumn();
		public PlanilhaControleColumn ProdutoDescricao { get; set; } = new PlanilhaControleColumn();
		public PlanilhaControleColumn QtdeEstoque { get; set; } = new PlanilhaControleColumn();
		public PlanilhaControleColumn ValorCustoMedio { get; set; } = new PlanilhaControleColumn();
		public PlanilhaControleColumn ValorMedioMercado { get; set; } = new PlanilhaControleColumn();
		public PlanilhaControleColumn ValorMinimoMercado { get; set; } = new PlanilhaControleColumn();
		public PlanilhaControleColumn PrecoFinal { get; set; } = new PlanilhaControleColumn();
	}

	class PlanilhaControleLinha
	{
		public string Sku { get; set; } = string.Empty;
		public string SkuFormatado { get; set; } = string.Empty;
		public string Asterisco { get; set; } = string.Empty;
		public string ProdutoDescricao { get; set; } = string.Empty;
		public double? QtdeEstoque { get; set; } = null;
		public double? ValorCustoMedio { get; set; } = null;
		public double? ValorMedioMercado { get; set; } = null;
		public double? ValorMinimoMercado { get; set; } = null;
		public decimal? PrecoFinal { get; set; } = null;
	}
}
