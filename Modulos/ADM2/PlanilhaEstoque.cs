using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADM2
{
	class PlanilhaEstoqueColumn
	{
		public int ColIndex { get; set; }
		public string ColTitle { get; set; } = string.Empty;
		public string ColTitleEsperado { get; set; } = string.Empty;
	}

	class PlanilhaEstoqueHeader
	{
		public PlanilhaEstoqueColumn Sku { get; set; } = new PlanilhaEstoqueColumn();
		public PlanilhaEstoqueColumn ProdutoDescricao { get; set; } = new PlanilhaEstoqueColumn();
		public PlanilhaEstoqueColumn QtdeEstoque { get; set; } = new PlanilhaEstoqueColumn();
		public PlanilhaEstoqueColumn ValorCustoMedio { get; set; } = new PlanilhaEstoqueColumn();
		public PlanilhaEstoqueColumn PrecoLista { get; set; } = new PlanilhaEstoqueColumn();
	}

	class PlanilhaEstoqueLinha
	{
		public string Sku { get; set; } = string.Empty;
		public string SkuFormatado { get; set; } = string.Empty;
		public string ProdutoDescricao { get; set; } = string.Empty;
		public double? QtdeEstoque { get; set; } = null;
		public double? ValorCustoMedio { get; set; } = null;
		public double? PrecoLista { get; set; } = null;
	}
}
