using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	class PlanilhaPrecosColumn
	{
		public int ColIndex { get; set; } = 0;
		public string ColTitle { get; set; } = string.Empty;
		public string ColTitleEsperado { get; set; } = string.Empty;
	}

	class PlanilhaPrecosHeader
	{
		public PlanilhaPrecosColumn Codigo { get; set; } = new PlanilhaPrecosColumn();
		public PlanilhaPrecosColumn ProdutoDescricao { get; set; } = new PlanilhaPrecosColumn();
		public PlanilhaPrecosColumn QtdeLojasConcorrentes { get; set; } = new PlanilhaPrecosColumn();
		public PlanilhaPrecosColumn Status { get; set; } = new PlanilhaPrecosColumn();
		public PlanilhaPrecosColumn SeuPreco { get; set; } = new PlanilhaPrecosColumn();
		public PlanilhaPrecosColumn Diferenca { get; set; } = new PlanilhaPrecosColumn();
		public PlanilhaPrecosColumn Regra { get; set; } = new PlanilhaPrecosColumn();
		public PlanilhaPrecosColumn Sugestao { get; set; } = new PlanilhaPrecosColumn();
		public List<PlanilhaPrecosColumn> ColunasPrecoConcorrente { get; set; } = new List<PlanilhaPrecosColumn>();
	}

	class PlanilhaPrecosPrecoConcorrente : PlanilhaPrecosColumn
	{
		public string CellValue { get; set; } = string.Empty;
		public decimal? Preco { get; set; } = null;
	}

	class PlanilhaPrecosLinha
	{
		public string Codigo { get; set; } = string.Empty;
		public string CodigoFormatado { get; set; } = string.Empty;
		public string ProdutoDescricao { get; set; } = string.Empty;
		public int QtdeLojas { get; set; } = 0;
		public string Status { get; set; } = string.Empty;
		public decimal? SeuPreco { get; set; } = null;
		public decimal? Diferenca { get; set; } = null;
		public string Regra { get; set; } = string.Empty;
		public decimal? PrecoSugestao { get; set; } = null;
		public List<PlanilhaPrecosPrecoConcorrente> ColunasPrecoConcorrente { get; set; } = new List<PlanilhaPrecosPrecoConcorrente>();
		public string NomeConcorrentePrecoMinimo { get; set; } = string.Empty;
		public decimal? PrecoMinimo { get; set; } = null;
		public decimal? PrecoMedio { get; set; } = null;
		public bool ProcessadoStatus { get; set; } = false;
	}
}
