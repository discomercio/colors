using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	public class RelEstoqueVendaEntity
	{
		public string FabricanteCodigo { get; set; }
		public string FabricanteDescricao { get; set; }
		public List<RelEstoqueVendaItemEntity> Produtos { get; set; } = new List<RelEstoqueVendaItemEntity>();
	}

	public class RelEstoqueVendaItemEntity
	{
		public string ProdutoCodigo { get; set; }
		public string ProdutoDescricao { get; set; }
		public string ProdutoDescricaoHtml { get; set; }
		public double Cubagem { get; set; }
		public int Qtde { get; set; }
		public decimal VlTotal { get; set; }
	}
}