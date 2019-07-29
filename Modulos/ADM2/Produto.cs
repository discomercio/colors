using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADM2
{
	#region [ ProdutoCadastroBasico ]
	public class ProdutoCadastroBasico
	{
		public string fabricante { get; set; } = string.Empty;
		public string produto { get; set; } = string.Empty;
		public string descricao { get; set; } = string.Empty;
		public bool isCadastrado { get; set; } = false;
		public bool isComposto { get; set; } = false;
	}
	#endregion

	#region [ ProdutoEstoqueVenda ]
	public class ProdutoEstoqueVenda
	{
		public string fabricante { get; set; } = string.Empty;
		public string produto { get; set; } = string.Empty;
		public string descricao { get; set; } = string.Empty;
		public int qtdeEstoqueVenda { get; set; } = 0;
		public decimal vlCustoIntermediario { get; set; } = 0m;
		public bool isCadastrado { get; set; } = false;
		public bool isComposto { get; set; } = false;
	}
	#endregion

	#region [ ProdutoEstoqueVendaLojaConsolidado ]
	public class ProdutoEstoqueVendaLojaConsolidado : ProdutoEstoqueVenda
	{
		public bool isVendavel { get; set; } = false;
		public int QtdeAmbientesComEstoqueDisponivel { get; set; } = 0;
		public List<EstoqueVendaAmbiente> listaEstoqueVendaAmbiente = new List<EstoqueVendaAmbiente>();
	}
	#endregion

	#region [ EstoqueVendaAmbiente ]
	public class EstoqueVendaAmbiente
	{
		public int qtdeEstoqueVenda { get; set; } = 0;
		public decimal vlCustoIntermediario { get; set; } = 0m;
		public string nomeAmbiente { get; set; } = string.Empty;

		#region [ Constructor ]
		public EstoqueVendaAmbiente(int QtdeEstoqueVenda, decimal ValorCustoIntermediario, string NomeAmbiente)
		{
			qtdeEstoqueVenda = QtdeEstoqueVenda;
			vlCustoIntermediario = ValorCustoIntermediario;
			nomeAmbiente = NomeAmbiente;
		}
		#endregion
	}
	#endregion

	#region [ ProdutoCompostoItem ]
	public class ProdutoCompostoItem
	{
		public string fabricante_composto { get; set; }
		public string produto_composto { get; set; }
		public string fabricante_item { get; set; }
		public string produto_item { get; set; }
		public int qtde { get; set; }
	}
	#endregion

	#region [ ProdutoCompostoItemCalculoPrecoLista ]
	public class ProdutoCompostoItemCalculoPrecoLista : ProdutoCompostoItem
	{
		public double razaoPrecoListaTotalPorUnidade { get; set; } = 0d;
		public double coeficienteCustoFinanceiro { get; set; } = 0d;
		public decimal preco_lista_a_vista_atual { get; set; } = 0m;
		public decimal preco_lista_a_prazo_novo { get; set; } = 0m;
		public decimal preco_lista_a_vista_novo { get; set; } = 0m;
		public ProdutoCompostoItemCalculoPrecoLista(string fabricanteComposto, string produtoComposto, string fabricanteItem, string produtoItem, int quantidade)
		{
			fabricante_composto = fabricanteComposto;
			produto_composto = produtoComposto;
			fabricante_item = fabricanteItem;
			produto_item = produtoItem;
			qtde = quantidade;
		}
	}
	#endregion

	#region [ ProdutoCompostoCalculaPrecoLista ]
	public class ProdutoCompostoCalculaPrecoLista
	{
		public decimal preco_lista_a_vista_total_atual { get; set; } = 0m;
		public decimal preco_lista_a_prazo_total_novo { get; set; } = 0m;
		public bool ocorreuErro { get; set; } = false;
		public string mensagem_erro { get; set; } = string.Empty;

		public List<ProdutoCompostoItemCalculoPrecoLista> Itens = new List<ProdutoCompostoItemCalculoPrecoLista>();
	}
	#endregion

	#region [ ProdutoLoja ]
	public class ProdutoLoja
	{
		public string fabricante { get; set; }
		public string produto { get; set; }
		public string loja { get; set; }
		public decimal preco_lista { get; set; }
		public double margem { get; set; }
		public double desc_max { get; set; }
		public double comissao { get; set; }
		public string vendavel { get; set; }
		public int qtde_max_venda { get; set; }
		public string cor { get; set; }
		public DateTime dt_cadastro { get; set; }
		public DateTime dt_ult_atualizacao { get; set; }
		public int excluido_status { get; set; }
	}
	#endregion

	#region [ ProdutoUnificadoPrecoLista ]
	public class ProdutoUnificadoPrecoLista
	{
		public string fabricante { get; set; } = string.Empty;
		public string produto { get; set; } = string.Empty;
		public string loja { get; set; } = string.Empty;
		public decimal preco_lista { get; set; } = 0;
		public string descricao { get; set; } = string.Empty;
		public bool isCadastrado { get; set; } = false;
		public bool isComposto { get; set; } = false;
	}
	#endregion

	#region [ PercentualCustoFinanceiroFornecedor ]
	public class PercentualCustoFinanceiroFornecedor
	{
		public string fabricante { get; set; }
		public string tipo_parcelamento { get; set; }
		public int qtde_parcelas { get; set; }
		public double coeficiente { get; set; }
	}
	#endregion
}
