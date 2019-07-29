using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Financeiro
{
    class PlanilhaRepasseMktplaceN1
    {
        public int Id { get; set; }
        public string Checksum { get; set; }
        public string Usuario { get; set; }
        public string NomeArquivo { get; set; }
        public string Path { get; set; }
        public string OrigemGrupo { get; set; }
        public DateTime DataCadastro { get; set; }
        public DateTime DataHoraCadastro { get; set; }
    }

    class PlanilhaRepasseMktplaceN2
    {
        public int Id { get; set; }
        public int MktplceRepasseN1Id { get; set; }
        public int LojaOrigemId { get; set; }
        public string Pedido { get; set; }
        public DateTime DataPedido { get; set; }
        public int TipoPagamentoId { get; set; }
        public int MeioPagamentoId { get; set; }
        public decimal ValorTotalPedido { get; set; }
        public decimal ValorTotalPedidoCorreto { get; set; }
        public bool PedidoExiste { get; set; }
        public bool ValorTotalConfere
        {
            get { return ValorTotalPedido == ValorTotalPedidoCorreto; }
        }
        public int Linha { get; set; }
    }

    class PlanilhaRepasseMktplaceN3
    {
        public int Id { get; set; }
        public int MktplceRepasseN2Id { get; set; }
        public int ProdutoId { get; set; }
        public decimal ValorFrete { get; set; }
        public decimal ValorItem { get; set; }
        public decimal ValorItemBruto { get; set; }
        public decimal PercComissao { get; set; }
        public int Linha { get; set; }

    }

    class PlanilhaRepasseMktplaceN4
    {
        public int MktplceRepasseN3Id { get; set; }
        public int TipoTransacao { get; set; }
        public int StatusTransacaoId { get; set; }
        public decimal Valor { get; set; }
        public DateTime DataPagamento { get; set; }
        public DateTime DataLiberacao { get; set; }
        public DateTime DataEstorno { get; set; }
        public int Linha { get; set; }
    }

    #region [ MktplaceRepasseTipoTransacao ]
    /// <summary>
    /// Type safe enum pattern
    /// Esta classe contém todos os tipos de transação conhecidas nas planilhas de pagamento marketplace
    /// Que podem ser recuperados através dos métodos de acesso
    /// </summary>
    public sealed class MktplaceRepasseTipoTransacao
    {
        private readonly int Id;
        private readonly string Descricao;
        private readonly string Empresa;

        public int inteiroTeste { get; }

        public static MktplaceRepasseTipoTransacao Venda = new MktplaceRepasseTipoTransacao(1, "Venda", "B2W");
        public static MktplaceRepasseTipoTransacao Comissao = new MktplaceRepasseTipoTransacao(2, "Comissao", "B2W");
        public static MktplaceRepasseTipoTransacao EstornoVenda = new MktplaceRepasseTipoTransacao(3, "Estorno_Venda", "B2W");
        public static MktplaceRepasseTipoTransacao EstornoComissao = new MktplaceRepasseTipoTransacao(4, "Estorno_Comissao", "B2W");
        public static MktplaceRepasseTipoTransacao EstornoComissaoSemDesbloqueio = new MktplaceRepasseTipoTransacao(5, "Estorno_Comissao_Sem_Desbloqueio", "B2W");
        public static MktplaceRepasseTipoTransacao ComissaoSemDesbloqueio = new MktplaceRepasseTipoTransacao(6, "Comissao_Sem_Desbloqueio", "B2W");
        public static MktplaceRepasseTipoTransacao Bonus = new MktplaceRepasseTipoTransacao(7, "Bonus", "B2W");
        
        private MktplaceRepasseTipoTransacao(int Id, string Descricao, string Empresa)
        {
            this.Id = Id;
            this.Descricao = Descricao;
            this.Empresa = Empresa;
        }

        public string GetDescricao()
        {
            return Descricao;
        }

        public string GetEmpresa()
        {
            return Empresa;
        }

        public int GetId()
        {
            return Id;
        }

        public static int GetId(string descricao, string empresa, ref string strMsgErro)
        {
            strMsgErro = "";
            if (descricao == null) return 0;

            Type type = typeof(MktplaceRepasseTipoTransacao);
            foreach (var p in type.GetFields())
            {
                var v = p.GetValue(null);
                if ((descricao.Equals(((MktplaceRepasseTipoTransacao)v).GetDescricao())) && (empresa.Equals(((MktplaceRepasseTipoTransacao)v).GetEmpresa()))) return ((MktplaceRepasseTipoTransacao)v).GetId();
            }

            if (descricao.Trim().Length > 0)
            {
                strMsgErro = "Tipo de transação Marketplace B2W desconhecida: " + descricao;
                return 0;
            }
            return 0;
        }
    } 
    #endregion
}
