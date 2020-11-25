namespace ART3WebAPI.Models.Entities
{
    public class Farol
    {
        public string Fabricante {get; set; }
        public string Produto { get; set; }
        public string Descricao { get; set; }
        public string Descricao_html { get; set; }
        public string Grupo { get; set; }
        public string Subgrupo { get; set; }
        public int? Potencia_BTU { get; set; }
        public string Ciclo { get; set; }
        public string Posicao_mercado { get; set; }
        public int Farol_qtde_comprada { get; set; }
        public int Qtde_vendida { get; set; }
        public int Qtde_devolvida { get; set; }
        public int Qtde_estoque_venda { get; set; }
        public int Saldo { get; set; }
        public int Qtde_composto_item { get; set; }
        public decimal Custo { get; set; }
        public int[] Meses { get; set; }
        public bool IsItemComposicao { get; set; } = false;

        public Farol() { }

        public Farol(string fabricante,
                     string produto,
                     string descricao,
                     string descricao_html,
                     string grupo,
                     string subgrupo,
                     int potencia_BTU,
                     string ciclo,
                     string posicao_mercado,
                     int farol_qtde_comprada,
                     int qtde_vendida,
                     int qtde_devolvida,
                     int qtde_estoque_venda,
                     decimal custo,
                     int[] meses
                     )
        {

            this.Fabricante = fabricante;
            this.Produto = produto;
            this.Descricao = descricao;
            this.Descricao_html = descricao_html;
            this.Grupo = grupo;
            this.Subgrupo = subgrupo;
            this.Potencia_BTU = potencia_BTU;
            this.Ciclo = ciclo;
            this.Posicao_mercado = posicao_mercado;
            this.Farol_qtde_comprada = farol_qtde_comprada;
            this.Qtde_vendida = qtde_vendida - qtde_devolvida;
            this.Qtde_devolvida = qtde_devolvida;
            this.Qtde_estoque_venda = qtde_estoque_venda;
            this.Custo = custo;
            this.Meses = meses;

        }
    }
}