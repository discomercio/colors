using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
    public class DevolucaoProduto2Entity
    {
        public DateTime DataPedido { get; set; }
        public DateTime DataDevolvido { get; set; }
        public DateTime DataBaixa { get; set; }
        public int Qtde { get; set; }
        public string Fabricante { get; set; }
        public string Produto { get; set; }
        public string Descricao { get; set; }
        public string Pedido { get; set; }
        public string Cliente { get; set; }
        public string Vendedor { get; set; }
        public string Indicador { get; set; }
        public string Motivo { get; set; }
    }
}