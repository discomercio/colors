using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
    public class Compras
    {
        public string NF { get; set; }
        public string Fabricante { get; set; }
        public string Produto { get; set; }
        public int Qtde { get; set; }
        public decimal Valor { get; set; }
        public decimal[] Meses { get; set; }

    }
}