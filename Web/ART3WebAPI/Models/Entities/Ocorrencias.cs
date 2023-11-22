using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
    public class Ocorrencias
    {
        public string  Pedido{ get; set; }
        public int  NF { get; set; }
        public string  Transportadora { get; set; }
        public string  Ocorrencia { get; set;}
        public string  TipoOcorrencia { get; set; }
		public string Loja { get; set; }
		public string Indicador { get; set; }
		public string Vendedor { get; set; }
	}
}