using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
    public class OcorrenciasStatus
    {
        public string  Loja { get; set; }
        public string  CD { get; set; }
        public string  Pedido { get; set; }
        public int  NF { get; set; }
        public string  Cliente { get; set; }
        public string  UF { get; set; }
        public string  Cidade { get; set; }
        public string  Transportadora { get; set; }
        public string  Contato { get; set; }
        public string  Telefone { get; set; }
        public string  Ocorrencia { get; set;}
        public string  TipoOcorrencia { get; set; }
        public string  Status { get; set; }
    }
}