using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
    public class Empresa
    {
        public string Cep { get; set; }
        public string Cidade { get; set; }
        public string Cnpj { get; set; }
        public string Contato { get; set; }
        public string Email { get; set; }
        public string Endereco { get; set; }
        public string NomeFantasia { get; set; }
        public string RazaoSocial { get; set; }
        public string TaxaAdm { get; set; }
        public string Telefone { get; set; }
        public string Uf { get; set; }
    }
}