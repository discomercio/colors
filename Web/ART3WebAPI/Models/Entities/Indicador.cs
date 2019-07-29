
namespace ART3WebAPI.Models.Entities
{
    public class Indicador
    {
        public string Nome { get; set; }
        public string TipoDocumento { get; set; }
        public string TipoPessoa { get; set; }
        public string CpfCnpj { get; set; }
        public string Banco { get; set; }
        public string Agencia { get; set; }
        public string DigitoAgencia { get; set; }
        public string Operacao { get; set; }
        public string Conta { get; set; }
        public string DigitoConta { get; set; }
        public string TipoConta { get; set; }
        public string TipoContaCSV { get; set; }
        public decimal Valor { get; set; }
        public string Vendedor { get; set; }
        public string Email { get; set; }
        public string Uf { get; set; }
    }
}