using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using ART3WebAPI.Models.Entities;

namespace ART3WebAPI.Models.Repository
{
    public class DataCiagroup
    {
        public Empresa empresa;
        
        public DataCiagroup()
        {
            empresa = GetEmpresa();
        }      

        public Indicador[] Get(int id)
        {
            List<Indicador> listaInd = new List<Indicador>();
            SqlConnection cn = new SqlConnection(BD.getConnectionString());
            cn.Open();

            try
            {
                StringBuilder sqlString = new StringBuilder();

                sqlString.AppendLine("SELECT t_ORCAMENTISTA_E_INDICADOR.favorecido, t_ORCAMENTISTA_E_INDICADOR.favorecido_cnpj_cpf, t_ORCAMENTISTA_E_INDICADOR.banco, ");
                sqlString.Append("t_ORCAMENTISTA_E_INDICADOR.agencia, t_ORCAMENTISTA_E_INDICADOR.agencia_dv, t_ORCAMENTISTA_E_INDICADOR.conta, t_ORCAMENTISTA_E_INDICADOR.conta_dv, ");
                sqlString.Append("t_ORCAMENTISTA_E_INDICADOR.conta_operacao, t_ORCAMENTISTA_E_INDICADOR.tipo_conta, t_COMISSAO_INDICADOR_N3.vl_total_pagto, t_COMISSAO_INDICADOR_N2.vendedor ");
                sqlString.AppendLine("FROM t_COMISSAO_INDICADOR_N1");
                sqlString.AppendLine("INNER JOIN t_COMISSAO_INDICADOR_N2 ON (t_COMISSAO_INDICADOR_N1.id = t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1)");
                sqlString.AppendLine("INNER JOIN t_COMISSAO_INDICADOR_N3 ON (t_COMISSAO_INDICADOR_N2.id = t_COMISSAO_INDICADOR_N3.id_comissao_indicador_n2)");
                sqlString.AppendLine("INNER JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_COMISSAO_INDICADOR_N3.indicador = t_ORCAMENTISTA_E_INDICADOR.apelido)");
                sqlString.AppendLine("WHERE (t_COMISSAO_INDICADOR_N1.id = " + id + ") AND (t_COMISSAO_INDICADOR_N3.st_tratamento_manual=0) AND (t_COMISSAO_INDICADOR_N3.vl_total_pagto > 0)");
                sqlString.AppendLine("ORDER BY t_COMISSAO_INDICADOR_N3.indicador");

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandText = sqlString.ToString();
                IDataReader reader = cmd.ExecuteReader();

                try
                {
                    int idxFavorecido = reader.GetOrdinal("favorecido");
                    int idxCnpjCpf = reader.GetOrdinal("favorecido_cnpj_cpf");
                    int idxBanco = reader.GetOrdinal("banco");
                    int idxAgencia = reader.GetOrdinal("agencia");
                    int idxAgencia_dv = reader.GetOrdinal("agencia_dv");
                    int idxConta = reader.GetOrdinal("conta");
                    int idxConta_dv = reader.GetOrdinal("conta_dv");
                    int idxOperacao = reader.GetOrdinal("conta_operacao");
                    int idxTipoConta = reader.GetOrdinal("tipo_conta");
                    int idxVlTotal = reader.GetOrdinal("vl_total_pagto");
                    int idxVendedor = reader.GetOrdinal("vendedor");


                    while (reader.Read())
                    {
                        Indicador _novo = new Indicador();
                        _novo.Nome = reader.GetString(idxFavorecido);
                        _novo.CpfCnpj = reader.IsDBNull(idxCnpjCpf) ? "" : reader.GetString(idxCnpjCpf);
                        _novo.Banco = reader.IsDBNull(idxBanco) ? "" : reader.GetString(idxBanco);
                        _novo.Agencia = reader.IsDBNull(idxAgencia) ? "" : reader.GetString(idxAgencia);
                        _novo.DigitoAgencia = reader.IsDBNull(idxAgencia_dv) ? "" : reader.GetString(idxAgencia_dv);
                        _novo.Conta = reader.IsDBNull(idxConta) ? "" : reader.GetString(idxConta);
                        _novo.DigitoConta = reader.IsDBNull(idxConta_dv) ? "" : reader.GetString(idxConta_dv);
                        _novo.Operacao = reader.IsDBNull(idxOperacao) ? "" : reader.GetString(idxOperacao);
                        _novo.Valor = reader.GetDecimal(idxVlTotal);
                        _novo.TipoConta = reader.IsDBNull(idxTipoConta) ? "" : reader.GetString(idxTipoConta);
                        _novo.Vendedor = reader.GetString(idxVendedor);

                        if ((_novo.CpfCnpj.Length) == 11)
                        {
                            _novo.TipoDocumento = "1";
                            _novo.TipoPessoa = "F";
                        }
                        else if ((_novo.CpfCnpj.Length) == 14)
                        {
                            _novo.TipoDocumento = "2";
                            _novo.TipoPessoa = "J";
                        }

                        if (_novo.TipoConta.Equals("C"))
                        {
                            _novo.TipoContaCSV = "CC";
                        }
                        else if (_novo.TipoConta.Equals("P"))
                        {
                            _novo.TipoContaCSV = "CP";
                        }

                        listaInd.Add(_novo);

                    }
                }
                finally
                {
                    reader.Close();
                }
            }
            finally
            {
                cn.Close();
            }

            return listaInd.ToArray();
        }


        public Empresa GetEmpresa()
        {
            Empresa empresa = new Empresa();
            SqlConnection cn = new SqlConnection(BD.getConnectionString());
            cn.Open();

            try
            {

                StringBuilder sqlString = new StringBuilder();

                sqlString.AppendLine("SELECT id, campo_texto FROM t_PARAMETRO WHERE id LIKE 'Ciagroup%'");

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandText = sqlString.ToString();
                IDataReader reader = cmd.ExecuteReader();

                try
                {
                    string id = "";

                    while (reader.Read())
                    {
                        id = reader.GetString(0);

                        if (id.Equals("Ciagroup_Cedente_1__CEP"))
                            empresa.Cep = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__Cidade"))
                            empresa.Cidade = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__CNPJ"))
                            empresa.Cnpj = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__Contato"))
                            empresa.Contato = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__Email"))
                            empresa.Email = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__Endereco"))
                            empresa.Endereco = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__NomeFantasia"))
                            empresa.NomeFantasia = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__RazaoSocial"))
                            empresa.RazaoSocial = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__TaxaAdm"))
                            empresa.TaxaAdm = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__Telefone"))
                            empresa.Telefone = reader.GetString(1);
                        if (id.Equals("Ciagroup_Cedente_1__UF"))
                            empresa.Uf = reader.GetString(1);
                    }
                }
                finally
                {
                    reader.Close();
                }
            }
            finally
            {
                cn.Close();
            }

            return empresa;
        }
    }
}