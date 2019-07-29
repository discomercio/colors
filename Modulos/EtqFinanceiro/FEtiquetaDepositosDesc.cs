#region [ using ]
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Media;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
#endregion

namespace EtqFinanceiro
{
    public partial class FEtiquetaDepositosDesc : FModelo
    {
        #region [ Constantes ]
        const string GRID_PESQ_COL_CHECKBOX = "colGrdPesqCheckBox";
        const string GRID_PESQ_COL_DATA_HORA = "colGrdPesqDataHora";
        const string GRID_PESQ_COL_VENDEDORES = "colGrdPesqVendedores";
        const string GRID_PESQ_COL_ID = "colGrdPesqId";
        const string GRID_DADOS_COL_CHECKBOX = "colGrdDadosCheckBox";
        const string GRID_DADOS_COL_VENDEDOR = "colGrdDadosVendedor";
        const string GRID_DADOS_COL_INDICADOR = "colGrdDadosIndicador";
        const string GRID_DADOS_COL_BANCO = "colGrdDadosBanco";
        const string GRID_DADOS_COL_AGENCIA = "colGrdDadosAgencia";
        const string GRID_DADOS_COL_CONTA = "colGrdDadosConta";
        const string GRID_DADOS_COL_FAVORECIDO = "colGrdDadosFavorecido";
        const string GRID_DADOS_COL_VALOR = "colGrdDadosValor";
        const string GRID_DADOS_COL_TIPO_CONTA = "colGrdDadosTipoConta";
        const string GRID_DADOS_COL_CONTA_OPERACAO = "colGrdDadosContaOperacao";
        #endregion

        #region [ Atributos ]
        private bool _emProcessamento = false;
        private bool _InicializacaoOk;

        public bool inicializacaoOk
        {
            get { return _InicializacaoOk; }
        }

        private bool _OcorreuExceptionNaInicializacao;

        public bool ocorreuExceptionNaInicializacao
        {
            get { return _OcorreuExceptionNaInicializacao; }
        }

        private DataTable _dtbConsulta = new DataTable();
        private int _idUltimoSelecionado = 0;
        private DateTime _dtHrIdUltimoSelecionado = DateTime.MinValue;
        private List<EtiquetaDepositosDescDados> _listaEtqCompleta = new List<EtiquetaDepositosDescDados>();
        private string bancosAIgnorarConfig = ConfigurationManager.AppSettings["bancosAIgnorar"];
        #endregion

        #region [ Impressão ]
        const char CODIGO_SOH = (char)0x01;
        const char CODIGO_STX = (char)0x02;
        const char CODIGO_CR = (char)0x0D;
        #endregion

        #region [ Construtor ]
        public FEtiquetaDepositosDesc()
        {
            InitializeComponent();
        }
        #endregion

        #region [ Métodos Privados ]

        #region [ executaConsulta ]
        private bool executaConsulta()
        {
            #region [ Declarações ]
            const string NOME_DESTA_ROTINA = "executaConsulta()";
            int idSelecionado = 0;
            string strSql;
            string strMsgErro = "";
            string strAux = "";
            SqlCommand cmCommand;
            SqlDataAdapter daAdapter;
            DataRow rowConsulta;
			EtiquetaDepositosDescDados etiqueta;
            string[] bancosAIgnorar = this.bancosAIgnorarConfig.Split(';');
            #endregion

            try
            {
                info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

                #region [ Há linha selecionada? ]
                for (int i = 0; i < grdPesquisa.Rows.Count; i++)
                {
                    if (grdPesquisa.Rows[i].Selected)
                    {
                        idSelecionado = (int)Global.converteInteiro(grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_ID].Value.ToString());
                        break;
                    }
                }

                if (idSelecionado == 0)
                {
                    avisoErro("Nenhum relatório foi selecionado!!");
                    return false;
                }
                #endregion

                #region [ É a mesma consulta que a anterior? ]
                if ((idSelecionado == _idUltimoSelecionado) && (grdDados.Rows.Count > 0))
                {
                    if (Global.calculaTimeSpanSegundos(DateTime.Now - _dtHrIdUltimoSelecionado) <= 10)
                        return true;
                }
                #endregion

                #region [ Limpa listas que armazenam os dados p/ impressão das etiquetas ]
                _listaEtqCompleta.Clear();
                #endregion

                #region [ Cria objetos de BD ]
                cmCommand = BD.criaSqlCommand();
                daAdapter = BD.criaSqlDataAdapter();
                #endregion

                #region [ Monta restrições de bancos ]
                for (int i = 0; i < bancosAIgnorar.Length; i++)
                {
                    if (strAux.Length > 0) strAux += ", ";
                    strAux += "'" + bancosAIgnorar[i] + "'";
                }
                strAux = " AND(banco NOT IN (" + strAux + "))";
                #endregion

                #region [ Monta o SQL ]
                strSql = "SELECT tn1.id" +
                               " ,vendedor" +
                               " ,indicador" +
                               " ,meio_pagto" +
                               " ,banco" +
                               " ,agencia" +
                               " ,agencia_dv" +
                               " ,conta" +
                               " ,conta_dv" +
                               " ,favorecido" +
							   " ,vl_total_pagto_liquido_NFS AS vl_total_pagto" +
                               " ,tipo_conta" +
                               " ,conta_operacao" +
                            " FROM t_COMISSAO_INDICADOR_N3 tn3" +
                            " INNER JOIN t_COMISSAO_INDICADOR_N2 tn2 ON(tn2.id = tn3.id_comissao_indicador_n2)" +
                            " INNER JOIN t_COMISSAO_INDICADOR_N1 tn1 ON(tn1.id = tn2.id_comissao_indicador_n1)" +
                            " WHERE(tn1.id = '" + idSelecionado.ToString() + "')" +
                               " AND(tn3.st_tratamento_manual = 0)" +
                               " AND(tn3.vl_total_pagto > 0)" +
                               strAux +
                               " ORDER BY vendedor, indicador";

                #endregion

                #region [ Executa a consulta no BD ]
                _dtbConsulta.Reset();
                cmCommand.CommandText = strSql;
                daAdapter.SelectCommand = cmCommand;
                daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
                daAdapter.Fill(_dtbConsulta);
                #endregion

                #region [ Verifica se retornou algum registro ]
                if(_dtbConsulta.Rows.Count == 0)
                {
                    grdDados.Rows.Clear();
                    info(ModoExibicaoMensagemRodape.Normal);
                    avisoErro("Nenhum depósito para impressão foi encontrado!!");
                    return false;
                }
                #endregion

                #region [ Carrega dados no grid ]
                try
                {
                    info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");

                    grdDados.Rows.Clear();
                    if (_dtbConsulta.Rows.Count > 0) grdDados.Rows.Add(_dtbConsulta.Rows.Count);

                    for (int i = 0; i < _dtbConsulta.Rows.Count; i++)
                    {
                        rowConsulta = _dtbConsulta.Rows[i];

                        etiqueta = new EtiquetaDepositosDescDados();
                        etiqueta.Id = BD.readToInt(rowConsulta["id"]);
                        etiqueta.Vendedor = BD.readToString(rowConsulta["vendedor"]);
                        etiqueta.Indicador = BD.readToString(rowConsulta["indicador"]);
                        etiqueta.Banco = BD.readToString(rowConsulta["banco"]) + " - " + Global.filtraAcentuacao(Texto.iniciaisEmMaiusculas(BD.getBancoDescricao(BD.readToString(rowConsulta["banco"]), out strMsgErro)));
                        etiqueta.Agencia = BD.readToString(rowConsulta["agencia"]) + (BD.readToString(rowConsulta["agencia_dv"]) != "" ? "-" + BD.readToString(rowConsulta["agencia_dv"]) : "");
                        etiqueta.Conta = (BD.readToString(rowConsulta["conta_operacao"]) != "" ? BD.readToString(rowConsulta["conta_operacao"]) + "-" : "") + BD.readToString(rowConsulta["conta"]) + (BD.readToString(rowConsulta["conta_dv"]) != "" ? "-" + BD.readToString(rowConsulta["conta_dv"]) : ""); ;
                        etiqueta.Favorecido = BD.readToString(rowConsulta["favorecido"]);
                        etiqueta.Valor = BD.readToDecimal(rowConsulta["vl_total_pagto"]);
                        etiqueta.MeioPagto = BD.readToString(rowConsulta["meio_pagto"]);
                        etiqueta.TipoConta = BD.readToString(rowConsulta["tipo_conta"]);
                        etiqueta.ContaOperacao = BD.readToString(rowConsulta["conta_operacao"]);

                        _listaEtqCompleta.Add(etiqueta);

                        preencheLinhaGrid(i, etiqueta);
                    }

                    #region [ Exibe o grid sem nenhuma linha pré-selecionada ]
                    grdDados.ClearSelection();
                    #endregion
                }
                finally
                {
                    grdDados.ResumeLayout();
                }
                #endregion

                #region [ Totais ]
                lblTotalRegistros.Text = Global.formataInteiro(_dtbConsulta.Rows.Count);
                #endregion

                grdDados.Focus();

                SystemSounds.Exclamation.Play();

                _dtHrIdUltimoSelecionado = DateTime.Now;
                _idUltimoSelecionado = idSelecionado;

                return true;
            }
            catch (Exception ex)
            {
                Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
                avisoErro(ex.ToString());
                Close();
                return false;
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);
            }
        }
        #endregion

        #region [ executaPesquisa ]
        private bool executaPesquisa()
        {
            #region [ Declarações ]
            const string NOME_DESTA_ROTINA = "executaPesquisa()";
            string strSql;
            string strWhere = "";
            string strAno;
            string strMes;
            SqlCommand cmCommand;
            SqlDataAdapter daAdapter;
            DataTable dtbConsulta = new DataTable();
            DataRow rowConsulta;
            #endregion

            try
            {
                #region [ Limpa campos dos dados de resposta ]
                limpaCamposDados();
                #endregion

                #region [ Monta restrições da cláusula 'Where' ]
                if (!string.IsNullOrEmpty(txtMesCompetencia.Text))
                {
                    strAno = Global.digitos(txtMesCompetencia.Text.Substring(2));
                    strMes = Global.digitos(txtMesCompetencia.Text.Substring(0, 2));

                    if (strWhere.Length > 0) strWhere += " AND";
                    strWhere += " t_COMISSAO_INDICADOR_N2.competencia_ano = '" + strAno + "'" +
                        " AND t_COMISSAO_INDICADOR_N2.competencia_mes = '" + strMes + "'";
                }
                #endregion

                #region [ Há restrições definidas? ]
                if (strWhere.Length == 0)
                {
                    avisoErro("Informe o mês de competência!!");
                    txtMesCompetencia.Focus();
                    return false;
                }
                #endregion

                info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

                #region [ Cria objetos de BD ]
                cmCommand = BD.criaSqlCommand();
                daAdapter = BD.criaSqlDataAdapter();
                #endregion

                #region [ Monta o SQL ]
                strSql = "SELECT t_COMISSAO_INDICADOR_N1.id" +
                           " ,t_COMISSAO_INDICADOR_N1.dt_hr_cadastro" +
                           " ,t_COMISSAO_INDICADOR_N2.vendedor" +
                           " ,t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1" +
                       " FROM t_COMISSAO_INDICADOR_N1" +
                       " INNER JOIN t_COMISSAO_INDICADOR_N2 ON(t_COMISSAO_INDICADOR_N1.id = t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1)" +
                       " WHERE" +
                            strWhere +
                       " ORDER BY dt_hr_cadastro";
                #endregion

                #region [ Executa a consulta no BD ]
                cmCommand.CommandText = strSql;
                daAdapter.SelectCommand = cmCommand;
                daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
                daAdapter.Fill(dtbConsulta);
                #endregion

                #region [ Carrega dados no grid ]
                try
                {
                    info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
                    grdPesquisa.SuspendLayout();

                    grdPesquisa.Rows.Clear();
                    if (dtbConsulta.Rows.Count > 0) grdPesquisa.Rows.Add(dtbConsulta.Rows.Count);

                    for (int i = 0; i < dtbConsulta.Rows.Count; i++)
                    {
                        rowConsulta = dtbConsulta.Rows[i];
                        grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_DATA_HORA].Value = Global.formataDataDdMmYyyyHhMmSsComSeparador(BD.readToDateTime(rowConsulta["dt_hr_cadastro"]));
                        grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_VENDEDORES].Value = BD.readToString(rowConsulta["vendedor"]).ToString();
                        grdPesquisa.Rows[i].Cells[GRID_PESQ_COL_ID].Value = BD.readToString(rowConsulta["id"]).ToString();
                    }

                    #region [ Exibe o grid sem nenhuma linha pré-selecionada ]
                    grdPesquisa.ClearSelection();
                    #endregion
                }
                finally
                {
                    grdPesquisa.ResumeLayout();
                }
                #endregion

                grdPesquisa.Focus();

                SystemSounds.Exclamation.Play();

                return true;
            }
            catch (Exception ex)
            {
                Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
                avisoErro(ex.ToString());
                Close();
                return false;
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);
            }
        }
        #endregion

        #region [ montaDadosImpressaoEtiqueta ]
        private bool montaDadosImpressaoEtiqueta(EtiquetaDepositosDescDados etiqueta, out string textoImpressaoEtiqueta, out string strMsgErro)
        {
            #region [ Declarações ]
            const string NOME_DESTA_ROTINA = "montaDadosImpressaoEtiqueta()";
            string strVendedor;
            string strIndicador;
            string strBanco;
            string strAgencia;
            string strConta;
            string strFavorecido;
            string strValor;
            string strTipoConta;
            const string COORDENADAS_MARGEM_X = "0045";
            #region [ Padrão ]
            const string FONTE_PADRAO = "9";
            const string SUB_FONTE_PADRAO = "004";
            const string MULTIPLICADOR_HORIZONTAL_PADRAO = "1";
            const string MULTIPLICADOR_VERTICAL_PADRAO = "1";
            #endregion
            #region [ Reduzida ]
            const string FONTE_REDUZIDA = "9";
            const string SUB_FONTE_REDUZIDA = "002";
            const string MULTIPLICADOR_HORIZONTAL_REDUZIDA = "1";
            const string MULTIPLICADOR_VERTICAL_REDUZIDA = "1";
            #endregion
            #endregion

            textoImpressaoEtiqueta = "";
            strMsgErro = "";

            if (etiqueta == null)
            {
                strMsgErro = "Não há dados para a impressão da etiqueta!!";
                return false;
            }

            strVendedor = Global.filtraAcentuacao(etiqueta.Vendedor);
            strIndicador = Global.filtraAcentuacao(etiqueta.Indicador);
            strBanco = etiqueta.Banco;
            strAgencia = etiqueta.Agencia;
            strConta = etiqueta.Conta;
            strFavorecido = Texto.iniciaisEmMaiusculas(Global.filtraAcentuacao(etiqueta.Favorecido));
            strValor = Global.formataMoeda(etiqueta.Valor);
            strTipoConta = etiqueta.TipoConta;

            try
            {
                textoImpressaoEtiqueta =
                    CODIGO_STX + "L" + CODIGO_CR + // Enters label formatting state
                    "H12" + CODIGO_CR + // Temperatura da cabeça p/ controlar o contraste (padrão: H10, máximo: H20, máximo recomendável: H16)
                    "D11" + CODIGO_CR + // Sets width and height pixel size (default: D22)
                                        // Valor
                    "1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape 
                    FONTE_PADRAO + // Fonte
                    MULTIPLICADOR_HORIZONTAL_PADRAO + // Multiplicador Horizontal
                    MULTIPLICADOR_VERTICAL_PADRAO + // Multiplicador Vertical
                    SUB_FONTE_PADRAO + // Subtipo da fonte
                    "0020" + // Coordenadas Y
                    "0580" + // Coordenadas X
                    Global.Cte.Etc.SIMBOLO_MONETARIO + " " + strValor + CODIGO_CR + // Texto a ser impresso
                                                   // Favorecido
                    "1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
                    FONTE_REDUZIDA + // Fonte
                    MULTIPLICADOR_HORIZONTAL_REDUZIDA + // Multiplicador Horizontal
                    MULTIPLICADOR_VERTICAL_REDUZIDA + // Multiplicador Vertical
                    SUB_FONTE_REDUZIDA + // Subtipo da fonte
                    "0070" + // Coordenadas Y
                    COORDENADAS_MARGEM_X + // Coordenadas X
                    "Favorecido: " + strFavorecido + CODIGO_CR + // Texto a ser impresso
                                                                 // Agência/Conta
                    "1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
                    FONTE_REDUZIDA + // Fonte
                    MULTIPLICADOR_HORIZONTAL_REDUZIDA + // Multiplicador Horizontal
                    MULTIPLICADOR_VERTICAL_REDUZIDA + // Multiplicador Vertical
                    SUB_FONTE_REDUZIDA + // Subtipo da fonte
                    "0120" + // Coordenadas Y
                    COORDENADAS_MARGEM_X + // Coordenadas X
                    "Agencia: " + strAgencia + "          " + (strTipoConta != "" ? "C/" + strTipoConta : "Conta") + ": " + strConta + CODIGO_CR + // Texto a ser impresso
                                                                                                                                             // Banco
                    "1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
                    FONTE_REDUZIDA + // Fonte
                    MULTIPLICADOR_HORIZONTAL_REDUZIDA + // Multiplicador Horizontal
                    MULTIPLICADOR_VERTICAL_REDUZIDA + // Multiplicador Vertical
                    SUB_FONTE_REDUZIDA + // Subtipo da fonte
                    "0170" + // Coordenadas Y
                    COORDENADAS_MARGEM_X + // Coordenadas X
                    "Banco: " + strBanco + CODIGO_CR + // Texto a ser impresso
                                                       // Indicador/Vendedor
                    "1" + // Orientação: 1=Portrait, 2=Reverse Landscape, 3=Reverse Portrait, 4=Landscape
                    FONTE_REDUZIDA + // Fonte
                    MULTIPLICADOR_HORIZONTAL_REDUZIDA + // Multiplicador Horizontal
                    MULTIPLICADOR_VERTICAL_REDUZIDA + // Multiplicador Vertical
                    SUB_FONTE_REDUZIDA + // Subtipo da fonte
                    "0220" + // Coordenadas Y
                    COORDENADAS_MARGEM_X + // Coordenadas X
                    "Indicador: " + strIndicador + "               Vendedor: " + strVendedor + CODIGO_CR + // Texto a ser impresso
                                                                                                                // Comandos de finalização da etiqueta
                    "Q0001" + CODIGO_CR + // Sets the quantity of labels to print
                    "E" + CODIGO_CR // Ends the job and exit from label formatting mode
                    ;

                return true;
            }
            catch (Exception ex)
            {
                strMsgErro = ex.ToString();
                Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
                return false;
            }
        }
        #endregion

        #region [ limpaCamposPesquisa ]
        private void limpaCamposPesquisa()
        {
            grdPesquisa.Rows.Clear();
            txtMesCompetencia.Text = "";
        }
        #endregion

        #region [ limpaCamposDados ]
        private void limpaCamposDados()
        {
            grdDados.Rows.Clear();
            lblTotalRegistros.Text = "";
        }
        #endregion

        #region [ preencheLinhaGrid ]
        private bool preencheLinhaGrid(int rowIndex, EtiquetaDepositosDescDados etiqueta)
        {
            try
            {
                grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_VENDEDOR].Value = etiqueta.Vendedor;
                grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_INDICADOR].Value = etiqueta.Indicador;
                grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_BANCO].Value = etiqueta.Banco;
                grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_AGENCIA].Value = etiqueta.Agencia;
                grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_CONTA].Value = etiqueta.Conta;
                grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_FAVORECIDO].Value = etiqueta.Favorecido;
                grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_VALOR].Value = Global.formataMoeda(etiqueta.Valor);
                grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_TIPO_CONTA].Value = etiqueta.TipoConta;
                grdDados.Rows[rowIndex].Cells[GRID_DADOS_COL_CONTA_OPERACAO].Value = etiqueta.ContaOperacao;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        #endregion

        #region [ trataBotaoPrinterDialog ]
        private void trataBotaoPrinterDialog()
        {
            printDialog.ShowDialog();
        }
        #endregion

        #region [ trataBotaoConsultar ]
        private bool trataBotaoConsultar()
        {

            #region [ Verifica se a conexão c/ o BD está ok ]
            if (!BD.isConexaoOk())
            {
                if (!FMain.reiniciaBancoDados())
                {
                    avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
                    return false;
                }
            }
            #endregion

            return executaConsulta();
        }
        #endregion

        #region [ trataBotaoPesquisar ]
        private bool trataBotaoPesquisar()
        {

            #region [ Verifica se a conexão c/ o BD está ok ]
            if (!BD.isConexaoOk())
            {
                if (!FMain.reiniciaBancoDados())
                {
                    avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
                    return false;
                }
            }
            #endregion

            if (!executaPesquisa()) return false;

            if (grdPesquisa.Rows.Count == 1)
            {
                grdPesquisa.Rows[0].Selected = true;
                return executaConsulta();
            }

            return true;
        }
        #endregion

        #region [ trataBotaoImprimirCompleto ]
        private void trataBotaoImprimirCompleto()
        {
            #region [ Declarações ]
            const string NOME_DESTA_ROTINA = "trataBotaoImprimirCompleto()";
            string strDescricaoLog;
            string strMsgErroLog = "";
            string strMsgErro;
            string textoEtiqueta;
            string textoEtiquetaBatch;
            StringBuilder sbEtiqueta = new StringBuilder();
            Log log = new Log();
            #endregion

            #region [ Verifica se a conexão c/ o BD está ok ]
            if (!BD.isConexaoOk())
            {
                if (!FMain.reiniciaBancoDados())
                {
                    avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
                    return;
                }
            }
            #endregion

            try
            {
                if (_listaEtqCompleta.Count == 0)
                {
                    avisoErro("Não há dados para imprimir!!");
                    return;
                }

                if (grdDados.Rows.Count == 0)
                {
                    avisoErro("Escolha um relatório para consultar a relação de depósitos!!");
                    return;
                }

                #region [ Impressora correta? ]
                while (true)
                {
                    if (!printDialog.PrinterSettings.PrinterName.ToUpper().Contains("ARGOX"))
                    {
                        if (confirma("A impressora selecionada (" + printDialog.PrinterSettings.PrinterName + ") não parece ser a impressora de etiquetas!!\nDeseja selecionar outra impressora?"))
                        {
                            printDialog.ShowDialog();
                        }
                        else break;
                    }
                    else break;
                }
                #endregion

                #region [ Confirma impressão? ]
                if (!confirma("Confirma a impressão da listagem COMPLETA (" + Global.formataInteiro(_dtbConsulta.Rows.Count).ToString() + " etiquetas)?")) return;
                #endregion

                info(ModoExibicaoMensagemRodape.EmExecucao, "preparando dados das etiquetas");

                #region [ Monta os dados e imprime ]
                for (int i = 0; i < _listaEtqCompleta.Count; i++)
                {
                    if (!montaDadosImpressaoEtiqueta(_listaEtqCompleta[i], out textoEtiqueta, out strMsgErro))
                    {
                        throw new Exception(strMsgErro);
                    }

                    sbEtiqueta.Append(textoEtiqueta);
                }

                textoEtiquetaBatch =
                    CODIGO_STX + "m" + CODIGO_CR + // Sets measurement to metric
                    CODIGO_STX + "r" + CODIGO_CR + // Selects reflective sensor for gap
                    CODIGO_STX + "V0" + CODIGO_CR + // Sets cutter and dispenser configuration ('0': no cutter and peeler function; '1': Enables cutter function; '4': Enables peeler function)
                    CODIGO_SOH + "D" + CODIGO_CR + // Disables the interaction command
                    CODIGO_STX + "f920" + CODIGO_CR + // Sets stop position and automatic back-feed for the label stock (Back-feed will not be activated if xxx is less than 220)
                    sbEtiqueta.ToString();

                info(ModoExibicaoMensagemRodape.EmExecucao, "enviando dados para a impressora");

                RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, textoEtiquetaBatch);
                #endregion

                #region [ Grava log no BD ]
                strDescricaoLog = "Relatório de Relação de Depósitos: impressão completa das etiquetas do relatório de ID " + _idUltimoSelecionado.ToString() + " (" + _listaEtqCompleta.Count.ToString() + " etiquetas impressas)";
                log.usuario = Global.Usuario.usuario;
                log.operacao = Global.Cte.EtqFinanceiro.LogOperacao.ETIQUETA_FIN_IMPRESSAO_COMPLETA;
                log.complemento = strDescricaoLog;
                LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
                #endregion

                // Feedback da conclusão da consulta
                SystemSounds.Exclamation.Play();
            }
            catch (Exception ex)
            {
                strMsgErro = ex.ToString();
                avisoErro(strMsgErro);
                Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + strMsgErro);
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);
            }
        }
        #endregion

        #region [ trataBotaoImprimirSelecionado ]
        private void trataBotaoImprimirSelecionado()
        {
            #region [ Declarações ]
            const string NOME_DESTA_ROTINA = "trataBotaoImprimirSelecionado()";
            string strMsgErroLog = "";
            string strDescricaoLog;
            string strMsgErro;
            string textoEtiqueta;
            string textoEtiquetaBatch;
            StringBuilder sbEtiqueta = new StringBuilder();
            Log log = new Log();
            EtiquetaDepositosDescDados etiqueta = new EtiquetaDepositosDescDados();
            #endregion

            #region [ Verifica se a conexão c/ o BD está ok ]
            if (!BD.isConexaoOk())
            {
                if (!FMain.reiniciaBancoDados())
                {
                    avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!");
                    return;
                }
            }
            #endregion

            try
            {
                #region [ Recupera dados da linha selecionada ]
                for (int i = 0; i < grdDados.Rows.Count; i++)
                {
                    if (grdDados.Rows[i].Selected)
                    {
                        etiqueta.Vendedor = grdDados.Rows[i].Cells[GRID_DADOS_COL_VENDEDOR].Value.ToString();
                        etiqueta.Indicador = grdDados.Rows[i].Cells[GRID_DADOS_COL_INDICADOR].Value.ToString();
                        etiqueta.Banco = grdDados.Rows[i].Cells[GRID_DADOS_COL_BANCO].Value.ToString();
                        etiqueta.Agencia = grdDados.Rows[i].Cells[GRID_DADOS_COL_AGENCIA].Value.ToString();
                        etiqueta.Conta = grdDados.Rows[i].Cells[GRID_DADOS_COL_CONTA].Value.ToString();
                        etiqueta.Favorecido = grdDados.Rows[i].Cells[GRID_DADOS_COL_FAVORECIDO].Value.ToString();
                        etiqueta.Valor = Convert.ToDecimal(grdDados.Rows[i].Cells[GRID_DADOS_COL_VALOR].Value);
                        etiqueta.TipoConta = grdDados.Rows[i].Cells[GRID_DADOS_COL_TIPO_CONTA].Value.ToString();
                        break;
                    }
                }

                if (string.IsNullOrEmpty(etiqueta.Vendedor))
                {
                    avisoErro("Nenhum dado foi selecionado para impressão!!");
                    return;
                }
                #endregion

                #region [ Verifica se a impressora selecionada é a correta ]
                while (true)
                {
                    if (!printDialog.PrinterSettings.PrinterName.ToUpper().Contains("ARGOX"))
                    {
                        if (confirma("A impressora selecionada (" + printDialog.PrinterSettings.PrinterName + ") não parece ser a impressora de etiquetas!!\nDeseja selecionar outra impressora?"))
                        {
                            printDialog.ShowDialog();
                        }
                        else break;
                    }
                    else break;
                } 
                #endregion

                if (!confirma("Confirma a impressão da etiqueta selecionada (" + etiqueta.Indicador + ")?")) return;

                info(ModoExibicaoMensagemRodape.EmExecucao, "preparando dados da etiqueta");

                #region [ Monta os dados e imprime ]
                if (!montaDadosImpressaoEtiqueta(etiqueta, out textoEtiqueta, out strMsgErro))
                {
                    throw new Exception(strMsgErro);
                }

                textoEtiquetaBatch =
                    CODIGO_STX + "m" + CODIGO_CR + // Sets measurement to metric
                    CODIGO_STX + "r" + CODIGO_CR + // Selects reflective sensor for gap
                    CODIGO_STX + "V0" + CODIGO_CR + // Sets cutter and dispenser configuration ('0': no cutter and peeler function; '1': Enables cutter function; '4': Enables peeler function)
                    CODIGO_SOH + "D" + CODIGO_CR + // Disables the interaction command
                    CODIGO_STX + "f920" + CODIGO_CR + // Sets stop position and automatic back-feed for the label stock (Back-feed will not be activated if xxx is less than 220)
                    textoEtiqueta;

                info(ModoExibicaoMensagemRodape.EmExecucao, "enviando dados para a impressora");

                RawPrinterHelper.SendStringToPrinter(printDialog.PrinterSettings.PrinterName, textoEtiquetaBatch);
                #endregion

                #region [ Grava log no BD ]
                strDescricaoLog = "Relatório de Relação de Depósitos: impressão avulsa de etiqueta do relatório de ID " + _idUltimoSelecionado.ToString() + " (indicador " + etiqueta.Indicador + ")";
                log.usuario = Global.Usuario.usuario;
                log.operacao = Global.Cte.EtqFinanceiro.LogOperacao.ETIQUETA_FIN_IMPRESSAO_SELECIONADO;
                log.complemento = strDescricaoLog;
                LogDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
                #endregion

                // Feedback da conclusão da consulta
                SystemSounds.Exclamation.Play();
            }
            catch (Exception ex)
            {
                strMsgErro = ex.ToString();
                avisoErro(strMsgErro);
                Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + strMsgErro);
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);
            }
        }
		#endregion

		#endregion

		#region [ Eventos ]

		#region [ FEtiquetaDepositosDescDados ]

		#region [ EFtiquetaDepositosDescDados ]
		private void EFtiquetaDepositosDescDados_Load(object sender, EventArgs e)
        {
            bool blnSucesso = false;

            try
            {
                limpaCamposPesquisa();
                limpaCamposDados();
                //limpaCamposFiltro();

                blnSucesso = true;
            }
            catch (Exception ex)
            {
                _OcorreuExceptionNaInicializacao = true;
                avisoErro(ex.ToString());
                return;
            }
            finally
            {
                if (!blnSucesso) Close();
            }
        }
        #endregion

        #region [ FEtiquetaDepositosDescDados_Shown ]
        private void FEtiquetaDepositosDescDados_Shown(object sender, EventArgs e)
        {
            try
            {
                #region [ Executa rotinas de inicialização ]
                if (!_InicializacaoOk)
                {
                    #region [ Ajusta layout do header do grid (resultado da pesquisa ) ]
                    grdPesquisa.Columns[GRID_PESQ_COL_CHECKBOX].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    grdPesquisa.Columns[GRID_PESQ_COL_DATA_HORA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    grdPesquisa.Columns[GRID_PESQ_COL_DATA_HORA].ReadOnly = true;
                    grdPesquisa.Columns[GRID_PESQ_COL_VENDEDORES].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    grdPesquisa.Columns[GRID_PESQ_COL_VENDEDORES].ReadOnly = true;
                    #endregion

                    #region [ Ajusta o layout do header do grid (dados de depósitos) ]
                    grdDados.ReadOnly = true;
                    grdDados.Columns[GRID_DADOS_COL_CHECKBOX].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    grdDados.Columns[GRID_DADOS_COL_VENDEDOR].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    grdDados.Columns[GRID_DADOS_COL_INDICADOR].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    grdDados.Columns[GRID_DADOS_COL_BANCO].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    grdDados.Columns[GRID_DADOS_COL_AGENCIA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    grdDados.Columns[GRID_DADOS_COL_CONTA].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    grdDados.Columns[GRID_DADOS_COL_FAVORECIDO].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    grdDados.Columns[GRID_DADOS_COL_VALOR].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                    #endregion

                    #region [ Posiciona foco ]
                    btnDummy.Focus();
                    #endregion

                    _InicializacaoOk = true;
                }
                #endregion
            }
            catch (Exception ex)
            {
                _OcorreuExceptionNaInicializacao = true;
                avisoErro(ex.ToString());
                Close();
                return;
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);
                if (!_InicializacaoOk) Close();
            }
        }
		#endregion

		#region [ FEtiquetaImprime_FormClosing ]
		private void FEtiquetaImprime_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_emProcessamento)
            {
                SystemSounds.Exclamation.Play();
                e.Cancel = true;
                return;
            }

            FMain.fMain.Location = this.Location;
            FMain.fMain.Visible = true;
            this.Visible = false;
        }
        #endregion

        #endregion

        #region [ txtMesCompetencia ]

        #region [ txtMesCompetencia_Enter ]
        private void txtMesCompetencia_Enter(object sender, EventArgs e)
        {
            txtMesCompetencia.Select(0, txtMesCompetencia.Text.Length);
        }
        #endregion

        #region [ txtMesCompetencia_KeyDown ]
        private void txtMesCompetencia_KeyDown(object sender, KeyEventArgs e)
        {
            Global.trataTextBoxKeyDown(sender, e, btnPesquisar);
        }
        #endregion

        #region [ txtMesCompetencia_KeyPress ]
        private void txtMesCompetencia_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Global.filtraDigitacaoData(e.KeyChar);
        }
        #endregion

        #region [ txtMesCompetencia_Leave ]
        private void txtMesCompetencia_Leave(object sender, EventArgs e)
        {
            txtMesCompetencia.Text = Global.formataDataDigitadaParaMMYYYYComSeparador(txtMesCompetencia.Text);
            if (!Global.isDataMMYYYYOk(txtMesCompetencia.Text))
            {
                avisoErro("Formato inválido!!");
                txtMesCompetencia.Focus();
                return;
            }
        }
        #endregion

        #endregion

        #region [ btnPrinterDialog ]

        #region [ btnPrinterDialog_Click ]
        private void btnPrinterDialog_Click(object sender, EventArgs e)
        {
            trataBotaoPrinterDialog();
        }
        #endregion

        #endregion

        #region [ btnConsultar ]

        #region [ btnConsultar_Click ]
        private void btnConsultar_Click(object sender, EventArgs e)
        {
            trataBotaoConsultar();
        }
        #endregion

        #endregion

        #region [ btnPesquisar ]

        #region [ btnPesquisar_Click ]
        private void btnPesquisar_Click(object sender, EventArgs e)
        {
            trataBotaoPesquisar();
        }
        #endregion

        #endregion

        #region [ btnImprimirCompleto ]

        #region [ btnImprimirCompleto_Click ]
        private void btnImprimirCompleto_Click(object sender, EventArgs e)
        {
            trataBotaoImprimirCompleto();
        }
        #endregion

        #endregion

        #region [ btnImprimirSelecionado ]

        #region [ btnImprimirSelecionado_Click ]
        private void btnImprimirSelecionado_Click(object sender, EventArgs e)
        {
            trataBotaoImprimirSelecionado();
        }
        #endregion

        #endregion

        #region [ grdPesquisa ]

        #region [ grdPesquisa_CellContentClick ]
        private void grdPesquisa_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e == null) return;
            if (e.ColumnIndex == 0)
            {
                DataGridViewCheckBoxCell chkBox = (DataGridViewCheckBoxCell)this.grdPesquisa[e.ColumnIndex, e.RowIndex];
                if (chkBox.EditingCellFormattedValue.ToString().ToUpper().Equals("TRUE"))
                {
                    this.grdPesquisa.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightSkyBlue;
                }
                else
                {
                    this.grdPesquisa.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Empty;
                }
            }
        }
        #endregion

        #region [ grdPesquisa_CellDoubleClick ]
        private void grdPesquisa_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            trataBotaoConsultar();
        }
        #endregion

        #endregion

        #endregion

    }

    #region [ Declaração de Classes Auxiliares ]

    #region [ EtiquetaDepositosDescDados ]
    public class EtiquetaDepositosDescDados
    {
        public int Id { get; set; }

        public string Indicador { get; set; }

        public string Banco { get; set; }

        public string Agencia { get; set; }

        public string Conta { get; set; }

        public string Favorecido { get; set; }

        public string Vendedor { get; set; }

        public decimal Valor { get; set; }

        public string MeioPagto { get; set; }

        public string TipoConta { get; set; }

        public string ContaOperacao { get; set; }
    }
    #endregion

    #endregion
}
