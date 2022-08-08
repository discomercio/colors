#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
#endregion

namespace Financeiro
{
    public partial class FPlanilhaPagamentoMarketplaceB2W : FModelo
    {
        #region [ Atributos ]
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

        private string[] _linhasArqPlanilha;
        DataTable dtResultado;
        #endregion

        #region [ Construtor ]
        public FPlanilhaPagamentoMarketplaceB2W()
        {
            InitializeComponent();

            dtResultado = new DataTable("Resultado");
            dtResultado.Columns.Add("Linha", typeof(int));
            dtResultado.Columns.Add("Pedido", typeof(string));
            dtResultado.Columns.Add("DataPedido", typeof(string));
            dtResultado.Columns.Add("ValorTotalPedido", typeof(string));
            dtResultado.Columns.Add("Tipo", typeof(string));
            dtResultado.Columns.Add("Valor", typeof(string));
            dtResultado.Columns.Add("Observacao", typeof(string));
        }
        #endregion

        #region [ Métodos ]

        #region [ limpaCampos ]
        private void limpaCampos()
        {
            txtPlanilha.Clear();
            gridDados.DataSource = null;
        }
        #endregion

        #region [ trataBotaoSelecionaArqPlanilha ]
        private void trataBotaoSelecionaArqPlanilha()
        {
            #region [ Declarações ]
            DialogResult dr;
            #endregion
            
            dr = openFileDialog.ShowDialog();
            if (dr != DialogResult.OK) return;

            #region [ Limpa campos ]
            limpaCampos();
            #endregion

            txtPlanilha.Text = openFileDialog.FileName;
        }
        #endregion

        #region [ trataBotaoConfirmar ]
        private void trataBotaoConfirmar()
        {
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "FPlanilhaPagamentoMarketplaceB2W.trataBotaoConfirmar()";
			bool blnAux;
            bool blnPedidoExiste;
            string strMsgErro = "";
            string strPedidoAux;
            string strValorPedidoAux;
            int lojaOrigemIdAux;
            int meioPagtoAux;
            int intTipoTransacaoId;
            Dictionary<string, int> LojaOrigemLista = new Dictionary<string, int>();
            Dictionary<string, int> MeioPagamentoLista = new Dictionary<string, int>();
            Dictionary<string, int> TipoPagamentoLista = new Dictionary<string, int>();
            PlanilhaRepasseMktplaceN1 planilhaN1 = new PlanilhaRepasseMktplaceN1();
            PlanilhaRepasseMktplaceN2 planilhaN2 = null;
            PlanilhaRepasseMktplaceN3 planilhaN3 = null;
            PlanilhaRepasseMktplaceN4 planilhaN4;
            List<PlanilhaRepasseMktplaceN2> listPlanilhaN2 = new List<PlanilhaRepasseMktplaceN2>();
            List<PlanilhaRepasseMktplaceN3> listPlanilhaN3 = new List<PlanilhaRepasseMktplaceN3>();
            List<PlanilhaRepasseMktplaceN4> listPlanilhaN4 = new List<PlanilhaRepasseMktplaceN4>();
            FileInfo fileInfo;
            Encoding encode = Encoding.GetEncoding("Windows-1252");
            #endregion

            #region [ Selecionou algum arquivo? ]
            if (txtPlanilha.Text.Trim().Length == 0)
            {
                avisoErro("Selecione um arquivo de planilha de pagamentos!");
                btnSelecionaArqPlanilha.Focus();
                return;
            } 
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

            BD.iniciaTransacao();

            #region [ Rotinas de carregamento do arquivo ]
            try
            {
                info(ModoExibicaoMensagemRodape.EmExecucao, "gravando informações do arquivo");

                #region [ Captura informações sobre o arquivo selecionado ]
                fileInfo = new FileInfo(openFileDialog.FileName);

                #region [ Selecionou arquivo com extensão CSV? ]
                if (!Equals(fileInfo.Extension.ToLower(), ".csv"))
                {
                    BD.rollbackTransacao();
                    avisoErro("Selecione um arquivo com extensão .CSV!");
                    btnSelecionaArqPlanilha.Focus();
                    return;
                }
                #endregion

                planilhaN1.NomeArquivo = fileInfo.Name;
                planilhaN1.Path = fileInfo.DirectoryName;
                planilhaN1.OrigemGrupo = "B2W";
                // calcula checksum do arquivo
                planilhaN1.Checksum = Global.GetChecksum(txtPlanilha.Text);
                #endregion

                #region [ Verifica se o arquivo já foi processado ]
                //Classe Tuple
                //item1 (boolean): retorna true se o arquivo já foi processado
                //item2 (string): retorna o usuario que processou o arquivo (caso já foi processado)
                //item3 (string): retorna a data em que foi processado o arquivo (caso já foi processado)
                var retornoArquivoProcessado = PlanilhaRepasseMktplaceDAO.isArquivoProcessado(planilhaN1.Checksum);
                if (retornoArquivoProcessado.Item1 == true)
                {
                    BD.rollbackTransacao();
                    info(ModoExibicaoMensagemRodape.Normal);
                    avisoErro("O arquivo selecionado já foi processado por " + retornoArquivoProcessado.Item2 + " em " + retornoArquivoProcessado.Item3);
                    return;
                }
                #endregion
                
                #region [ Grava t_MKTP_REPASSE_N1 ]
                blnAux = PlanilhaRepasseMktplaceDAO.insereN1(Global.Usuario.usuario, planilhaN1, ref strMsgErro);
                if (!blnAux)
                {
                    BD.rollbackTransacao();
                    info(ModoExibicaoMensagemRodape.Normal);
                    avisoErro(strMsgErro);
                    return;
                }
                #endregion

                #region [ Carrega dados do arquivo em array ]
                info(ModoExibicaoMensagemRodape.EmExecucao, "lendo dados do arquivo");
                _linhasArqPlanilha = File.ReadAllLines(txtPlanilha.Text, encode);
                #endregion

                info(ModoExibicaoMensagemRodape.Normal);
            }
            catch (IOException ioex)
            {
                BD.rollbackTransacao();
                Global.gravaLogAtividade("Falha ao tentar carregar planilha de repasse marketplace\n" + ioex.Message);
                info(ModoExibicaoMensagemRodape.Normal);
                avisoErro("Não foi possível carregar o arquivo.\n\nSe o arquivo estiver aberto, feche-o e tente carregá-lo novamente");
                return;
            }
            catch (Exception ex)
            {
                BD.rollbackTransacao();
                Global.gravaLogAtividade("Falha ao tentar carregar planilha de repasse marketplace\n" + ex.Message);
                info(ModoExibicaoMensagemRodape.Normal);
                avisoErro("Não foi possível carregar o arquivo.\n\n" + ex.Message);
                return;
            }
            #endregion

            info(ModoExibicaoMensagemRodape.EmExecucao, "gravando dados do arquivo no banco de dados");

            try
            {
                #region [ Percorre as linhas do array ]
                for (int i = 1; i < _linhasArqPlanilha.Length; i++)
                {
                    string[] Colunas = _linhasArqPlanilha[i].Split(';');

                    strPedidoAux = Colunas[9].Trim('"');
                    intTipoTransacaoId = MktplaceRepasseTipoTransacao.GetId(Colunas[10].Trim('"'), "B2W", ref strMsgErro);
                    planilhaN4 = new PlanilhaRepasseMktplaceN4();
                    planilhaN4.Linha = i + 1;
                    if (!listPlanilhaN2.Any(x => x.Pedido == strPedidoAux))
                    {
                        planilhaN2 = new PlanilhaRepasseMktplaceN2();
                        planilhaN2.MktplceRepasseN1Id = planilhaN1.Id;
                        planilhaN2.Linha = i + 1;
                        if ((intTipoTransacaoId == MktplaceRepasseTipoTransacao.Venda.GetId())||(intTipoTransacaoId == MktplaceRepasseTipoTransacao.EstornoVenda.GetId()))
                        {
                            strValorPedidoAux = Colunas[12].Trim('"').Replace("-", string.Empty);
                            planilhaN2.ValorTotalPedido = Global.converteNumeroDecimal(strValorPedidoAux);
                        }

                        planilhaN3 = new PlanilhaRepasseMktplaceN3();
                        planilhaN3.Linha = i + 1;
                        blnPedidoExiste = false;
                    }
                    else
                    {
                        blnPedidoExiste = true;
                    }

                    #region [ Percorre as colunas da linha atual ]
                    for (int j = 0; j < Colunas.Length; j++)
                    {
                        // Retira as aspas
                        Colunas[j] = Colunas[j].Trim('"');

                        #region [ Recupera valor das células conforme índice das colunas na planilha ]

                        #region [ Recupera valores das planilhas N2 e N3 ]
                        if (!blnPedidoExiste)
                        {
                            switch (j)
                            {
                                case 0: //Marca
                                    if (!LojaOrigemLista.ContainsKey(Colunas[j]))
                                    {
                                        lojaOrigemIdAux = PlanilhaRepasseMktplaceDAO.getLojaOrigemId(Colunas[j], "B2W");
                                        if (lojaOrigemIdAux > 0)
                                        {
                                            LojaOrigemLista.Add(Colunas[j], lojaOrigemIdAux);
                                        }
                                    }
                                    planilhaN2.LojaOrigemId = LojaOrigemLista[Colunas[j]];
                                    break;
                                case 2: //Data pedido
                                    planilhaN2.DataPedido = Global.converteDdMmYyyyParaDateTime(Colunas[j]);
                                    break;
                                case 9: //Pedido
                                    planilhaN2.Pedido = Colunas[j];
                                    break;
                                case 14: //Meio Pgto
                                    if (!MeioPagamentoLista.ContainsKey(Colunas[j]))
                                    {
                                        meioPagtoAux = PlanilhaRepasseMktplaceDAO.getMeioPagamentoId(Colunas[j], "B2W");
                                        if (meioPagtoAux > 0)
                                        {
                                            MeioPagamentoLista.Add(Colunas[j], meioPagtoAux);
                                        }
                                    }
                                    planilhaN2.MeioPagamentoId = MeioPagamentoLista[Colunas[j]];
                                    break;
                                default:
                                    break;
                            } 
                        }
                        #endregion

                        #region [ Recupera dados da planilha N4 ]
                        switch (j)
                        {
                            case 3: //Data Pagamento
                                planilhaN4.DataPagamento = Global.converteDdMmYyyyParaDateTime(Colunas[j]);
                                break;
                            case 4: // Data Estorno
                                planilhaN4.DataEstorno = Global.converteDdMmYyyyParaDateTime(Colunas[j]);
                                break;
                            case 5: //Data Liberação
                                planilhaN4.DataLiberacao = Global.converteDdMmYyyyParaDateTime(Colunas[j]);
                                break;
                            case 10: //Tipo
                                intTipoTransacaoId = MktplaceRepasseTipoTransacao.GetId(Colunas[j], "B2W", ref strMsgErro);
                                if(intTipoTransacaoId == 0)
                                {
                                    BD.rollbackTransacao();
                                    avisoErro(strMsgErro + "\n\nA operação não será concluída!");
                                    return;
                                }
                                planilhaN4.TipoTransacao = intTipoTransacaoId;
                                break;
                            case 11: //Status
                                planilhaN4.StatusTransacaoId = PlanilhaRepasseMktplaceDAO.getStatusTransacaoId(Colunas[j], "B2W");
                                break;
                            case 12: //Valor
                                planilhaN4.Valor = Global.converteNumeroDecimal(Colunas[j]);
                                break;
                            default:
                                break;
                        }
                        #endregion

                        #endregion
                    }
                    #endregion

                    if (!blnPedidoExiste)
                    {
                        blnAux = PlanilhaRepasseMktplaceDAO.insereN2(planilhaN2, ref strMsgErro);
                        if (!blnAux)
                        {
                            BD.rollbackTransacao();
                            avisoErro("Erro ao gravar linha específica da planilha no banco de dados!\n\nA operação será interrompida!\n\n" + strMsgErro);
                            return;
                        }
                        listPlanilhaN2.Add(planilhaN2);

                        planilhaN3.MktplceRepasseN2Id = planilhaN2.Id;
                        blnAux = PlanilhaRepasseMktplaceDAO.insereN3(planilhaN3, ref strMsgErro);
                        listPlanilhaN3.Add(planilhaN3);
                    }

                    planilhaN2 = listPlanilhaN2.Find(x => x.Pedido == strPedidoAux);
                    planilhaN3 = listPlanilhaN3.Find(x => x.MktplceRepasseN2Id == planilhaN2.Id);
                    planilhaN4.MktplceRepasseN3Id = planilhaN3.Id;
                    PlanilhaRepasseMktplaceDAO.insereN4(planilhaN4, ref strMsgErro);
                    if (!blnAux)
                    {
                        BD.rollbackTransacao();
                        avisoErro("Erro ao gravar linha específica da planilha no banco de dados!\n\nA operação será interrompida!\n\n" + strMsgErro);
                        return;
                    }
                    listPlanilhaN4.Add(planilhaN4);
                }
                #endregion

            }
            catch (Exception ex)
            {
                BD.rollbackTransacao();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
                avisoErro("Erro ao gravar dados da planilha no banco de dados!");
                return;
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);
            }


            #region [ Verifica se há divergências nos dados da planilha ]
            try
            {
                info(ModoExibicaoMensagemRodape.EmExecucao, "verificando divergências nos dados da planilha");

                #region [ checa consistência nas informações do pedido ]
                foreach (PlanilhaRepasseMktplaceN2 item in listPlanilhaN2)
                {
                    DataRow row;
                    PlanilhaRepasseMktplaceDAO.checkPedidoDados(item, out strMsgErro);
                    if (strMsgErro.Length > 0)
                    {
                        avisoErro("Falha ao fazer a checagem de um pedido específico.\nA operação será interrompida!!\n\n" + strMsgErro);
                        return;
                    }
                    // Checa se o pedido existe
                    if (!item.PedidoExiste)
                    {
                        if (item.Pedido.Length > 0 && (item.DataPedido != DateTime.MinValue))
                        {
                            row = dtResultado.NewRow();
                            row["Linha"] = item.Linha;
                            row["Pedido"] = item.Pedido;
                            row["DataPedido"] = item.DataPedido == DateTime.MinValue ? "" : item.DataPedido.ToString("dd/MM/yyyy");
                            row["ValorTotalPedido"] = Global.formataMoeda(item.ValorTotalPedido);
                            row["Observacao"] = "Pedido não existe.";
                            dtResultado.Rows.Add(row);
                            continue; 
                        }
                    }
                    // Checa se o valor total do pedido confere
                    if (!item.ValorTotalConfere && item.ValorTotalPedidoCorreto > 0)
                    {
                        row = dtResultado.NewRow();
                        row["Linha"] = item.Linha;
                        row["Pedido"] = item.Pedido;
                        row["DataPedido"] = item.DataPedido == DateTime.MinValue ? "" : item.DataPedido.ToString("dd/MM/yyyy");
                        row["ValorTotalPedido"] = Global.formataMoeda(item.ValorTotalPedido);
                        row["Observacao"] = "Valor total do pedido não confere. Valor correto = " + Global.formataMoeda(item.ValorTotalPedidoCorreto) + ".";
                        dtResultado.Rows.Add(row);
                        continue;
                    }
                }
                #endregion

                #region [ checa consistência nas operações de repasse ]
                var query = from n4 in listPlanilhaN4
                            join n3 in listPlanilhaN3 on n4.MktplceRepasseN3Id equals n3.Id
                            join n2 in listPlanilhaN2 on n3.MktplceRepasseN2Id equals n2.Id

                            select new
                            {
                                n2.Pedido,
                                n2.DataPedido,
                                n2.PedidoExiste,
                                n2.ValorTotalConfere,
                                n2.ValorTotalPedido,
                                LinhaPlanilhaN2 = n2.Linha,
                                n4.TipoTransacao,
                                n4.Valor,
                                LinhaPlanilhaN4 = n4.Linha
                            };

                // checa se pedido com estorno está realmente cancelado ou devolvido
                foreach (var item in query.Where(p => p.TipoTransacao == MktplaceRepasseTipoTransacao.EstornoVenda.GetId() && p.PedidoExiste == true))
                {
                    DataRow row;
                    blnAux = PlanilhaRepasseMktplaceDAO.isPedidoCancelado(item.Pedido, out strMsgErro);
                    if (strMsgErro.Length > 0)
                    {
                        avisoErro("Falha ao fazer a checagem de um pedido específico.\nA operação será interrompida!!\n\n" + strMsgErro);
                        return;
                    }
                    if (!blnAux)
                    {
                        row = dtResultado.NewRow();
                        row["Linha"] = item.LinhaPlanilhaN4;
                        row["Pedido"] = item.Pedido;
                        row["DataPedido"] = item.DataPedido == DateTime.MinValue ? "" : item.DataPedido.ToString("dd/MM/yyyy");
                        row["ValorTotalPedido"] = Global.formataMoeda(item.ValorTotalPedido);
                        row["Tipo"] = MktplaceRepasseTipoTransacao.EstornoVenda.GetDescricao();
                        row["Valor"] = Global.formataMoeda(item.Valor);
                        row["Observacao"] = "Pedido com estorno de venda não consta como cancelado ou devolvido.";
                        dtResultado.Rows.Add(row);
                    }
                }

                // checa se pedido com estorno comissão sem desbloqueio consta no sistema e não está cancelado ou devolvido
                foreach (var item in query.Where(p => p.TipoTransacao == MktplaceRepasseTipoTransacao.ComissaoSemDesbloqueio.GetId() && p.PedidoExiste))
                {
                    DataRow row;
                    blnAux = PlanilhaRepasseMktplaceDAO.isPedidoCancelado(item.Pedido, out strMsgErro);
                    if (strMsgErro.Length > 0)
                    {
                        avisoErro("Falha ao fazer a checagem de um pedido específico.\nA operação será interrompida!!\n\n" + strMsgErro);
                        return;
                    }
                    if (!blnAux)
                    {
                        row = dtResultado.NewRow();
                        row["Linha"] = item.LinhaPlanilhaN4;
                        row["Pedido"] = item.Pedido;
                        row["DataPedido"] = item.DataPedido == DateTime.MinValue ? "" : item.DataPedido.ToString("dd/MM/yyyy");
                        row["ValorTotalPedido"] = Global.formataMoeda(item.ValorTotalPedido);
                        row["Tipo"] = MktplaceRepasseTipoTransacao.ComissaoSemDesbloqueio.GetDescricao();
                        row["Valor"] = Global.formataMoeda(item.Valor);
                        row["Observacao"] = "Estorno de comissão sem desbloqueio de pedido NÃO cancelado.";
                        dtResultado.Rows.Add(row);
                    }
                }
                #endregion

                info(ModoExibicaoMensagemRodape.EmExecucao, "preenchendo o grid de dados");

                #region [ Preenche grid com o resultado ]
                dtResultado.DefaultView.Sort = "Linha";
                dtResultado = dtResultado.DefaultView.ToTable();
                gridDados.DataSource = dtResultado;
                #endregion
            }
            catch (Exception ex)
            {
                BD.rollbackTransacao();
                avisoErro("Erro ao verificar divergências nos dados da planilha!\n\nToda a operação será cancelada!" + ex.Message.ToString());
                return;
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);
            } 
            #endregion

            BD.commitTransacao();
        }
        #endregion

        #endregion

        #region [ Eventos ]

        #region [ FPlanilhaPagamentoMarketplaceB2W ]

        #region [ FPlanilhaPagamentoMarketplaceB2W_Load ]
        private void FPlanilhaPagamentoMarketplaceB2W_Load(object sender, EventArgs e)
        {
            bool blnSucesso = false;

            try
            {
                limpaCampos();

                blnSucesso = true;
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
                if (!blnSucesso) Close();
            }
        }
        #endregion

        #region [ FPlanilhaPagamentoMarketplaceB2W_Shown ]
        private void FPlanilhaPagamentoMarketplaceB2W_Shown(object sender, EventArgs e)
        {
            try
            {
                #region[ Executa rotinas de inicialização ]
                if (!_InicializacaoOk)
                {
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
                // Se não inicializou corretamente, assegura-se de que o painel será fechado
                if (!_InicializacaoOk) Close();
            }
        }
        #endregion

        #region [ FPlanilhaPagamentoMarketplaceB2W_FormClosing ]
        private void FPlanilhaPagamentoMarketplaceB2W_FormClosing(object sender, FormClosingEventArgs e)
        {
            FMain.fMain.Location = this.Location;
            FMain.fMain.Visible = true;
            this.Visible = false;
        }

        #endregion

        #endregion

        #region [ btnSelecionaArqPlanilha_Click ]
        private void btnSelecionaArqPlanilha_Click(object sender, EventArgs e)
        {
            trataBotaoSelecionaArqPlanilha();
        }
        #endregion

        #region [ btnConfirmar_Click ]
        private void btnConfirmar_Click(object sender, EventArgs e)
        {
            trataBotaoConfirmar();
        } 
        #endregion

        #endregion

    }
}
