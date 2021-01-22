using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Media;
using System.Text;
using System.Windows.Forms;
using System.Xml;


namespace ADM2
{
    public partial class FImportXML : ADM2.FModelo
    {

        #region [ Constantes ]
        const String GRID_COL_ID_ESTOQUE = "colIdEstoque";
        const String GRID_COL_DATA_ENTRADA = "colDataEntrada";
        const String GRID_COL_CD = "colCD";
        const String GRID_COL_DOCUMENTO = "colDocumento";
        const String GRID_COL_FABRICANTE = "colFabricante";
        const String GRID_COL_DESCRICAO = "colDescricao";
        #endregion


        #region [ Atributos ]
        private bool _emProcessamento = false;
        private bool _InicializacaoOk;
        //private BancoDados _bd;
        public bool inicializacaoOk
        {
            get { return _InicializacaoOk; }
        }

        private bool _OcorreuExceptionNaInicializacao;
        public bool ocorreuExceptionNaInicializacao
        {
            get { return _OcorreuExceptionNaInicializacao; }
        }

        public DataTable dtbConsulta = new DataTable();

        #endregion

        public FImportXML()
        {
            InitializeComponent();
        }

        #region [ Métodos Privados ]

        #region [ executaPesquisa ]
        private bool executaPesquisa()
        {
            #region [ Declarações ]
            const string NOME_DESTA_ROTINA = "FImportXML.executaPesquisa()";
            String strSql;
            String strFrom;
            String strWhere = "";
            //BancoDados bd;
            //SqlConnection cnConexao;
            SqlCommand cmCommand;
            SqlDataAdapter daAdapter;
            #endregion

            try
            {
                #region [ Limpa campos dos dados de resposta ]
                //limpaCamposDados();
                #endregion

                #region [ Monta restrições da cláusula 'Where' ]
                strWhere = " WHERE (" +
                    "(t_ESTOQUE.data_entrada >= '2020-01-01')" +
                    " AND " +
                    "(t_ESTOQUE.data_entrada < '2021-02-01')" +
                ")" +
              " AND " +
                    "(t_ESTOQUE_XML.xml_prioridade = 1)";


                strFrom = " FROM t_ESTOQUE" +
                " INNER JOIN t_ESTOQUE_XML ON (t_ESTOQUE.id_estoque=t_ESTOQUE_XML.id_estoque)";
                #endregion

                this.Cursor = Cursors.WaitCursor;
                info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

                #region [ Cria objetos de BD ]
                //cnConexao = _bd.getNovaConexao();
                cmCommand = FMain.contextoBD.AmbienteBase.BD.criaSqlCommand();
                daAdapter = FMain.contextoBD.AmbienteBase.BD.criaSqlDataAdapter();
                #endregion

                #region [ Inicialização ]
                //nop
                #endregion

                #region [ Monta o SQL ]
                strSql = "SELECT" +
                " t_ESTOQUE.id_estoque, t_ESTOQUE.data_entrada, t_ESTOQUE.id_nfe_emitente, t_ESTOQUE.documento, t_ESTOQUE.fabricante, t_ESTOQUE.obs," +
                " t_ESTOQUE_XML.xml_conteudo, t_ESTOQUE_XML.xml_prioridade ";

                strSql +=
                strFrom +
                strWhere +
                " ORDER BY t_ESTOQUE.id_estoque, t_ESTOQUE.data_entrada";
                #endregion

                #region [ Executa a consulta no BD ]
                cmCommand.CommandText = strSql;
                daAdapter.SelectCommand = cmCommand;
                daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
                daAdapter.Fill(dtbConsulta);
                aviso("dtbConsulta.Rows.Count = " + dtbConsulta.Rows.Count.ToString());
                #endregion

                #region [ Carrega dados no grid ]
                //try
                //{
                //    info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
                //    grdDados.SuspendLayout();

                //    grdDados.Rows.Clear();
                //    if (dtbConsulta.Rows.Count > 0) grdDados.Rows.Add(dtbConsulta.Rows.Count);

                //    for (int i = 0; i < dtbConsulta.Rows.Count; i++)
                //    {
                //        rowConsulta = dtbConsulta.Rows[i];
                //        grdDados.Rows[i].Cells[GRID_COL_ID_ESTOQUE].Value = BD.readToString(rowConsulta["id_estoque"]);
                //        grdDados.Rows[i].Cells[GRID_COL_ID_ESTOQUE].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //        grdDados.Rows[i].Cells[GRID_COL_DATA_ENTRADA].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["data_entrada"]));
                //        grdDados.Rows[i].Cells[GRID_COL_DATA_ENTRADA].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //        grdDados.Rows[i].Cells[GRID_COL_CD].Value = BD.readToString(rowConsulta["id_nfe_emitente"]);
                //        grdDados.Rows[i].Cells[GRID_COL_CD].Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                //        grdDados.Rows[i].Cells[GRID_COL_DOCUMENTO].Value = BD.readToString(rowConsulta["documento"]);
                //        grdDados.Rows[i].Cells[GRID_COL_DOCUMENTO].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //        grdDados.Rows[i].Cells[GRID_COL_FABRICANTE].Value = BD.readToString(rowConsulta["fabricante"]);
                //        grdDados.Rows[i].Cells[GRID_COL_FABRICANTE].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //        grdDados.Rows[i].Cells[GRID_COL_DESCRICAO].Value = BD.readToString(rowConsulta["obs"]);
                //        grdDados.Rows[i].Cells[GRID_COL_DESCRICAO].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //    }

                //    //#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
                //    //for (int i = 0; i < grdDados.Rows.Count; i++)
                //    //{
                //    //	if (grdDados.Rows[i].Selected) grdDados.Rows[i].Selected = false;
                //    //}
                //    //#endregion
                //}
                //finally
                //{
                //    grdDados.ResumeLayout();
                //}
                #endregion

                #region [Totais]
                lblTotalRegistros.Text = Global.formataInteiro(dtbConsulta.Rows.Count);
                #endregion

                this.Cursor = Cursors.Default;

                grdDados.Focus();

                // Feedback da conclusão da pesquisa
                SystemSounds.Exclamation.Play();

                return true;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
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

        #region [ executaAtualizacaoDataEmissao ]
        private void executaAtualizacaoDataEmissao()
        {
            #region [ Declarações ]
            const string NOME_DESTA_ROTINA = "FImportXML.executaAtualizacaoDataEmissao";
            int qtdeRegistrosParaAtualizar = 0;
            int qtdeRegistrosAtualizados = 0;
            int qtdeRegistrosAtualizadosSucesso = 0;
            int qtdeRegistrosAtualizadosFalha = 0;
            int qtdeRegistrosAtualizadosSucessoUpdatePedidoRecebidoData = 0;
            int qtdeRegistrosAtualizadosFalhaUpdatePedidoRecebidoData = 0;
            int qtdeRegistrosAtualizadosSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = 0;
            int qtdeRegistrosAtualizadosFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = 0;
            int percProgresso;
            int percProgressoAnterior;
            bool blnFalhaAtualizacao;
            bool blnUpdatePedidoRecebidoData;
            bool blnUpdateMarketplacePedidoRecebidoRegistrarDataRecebido;
            string strMsg;
            string strMsgProgresso;
            string strMsgErroLog = "";
            string msg_erro = "";
            DateTime dtInicioProcessamento;
            TimeSpan tsDuracaoProcessamento;
            StringBuilder sbLogSucessoUpdatePedidoRecebidoData = new StringBuilder("");
            StringBuilder sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = new StringBuilder("");
            StringBuilder sbLogFalhaUpdatePedidoRecebidoData = new StringBuilder("");
            StringBuilder sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = new StringBuilder("");
            Log log = new Log();
            XmlDocument xmlDoc;
            XmlNodeList elemList;
            DataRow rowConsulta;
            #endregion

            try
            {
                //for (int iv = 0; iv < dtbConsulta.Rows.Count; iv++)
                //{
                //    if (dtbConsulta[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LIBERADO_PARA_REGISTRAR_PEDIDO_RECEBIDO_CLIENTE) qtdeRegistrosParaAtualizar++;
                //}

                //if (qtdeRegistrosParaAtualizar == 0)
                //{
                //    avisoErro("Não há nenhum registro para ser atualizado no banco de dados!");
                //    return;
                //}

                #region [ Solicita confirmação antes de executar a operação ]
                strMsg = "Confirma a atualização no banco de dados?";
                if (!confirma(strMsg)) return;
                #endregion

                #region [ Inicialização do processamento ]
                dtInicioProcessamento = DateTime.Now;
                strMsg = "Início da atualização no banco de dados";
                #endregion
                percProgressoAnterior = -1;
                xmlDoc = new XmlDocument();

                for (int iv = 0; iv < dtbConsulta.Rows.Count; iv++)
                {
                    rowConsulta = dtbConsulta.Rows[iv];
                    xmlDoc.LoadXml(BD.readToString(rowConsulta["xml_conteudo"]));
                    elemList = xmlDoc.GetElementsByTagName("dhEmi");
                    lbMensagem.Items.Add("id_estoque = " + BD.readToString(rowConsulta["id_estoque"]) + " - dthemi = " + elemList[0].InnerText);

                //if (dtbConsulta[iv].processo.Status == eRastreioPedidoRecebidoClienteProcessoStatus.LIBERADO_PARA_REGISTRAR_PEDIDO_RECEBIDO_CLIENTE)
                //{
                //    qtdeRegistrosAtualizados++;

                    //    #region [ Progresso ]
                    //    percProgresso = 100 * qtdeRegistrosAtualizados / qtdeRegistrosParaAtualizar;
                    //    if (percProgressoAnterior != percProgresso)
                    //    {
                    //        strMsgProgresso = "Atualizando pedidos no banco de dados: " + qtdeRegistrosAtualizados.ToString() + " / " + qtdeRegistrosParaAtualizar.ToString() + "   (" + percProgresso.ToString() + "%)";
                    //        info(ModoExibicaoMensagemRodape.EmExecucao, strMsgProgresso);
                    //        percProgressoAnterior = percProgresso;
                    //        Application.DoEvents();
                    //    }
                    //    #endregion

                    //    #region [ Executa atualização no banco de dados ]
                    //    blnFalhaAtualizacao = false;
                    //    blnUpdatePedidoRecebidoData = FMain.contextoBD.AmbienteBase.anotarPedidoRecebidoClienteDAO.UpdatePedidoRecebidoData(dtbConsulta[iv].processo.Pedido, dtbConsulta[iv].dadosNormalizado.dtDataEntrega, Global.Usuario.usuario, out msg_erro);
                    //    if (blnUpdatePedidoRecebidoData)
                    //    {
                    //        qtdeRegistrosAtualizadosSucessoUpdatePedidoRecebidoData++;
                    //        if (sbLogSucessoUpdatePedidoRecebidoData.Length > 0) sbLogSucessoUpdatePedidoRecebidoData.Append(", ");
                    //        sbLogSucessoUpdatePedidoRecebidoData.Append(dtbConsulta[iv].processo.Pedido);
                    //    }
                    //    else
                    //    {
                    //        if (sbLogFalhaUpdatePedidoRecebidoData.Length > 0) sbLogFalhaUpdatePedidoRecebidoData.Append(", ");
                    //        sbLogFalhaUpdatePedidoRecebidoData.Append(dtbConsulta[iv].processo.Pedido);

                    //        blnFalhaAtualizacao = true;
                    //        qtdeRegistrosAtualizadosFalhaUpdatePedidoRecebidoData++;
                    //        dtbConsulta[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.FALHA_REGISTRO_PEDIDO_RECEBIDO_CLIENTE;
                    //        strMsg = "Falha ao tentar atualizar o pedido " + dtbConsulta[iv].processo.Pedido + " (NF: " + dtbConsulta[iv].dadosNormalizado.NF + "): " + msg_erro;
                    //        dtbConsulta[iv].processo.MensagemErro = strMsg;
                    //        adicionaErro(strMsg);
                    //    }

                    //    if (blnUpdatePedidoRecebidoData)
                    //    {
                    //        if ((dtbConsulta[iv].processo.marketplace_codigo_origem.Trim().Length > 0) && (dtbConsulta[iv].processo.MarketplacePedidoRecebidoRegistrarStatus == 0))
                    //        {
                    //            blnUpdateMarketplacePedidoRecebidoRegistrarDataRecebido = FMain.contextoBD.AmbienteBase.anotarPedidoRecebidoClienteDAO.UpdateMarketplacePedidoRecebidoRegistrarDataRecebido(dtbConsulta[iv].processo.Pedido, dtbConsulta[iv].dadosNormalizado.dtDataEntrega, Global.Usuario.usuario, out msg_erro);
                    //            if (blnUpdateMarketplacePedidoRecebidoRegistrarDataRecebido)
                    //            {
                    //                qtdeRegistrosAtualizadosSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido++;
                    //                if (sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Length > 0) sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Append(", ");
                    //                sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Append(dtbConsulta[iv].processo.Pedido);
                    //            }
                    //            else
                    //            {
                    //                if (sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Length > 0) sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Append(", ");
                    //                sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Append(dtbConsulta[iv].processo.Pedido);

                    //                blnFalhaAtualizacao = true;
                    //                qtdeRegistrosAtualizadosFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido++;
                    //                dtbConsulta[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.FALHA_REGISTRO_PEDIDO_RECEBIDO_CLIENTE;
                    //                strMsg = "Falha ao tentar atualizar o pedido " + dtbConsulta[iv].processo.Pedido + " (NF: " + dtbConsulta[iv].dadosNormalizado.NF + "): " + msg_erro;
                    //                dtbConsulta[iv].processo.MensagemErro = strMsg;
                    //                adicionaErro(strMsg);
                    //            }
                    //        }
                    //    }
                    //#endregion

                    //    #region [ Atualiza status no grid ]
                    //    for (int jv = 0; jv < grid.Rows.Count; jv++)
                    //    {
                    //        if (grid.Rows[jv].Cells[GRID_COL_HIDDEN_GUID].Value.ToString().Equals(dtbConsulta[iv].processo.Guid))
                    //        {
                    //            if (blnFalhaAtualizacao)
                    //            {
                    //                grid.Rows[jv].Cells[GRID_COL_STATUS].Value = "FALHA";
                    //                grid.Rows[jv].Cells[GRID_COL_STATUS].Style.ForeColor = Color.Red;
                    //                grid.Rows[jv].Cells[GRID_COL_MENSAGEM].Value = dtbConsulta[iv].processo.MensagemErro;
                    //                grid.Rows[jv].Cells[GRID_COL_MENSAGEM].Style.ForeColor = Color.Red;
                    //            }
                    //            else
                    //            {
                    //                grid.Rows[jv].Cells[GRID_COL_STATUS].Value = "OK";
                    //                grid.Rows[jv].Cells[GRID_COL_STATUS].Style.ForeColor = Color.Green;
                    //            }
                    //            break;
                    //        }
                    //    }
                    //    #endregion

                    //    if (blnFalhaAtualizacao)
                    //    {
                    //        qtdeRegistrosAtualizadosFalha++;
                    //        // Prossegue para o próximo registro
                    //        continue;
                    //    }

                    //    qtdeRegistrosAtualizadosSucesso++;
                    //    dtbConsulta[iv].processo.Status = eRastreioPedidoRecebidoClienteProcessoStatus.SUCESSO_REGISTRO_PEDIDO_RECEBIDO_CLIENTE;
                    //    strMsg = "Sucesso na atualização do pedido " + dtbConsulta[iv].processo.Pedido + " (NF: " + dtbConsulta[iv].dadosNormalizado.NF + ")";
                    //    adicionaDisplay(strMsg);
                    //}
                }

                //    lblQtdeAtualizSucesso.Text = Global.formataInteiro(qtdeRegistrosAtualizadosSucesso);
                //lblQtdeAtualizFalha.Text = Global.formataInteiro(qtdeRegistrosAtualizadosFalha);

                //#region [ Grava o log ]
                //strMsg = "[Módulo ADM2] Operação 'Anotar Pedidos Recebidos pelo Cliente':";
                //if (sbLogSucessoUpdatePedidoRecebidoData.Length > 0) strMsg += "\nSucesso (campo 'PedidoRecebidoData') [" + Global.formataInteiro(qtdeRegistrosAtualizadosSucessoUpdatePedidoRecebidoData) + " pedidos]: " + sbLogSucessoUpdatePedidoRecebidoData.ToString();
                //if (sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Length > 0) strMsg += "\nSucesso (campo 'MarketplacePedidoRecebidoRegistrarDataRecebido') [" + Global.formataInteiro(qtdeRegistrosAtualizadosSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido) + " pedidos]: " + sbLogSucessoUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.ToString();
                //if (sbLogFalhaUpdatePedidoRecebidoData.Length > 0) strMsg += "\nFalha (campo 'PedidoRecebidoData') [" + Global.formataInteiro(qtdeRegistrosAtualizadosFalhaUpdatePedidoRecebidoData) + " pedidos]: " + sbLogFalhaUpdatePedidoRecebidoData.ToString();
                //if (sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.Length > 0) strMsg += "\nFalha (campo 'MarketplacePedidoRecebidoRegistrarDataRecebido') [" + Global.formataInteiro(qtdeRegistrosAtualizadosFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido) + " pedidos]: " + sbLogFalhaUpdateMarketplacePedidoRecebidoRegistrarDataRecebido.ToString();
                //strMsg += "\nArquivo processado: " + txtArquivoRastreio.Text.Trim() + " (contendo " + Global.formataInteiro(dtbConsulta.Count) + " registros)";
                //log.operacao = Global.Cte.ADM2.LogOperacao.OP_LOG_PEDIDO_RECEBIDO_VIA_ADM2;
                //log.usuario = Global.Usuario.usuario;
                //log.complemento = strMsg;
                //FMain.contextoBD.AmbienteBase.logDAO.insere(Global.Usuario.usuario, log, ref strMsgErroLog);
                //#endregion

                //tsDuracaoProcessamento = DateTime.Now - dtInicioProcessamento;

                #region [ Mensagem de sucesso ]
                info(ModoExibicaoMensagemRodape.Normal);
                //strMsg = "Atualização no banco de dados concluído com sucesso (duração: " + Global.formataDuracaoHMS(tsDuracaoProcessamento) + ")!";
                //aviso(strMsg);
            #endregion
        }
            catch (Exception ex)
            {
                //Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + ex.ToString());
                //adicionaErro(ex.Message);
                //avisoErro(ex.ToString());
                avisoErro("deu ruim");
                return;
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);
            }
        }
        #endregion

        #endregion


        #region [ Eventos ]

        private void FImportXML_Load(object sender, EventArgs e)
        {
            if (true) { };
        }

        private void FImportXML_Shown(object sender, EventArgs e)
        {
            executaPesquisa(); ;
        }

        private void FImportXML_FormClosing(object sender, FormClosingEventArgs e)
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

        private void BtnAtualizaDatas_Click(object sender, EventArgs e)
        {
            executaAtualizacaoDataEmissao();
        }
        #endregion

    }
}
