#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
#endregion

namespace ConsolidadorXlsEC 
{
    public partial class FIntegracaoMarketplace : FModelo
    {
        #region [ Constantes ]
        public const string PEDIDO_MAGENTO_STATUS_VALIDOS = "|despachado|rastreio_ic|";
        public const string ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_SKYHUB = "|B2W|Zoom|Magazine Luiza|Carrefour|CNOVA|";
        public const string ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_INTEGRACOMMERCE = "";
        public const string ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET = "|Leroy Merlin|";
        public const string COD_ST_PEDIDO_RECEBIDO_NAO = "0";
        public const string COD_ST_PEDIDO_RECEBIDO_SIM = "1";
        public const string COD_ST_PEDIDO_RECEBIDO_NAO_DEFINIDO = "10";
        #endregion

        #region [ Atributos ]
        private bool _InicializacaoOk;
        public bool inicializaoOk
        {
            get { return _InicializacaoOk; }
        }

        private bool _OcorreuExceptionNaInicializacao;
        public bool ocorreuExceptionNaInicializacao
        {
            get { return _OcorreuExceptionNaInicializacao; }
        }

        private SqlCommand _cmCommandPedidoRecebidoParaSim;
        #endregion

        #region [ Construtor ]
        public FIntegracaoMarketplace()
        {
            InitializeComponent();
        }
        #endregion

        #region [ Métodos ]

        #region [ limpaCampos ]
        private void limpaCampos()
        {
            cbTransportadora.SelectedIndex = -1;
            cbOrigemPedidoGrupo.SelectedIndex = -1;
            cbOrigemPedido.SelectedIndex = -1;
            txtLoja.Clear();
            grdDados.DataSource = null;
            lblTotalRegistros.Text = "";
        }
        #endregion

        #region [ montaClausulaWhere ]
        private string montaClausulaWhere()
        {
            #region [ Declarações ]
            StringBuilder sbWhere = new StringBuilder();
            string strAux;
            #endregion

            #region [ Critérios de Restrição ]
            strAux = "(p.st_entrega = '" + Global.Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE + "')" +
              " AND (p.MarketplacePedidoRecebidoRegistradoStatus = " + Global.Cte.StPedidoRecebido.COD_ST_PEDIDO_RECEBIDO_NAO + ")" +
              " AND (p.MarketplacePedidoRecebidoRegistrarStatus = " + Global.Cte.StPedidoRecebido.COD_ST_PEDIDO_RECEBIDO_SIM + ")" +
              " AND (p.marketplace_codigo_origem IS NOT NULL) AND (LEN(Coalesce(p.marketplace_codigo_origem,'')) > 0)";
            if (sbWhere.Length > 0) sbWhere.Append(" AND");
            sbWhere.Append(strAux);
            #endregion

            #region [ Transportadora ]
            if ((cbTransportadora.SelectedIndex > -1) && (cbTransportadora.SelectedValue.ToString().Trim().Length > 0))
            {
                strAux = " (p.transportadora_id = '" + cbTransportadora.SelectedValue.ToString() + "')";
                if (sbWhere.Length > 0) sbWhere.Append(" AND");
                sbWhere.Append(strAux);
            }
            #endregion

            #region [ Origem do Pedido (Grupo) ]
            if ((cbOrigemPedidoGrupo.SelectedIndex > -1) && (cbOrigemPedidoGrupo.SelectedValue.ToString().Length > 0))
            {
                strAux = getPedidoECommerceOrigemFmtSql(cbOrigemPedidoGrupo.SelectedValue.ToString());
                strAux = " p.marketplace_codigo_origem IN(" + strAux + ")";
                if (sbWhere.Length > 0) sbWhere.Append(" AND");
                sbWhere.Append(strAux);
            }
            #endregion

            #region [ Origem do Pedido ]
            if ((cbOrigemPedido.SelectedIndex > -1) && (cbOrigemPedido.SelectedValue.ToString().Length > 0))
            {
                strAux = " (p.marketplace_codigo_origem = '" + cbOrigemPedido.SelectedValue.ToString() + "')";
                if (sbWhere.Length > 0) sbWhere.Append(" AND");
                sbWhere.Append(strAux);
            }
            #endregion

            #region [ Loja ]
            if (txtLoja.Text.Trim().Length > 0)
            {
                strAux = " (p.numero_loja = " + txtLoja.Text + ")";
                if (sbWhere.Length > 0) sbWhere.Append(" AND");
                sbWhere.Append(strAux);
            }
            #endregion

            return sbWhere.ToString();
        }
        #endregion

        #region [ montaSqlConsulta ]
        private string montaSqlConsulta()
        {
            #region [ Declarações ]
            string strWhere;
            string strSql;
            #endregion

            #region [ Monta cláusula Where ]
            strWhere = montaClausulaWhere();
            if (strWhere.Length > 0) strWhere = " WHERE " + strWhere;
            #endregion

            #region [ Monta Select ]
            strSql = "SELECT" +
                        " p.transportadora_id," +
                        " p.pedido," +
                        " p.pedido_bs_x_ac," +
                        " p.pedido_bs_x_marketplace," +
                        " p.marketplace_codigo_origem," +
                        " p.loja," +
                        " p.MarketplacePedidoRecebidoRegistrarDataRecebido," +
                        " c.cidade," +
                        " c.uf," +
                        " c.nome_iniciais_em_maiusculas," +
                        " Sum(tPI.qtde*tPI.preco_venda) AS vl_pedido," +
                        " (SELECT descricao FROM t_CODIGO_DESCRICAO WHERE grupo = 'PedidoECommerce_Origem' AND codigo = p.marketplace_codigo_origem) AS marketplace_codigo_origem_descricao," +
                        " (SELECT codigo_pai FROM t_CODIGO_DESCRICAO WHERE grupo = 'PedidoECommerce_Origem' AND codigo = p.marketplace_codigo_origem) AS marketplace_codigo_origem_pai" +
                    " FROM t_PEDIDO p" +
                    " INNER JOIN t_PEDIDO_ITEM tPI ON (p.pedido=tPI.pedido)" +
                    " INNER JOIN t_CLIENTE c ON (p.id_cliente=c.id)" +
                        strWhere +
                    " GROUP BY p.transportadora_id" +
                       " ,p.pedido" +
                       " ,p.pedido_bs_x_ac" +
                       " ,p.pedido_bs_x_marketplace" +
                       " ,p.marketplace_codigo_origem" +
                       " ,p.loja" +
                       " ,p.MarketplacePedidoRecebidoRegistrarDataRecebido" +
                       " ,c.cidade" +
                       " ,c.uf" +
                       " ,c.nome_iniciais_em_maiusculas" +
                    " ORDER BY" +
                        " p.transportadora_id," +
                        " p.pedido";
            #endregion

            return strSql;
        }
        #endregion

        #region [ executaPesquisa ]
        private void executaPesquisa()
        {
            #region [ Declarações ]
            string strSql;
            string strMsgErroBDCompleto = "";
            string strMsgErroBDResumido = "";
            SqlCommand cmCommand;
            SqlDataAdapter daAdapter;
            DataTable dtbConsulta = new DataTable();
            DataRow rowConsulta;
            #endregion

            try
            {
                #region [ Verifica se a conexão com o BD está ok ]
                if (FMain.contextoBD.AmbienteBase.BD.isConexaoOk())
                {
                    if (!FMain.contextoBD.AmbienteBase.reiniciaBancoDados(out strMsgErroBDCompleto, out strMsgErroBDResumido))
                    {
                        avisoErro("Ocorreu uma falha na conexão com o Banco de Dados!!\nA tentativa de reconectar automaticamente falhou!!\nPor favor, aguarde alguns instantes e tente outra vez!!\n\n" + strMsgErroBDResumido);
                        return;
                    }
                }
                #endregion

                info(ModoExibicaoMensagemRodape.EmExecucao, "consultando banco de dados");

                #region [ Cria objetos de BD ]
                cmCommand = FMain.contextoBD.AmbienteBase.BD.criaSqlCommand();
                daAdapter = FMain.contextoBD.AmbienteBase.BD.criaSqlDataAdapter();
                #endregion

                #region [ Monta o SQL da consulta ]
                strSql = montaSqlConsulta();
                #endregion

                #region [ Executa a consulta no BD ]
                cmCommand.CommandText = strSql;
                daAdapter.SelectCommand = cmCommand;
                daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
                daAdapter.Fill(dtbConsulta);
                #endregion

                #region [ Prepara alguns campos que necessitam de formatação ]
                for (int i = 0; i < dtbConsulta.Rows.Count; i++)
                {
                    rowConsulta = dtbConsulta.Rows[i];
                    rowConsulta["vl_pedido"] = Global.formataMoeda((decimal)rowConsulta["vl_pedido"]);
                    rowConsulta["marketplace_codigo_origem_pai"] = this.getDescricaoTCodigoDescricao("PedidoECommerce_Origem_Grupo", (string)rowConsulta["marketplace_codigo_origem_pai"]);
                }
                #endregion

                #region [ Exibição dos dados no grid ]
                try
                {
                    info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");

                    grdDados.SuspendLayout();

                    #region [ Carrega os dados no grid ]
                    grdDados.DataSource = dtbConsulta;
                    #endregion

                    #region [ Exibe o grid sem nenhuma linha pré-selecionada ]
                    grdDados.ClearSelection();
                    #endregion                    
                }
                finally
                {
                    grdDados.ResumeLayout();
                }
                #endregion

                #region [ Exibe totalização ]
                lblTotalRegistros.Text = dtbConsulta.Rows.Count.ToString();
                #endregion

                btnConfirma.Enabled = true;

                grdDados.Focus();
            }
            catch (Exception ex)
            {
                avisoErro(ex.ToString());
                Close();
                return;
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);
            }
        }
        #endregion

        #region [ trataBotaoMarcarTodos ]
        private void trataBotaoMarcarTodos()
        {
            if (grdDados.Rows.Count == 0) return;

            foreach (DataGridViewRow row in grdDados.Rows)
            {
                row.Cells[colGrdDadosCheckBox.Name].Value = true;
            }
        }
        #endregion

        #region [ trataBotaoDesmarcarTodos ]
        private void trataBotaoDesmarcarTodos()
        {
            if (grdDados.Rows.Count == 0) return;

            foreach (DataGridViewRow row in grdDados.Rows)
            {
                row.Cells[colGrdDadosCheckBox.Name].Value = false;
            }
        }
        #endregion

        #region [ trataBotaoConfirma ]
        private void trataBotaoConfirma()
        {
            #region [ Declarações ]
            bool blnEnviouOk;
            int intCounter = 0;
            string xmlReqSoap;
            string xmlRespSoap;
            string sessionId;
            string msg_erro_aux;
            string strOrigemPedidoAux;
            string strIncrementId = "";
            string strStatus = "";
            string strComment = "";
            List<DataGridViewRow> linhasSelecionadas = new List<DataGridViewRow>();
            List<DataGridViewRow> salesOrderInfoComStatusOk = new List<DataGridViewRow>();
            List<DataGridViewRow> salesOrderInfoComStatusInvalido = new List<DataGridViewRow>();
            SalesOrderInfo salesOrderInfoAux = new SalesOrderInfo();
            SalesOrderAddCommentRequest addCommentRequest;
            SalesOrderAddCommentResponse addCommentResponse;
            FConfirmaPedidoStatus fConfirmaPedidoStatus;
            DialogResult drConfirmaPedidoStatus;
            #endregion
            
            #region [ Captura os pedidos selecionados ]
            foreach (DataGridViewRow row in grdDados.Rows)
                row.DefaultCellStyle.BackColor = Color.White;
            foreach (DataGridViewRow row in grdDados.Rows)
            {
                if (Convert.ToBoolean(row.Cells[colGrdDadosCheckBox.Name].Value) == true)
                {
                    strOrigemPedidoAux = row.Cells[colGrdDadosCodigoOrigemPai.Name].Value.ToString();
                    strOrigemPedidoAux = "|" + strOrigemPedidoAux + "|";

                    if ((ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_SKYHUB.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) == -1) &&
                        (ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_INTEGRACOMMERCE.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) == -1) &&
                        (ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) == -1))
                    {
                        avisoErro("Não é possível selecionar pedidos que não sejam SkyHub (" + ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_SKYHUB.Trim('|').Replace("|", ", ") + ") ou IntegraCommerce " +
                            "(" + ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_INTEGRACOMMERCE.Trim('|').Replace(" | ", ", ") +
                            ") ou AnyMarket (" + ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET.Trim('|').Replace(" | ", ", ") + ")!");
                        row.DefaultCellStyle.BackColor = Color.LightSalmon;
                        return;
                    }

                    linhasSelecionadas.Add(row);
                }
            }
            #endregion

            #region [ Há pedidos selecionados? ]
            if (linhasSelecionadas.Count == 0)
            {
                avisoErro("Nenhum pedido foi selecionado!!");
                return;
            }
            #endregion

            #region [ Confirma execução ]
            if (!confirma("Tem certeza de que deseja baixar todos os pedidos selecionados (" + linhasSelecionadas.Count + ") no Magento?")) return;
            #endregion

            info(ModoExibicaoMensagemRodape.EmExecucao, "verificando status dos pedidos no magento");

            #region [ Requisição para efetuar login ]
            xmlReqSoap = Magento.montaRequisicaoLogin(Global.Cte.Magento.USER_NAME, Global.Cte.Magento.PASSWORD);
            blnEnviouOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.login, out xmlRespSoap, out msg_erro_aux);
            #endregion

            if (!blnEnviouOk)
            {
                info(ModoExibicaoMensagemRodape.Normal);
                avisoErro("Erro ao efetuar logon na API do Magento: \n\n" + msg_erro_aux);
                return;
            }

            #region [ Obtém a SessionId ]
            sessionId = Magento.obtemSessionIdFromLoginResponse(xmlRespSoap, out msg_erro_aux);
            #endregion

            if ((sessionId ?? "").Length == 0)
            {
                info(ModoExibicaoMensagemRodape.Normal);
                avisoErro("Falha ao tentar obter o SessionId!!");
                return;
            }

            try
            {
                foreach (DataGridViewRow item in linhasSelecionadas)
                {
                    intCounter++;
                    info(ModoExibicaoMensagemRodape.EmExecucao, "verificando status dos pedidos no magento: " + intCounter + " de " + linhasSelecionadas.Count);

                    #region [ Recupera o status dos pedidos ]
                    xmlReqSoap = Magento.montaRequisicaoCallSalesOrderInfo(sessionId, item.Cells[colGrdDadosNumMagento.Name].Value.ToString());
                    blnEnviouOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.call, out xmlRespSoap, out msg_erro_aux);

                    if (blnEnviouOk)
                    {
                        salesOrderInfoAux = Magento.decodificaXmlSalesOrderInfoResponse(xmlRespSoap, out msg_erro_aux);
                        if (!salesOrderInfoAux.faultResponse.isFaultResponse)
                        {
                            if (isSalesOrderStatusValido(salesOrderInfoAux.status))
                            {
                                item.Cells[colGrdDadosStatus.Name].Value = salesOrderInfoAux.status;
                                item.Cells[colGrdDadosStatusDescricao.Name].Value = getDescricaoStatusMagento(salesOrderInfoAux.status);
                                salesOrderInfoComStatusOk.Add(item);
                            }
                            else
                            {
                                item.Cells[colGrdDadosStatus.Name].Value = salesOrderInfoAux.status;
                                item.Cells[colGrdDadosStatusDescricao.Name].Value = getDescricaoStatusMagento(salesOrderInfoAux.status);
                                item.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
                                item.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: status inválido";
                                salesOrderInfoComStatusInvalido.Add(item);
                            } 
                        }
                        else
                        {
                            item.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
                            item.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: falha ao tentar consultar o status no magento \t" + salesOrderInfoAux.faultResponse.faultstring;
                        }
                    }
                    else
                    {
                        info(ModoExibicaoMensagemRodape.Normal);
                        avisoErro("Erro ao enviar requisição de um dos pedidos!!\n\nA operação será cancelada!!\n\n" + msg_erro_aux);
                        executaPesquisa();
                        return;
                    }
                    #endregion
                }

                #region [ Exibe Form com os pedidos contendo status inválido ]
                if (salesOrderInfoComStatusInvalido.Count > 0)
                {
                    info(ModoExibicaoMensagemRodape.Normal);
                    fConfirmaPedidoStatus = new FConfirmaPedidoStatus(salesOrderInfoComStatusInvalido);
                    drConfirmaPedidoStatus = fConfirmaPedidoStatus.ShowDialog();
                    if (drConfirmaPedidoStatus != DialogResult.OK)
                    {
                        info(ModoExibicaoMensagemRodape.Normal);
                        avisoErro("Operação não realizada!!");
                        executaPesquisa();
                        return;
                    }
                }
                else
                    fConfirmaPedidoStatus = new FConfirmaPedidoStatus();
                #endregion

                #region [ Altera status dos pedidos no magento ]
                intCounter = 0;
                preparaSqlCommandPedidoRecebidoParaSim();
                foreach (DataGridViewRow row in salesOrderInfoComStatusOk)
                {
                    intCounter++;
                    info(ModoExibicaoMensagemRodape.EmExecucao, "alterando status dos pedidos no magento: " + intCounter + " de " + salesOrderInfoComStatusOk.Count);
                    strOrigemPedidoAux = row.Cells[colGrdDadosCodigoOrigemPai.Name].Value.ToString();
                    strOrigemPedidoAux = "|" + strOrigemPedidoAux + "|";


                    if (ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_SKYHUB.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) != -1)
                    {
                        #region [ Tratamento dos pedidos skyhub ]
                        strIncrementId = row.Cells[colGrdDadosNumMagento.Name].Value.ToString();
                        strStatus = "complete";
                        strComment = "";
                        #endregion
                    }
                    else if (ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_INTEGRACOMMERCE.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) != -1)
                    {
                        #region [ Tratamento dos pedidos Integra Commerce ]
                        strIncrementId = row.Cells[colGrdDadosNumMagento.Name].Value.ToString();
                        strStatus = "delivered";
                        strComment = Global.formataDataDdMmYyyyComSeparador(Convert.ToDateTime(row.Cells[colGrdDadosRecebido.Name].Value));
                        #endregion
                    }
                    else if (ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) != -1)
                    {
                        #region [ Tratamento dos pedidos AnyMarket ]
                        strIncrementId = row.Cells[colGrdDadosNumMagento.Name].Value.ToString();
                        strStatus = "complete";
                        strComment = "";
                        #endregion
                    }
                    else
                        continue;

                    #region [ Enviar requisição AddComment ]
                    addCommentRequest = new SalesOrderAddCommentRequest();
                    addCommentRequest.orderIncrementId = strIncrementId;
                    addCommentRequest.status = strStatus;
                    addCommentRequest.comment = strComment;

                    xmlReqSoap = Magento.montaRequisicaoSalesOrderAddComment(sessionId, addCommentRequest);
                    blnEnviouOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.call, out xmlRespSoap, out msg_erro_aux);

                    if (blnEnviouOk)
                    {
                        msg_erro_aux = "";
                        addCommentResponse = Magento.decodificaXmlSalesOrderAddCommentResponse(xmlRespSoap, out msg_erro_aux);
                        if (!addCommentResponse.faultResponse.isFaultResponse)
                        {
                            if (msg_erro_aux.Length > 0)
                            {
                                row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
                                row.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: falha ao tentar atualizar status no magento \t" + msg_erro_aux;
                            }
                            else
                            {
                                row.Cells[colGrdDadosMensagemStatus.Name].Value += "status atualizado no magento";
                                if (!setPedidoRecebidoParaSim(row.Cells[colGrdDadosNumPedido.Name].Value.ToString(), out msg_erro_aux))
                                {
                                    row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Orange;
                                    row.Cells[colGrdDadosMensagemStatus.Name].Value = "Parcial: " + row.Cells[colGrdDadosMensagemStatus.Name].Value;
                                    row.Cells[colGrdDadosMensagemStatus.Name].Value += "; falha ao tentar baixar no sistema interno \t" + msg_erro_aux;
                                }
                                else
                                {
                                    row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Green;
                                    row.Cells[colGrdDadosMensagemStatus.Name].Value = "Sucesso: " + row.Cells[colGrdDadosMensagemStatus.Name].Value;
                                    row.Cells[colGrdDadosMensagemStatus.Name].Value += "; baixa no sistema interno";
                                }
                            }
                        }
                        else
                        {
                            row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
                            row.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: falha ao atualizar status no magento \t" + addCommentResponse.faultResponse.faultstring;
                        }

                    }
                    else
                    {
                        row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
                        row.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: falha ao atualizar status no magento \t" + msg_erro_aux;
                    }
                    #endregion
                }
                #endregion

                #region [ Baixa no BD pedidos com status inválido ]
                intCounter = 0;
                foreach (DataGridViewRow row in fConfirmaPedidoStatus.LinhasSelecionadas)
                {
                    intCounter++;
                    info(ModoExibicaoMensagemRodape.EmExecucao, "baixando no banco de dados pedidos com status inválidos " + intCounter + " de " + fConfirmaPedidoStatus.LinhasSelecionadas.Count);

                    if (!setPedidoRecebidoParaSim(row.Cells[colGrdDadosNumPedido.Name].Value.ToString(), out msg_erro_aux))
                    {
                        row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
                        row.Cells[colGrdDadosMensagemStatus.Name].Value += "; falha ao tentar baixar no sistema interno \t" + msg_erro_aux;
                    }
                    else
                    {
                        row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Orange;
                        row.Cells[colGrdDadosMensagemStatus.Name].Value = row.Cells[colGrdDadosMensagemStatus.Name].Value.ToString().Replace("Falha:", "Parcial:");
                        row.Cells[colGrdDadosMensagemStatus.Name].Value += "; baixa no sistema interno";
                    }
                }
                #endregion
            }
            finally
            {
                info(ModoExibicaoMensagemRodape.Normal);

                #region [ Encerra sessão ]
                xmlReqSoap = Magento.montaRequisicaoEndSession(sessionId);
                blnEnviouOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.endSession, out xmlRespSoap, out msg_erro_aux);
                #endregion

                if (!blnEnviouOk)
                {
                    avisoErro(msg_erro_aux);
                }
            }
        }
        #endregion

        #region [ getDescricaoTCodigoDescricao ]
        private string getDescricaoTCodigoDescricao(string grupo, string codigo)
        {
            #region [ Declarações ]
            string strSql;
            string strDescricao = "";
            SqlCommand cmCommand;
            SqlDataReader drConsulta;
            #endregion

            if (string.IsNullOrEmpty(grupo)) return strDescricao;
            if (string.IsNullOrEmpty(codigo)) return strDescricao;

            strSql = "SELECT descricao FROM t_CODIGO_DESCRICAO WHERE (grupo='" + grupo + "') AND (codigo='" + codigo + "')";

            cmCommand = FMain.contextoBD.AmbienteBase.BD.criaSqlCommand();
            cmCommand.CommandText = strSql;
            drConsulta = cmCommand.ExecuteReader();
            try
            {
                if (drConsulta.Read())
                    strDescricao = drConsulta.GetString(0);
            }
            finally
            {
                drConsulta.Close();
            }

            return strDescricao;
        }
        #endregion

        #region [ getDescricaoStatusMagento ]
        private string getDescricaoStatusMagento(string status)
        {
            string statusDescricao;

            if ((status ?? "").Trim().Length == 0) return "";

            switch (status)
            {
                case "separando":
                    statusDescricao = "Liberar";
                    break;
                case "aguardando_nf_ic":
                    statusDescricao = "Aguardando NF IC";
                    break;
                case "processing":
                    statusDescricao = "Boleto Emitido";
                    break;
                case "separando2":
                    statusDescricao = "Separando";
                    break;
                case "lembrete":
                    statusDescricao = "Lembrete Pagamento";
                    break;
                case "boleto_pago":
                    statusDescricao = "Pago";
                    break;
                case "pgto_auth":
                    statusDescricao = "Pagamento Autorizado";
                    break;
                case "pagto_aprovado_integra":
                    statusDescricao = "Pagto Aprovado IC";
                    break;
                case "pending_payment":
                    statusDescricao = "Análise de Pagamento";
                    break;
                case "payment_review":
                    statusDescricao = "Análise de Pagamento";
                    break;
                case "fraud":
                    statusDescricao = "Suspected Fraud";
                    break;
                case "pending":
                    statusDescricao = "Pedido Realizado";
                    break;
                case "holded":
                    statusDescricao = "On Hold";
                    break;
                case "rastreio_ic":
                    statusDescricao = "Rastreio IC";
                    break;
                case "complete":
                    statusDescricao = "Completo";
                    break;
                case "shipexception":
                    statusDescricao = "Falha no Envio IC";
                    break;
                case "delivered":
                    statusDescricao = "Entregue IC";
                    break;
                case "despachado":
                    statusDescricao = "Enviado";
                    break;
                case "closed":
                    statusDescricao = "Estornado";
                    break;
                case "canceled":
                    statusDescricao = "Cancelado";
                    break;
                case "ip_delivery_failed":
                    statusDescricao = "Entrega Falhou";
                    break;
                case "ip_to_be_delivered":
                    statusDescricao = "Saiu para Entrega";
                    break;
                case "aprovado":
                    statusDescricao = "Pagamento Aprovado";
                    break;
                case "paypal_canceled_reversal":
                    statusDescricao = "PayPal Canceled Reversal";
                    break;
                case "ip_delivery_late":
                    statusDescricao = "Atraso na entrega";
                    break;
                case "pending_paypal":
                    statusDescricao = "Pending PayPal";
                    break;
                case "paypal_reversed":
                    statusDescricao = "PayPal Reversed";
                    break;
                case "ip_in_transit":
                    statusDescricao = "Em trânsito";
                    break;
                case "ip_delivered":
                    statusDescricao = "Entregue";
                    break;
                case "ip_shipped":
                    statusDescricao = "Despachado";
                    break;
                default:
                    statusDescricao = "";
                    break;
            }
            return statusDescricao;
        }
        #endregion

        #region [ getPedidoECommerceOrigemFmtSql ]
        private string getPedidoECommerceOrigemFmtSql(string codigo)
        {
            #region [ Declarações ]
            string strSql;
            string strAux;
            StringBuilder sbRetorno = new StringBuilder("");
            SqlCommand cmCommand;
            SqlDataReader drConsulta;
            #endregion

            if (string.IsNullOrEmpty(codigo)) return "";

            strSql = "SELECT codigo FROM t_CODIGO_DESCRICAO WHERE (codigo_pai = '" + codigo + "') AND grupo='PedidoECommerce_Origem'";

            cmCommand = FMain.contextoBD.AmbienteBase.BD.criaSqlCommand();
            cmCommand.CommandText = strSql;
            drConsulta = cmCommand.ExecuteReader();
            try
            {
                while (drConsulta.Read())
                {
                    strAux = "'" + drConsulta.GetString(0) + "'";
                    if (sbRetorno.Length > 0) sbRetorno.Append(", ");
                    sbRetorno.Append(strAux);
                }
            }
            finally
            {
                drConsulta.Close();
            }
            return sbRetorno.ToString();
        }
        #endregion

        #region [ isSalesOrderStatusValido ]
        private bool isSalesOrderStatusValido(string status)
        {
            if ((status ?? "").Trim().Length > 0)
                status = "|" + status + "|";

            if (PEDIDO_MAGENTO_STATUS_VALIDOS.IndexOf(status ?? "") != -1)
                return true;
            return false;
        }
        #endregion

        #region [ preparaSqlCommandPedidoRecebidoParaSim ]
        public void preparaSqlCommandPedidoRecebidoParaSim()
        {
            #region [ Declarações ]
            string strSql;
            #endregion

            strSql = "UPDATE t_PEDIDO SET" +
                " MarketplacePedidoRecebidoRegistradoStatus=@PedidoRecebidoRegistradoStatus," +
                " MarketplacePedidoRecebidoRegistradoDataHora=getdate()," +
                " MarketplacePedidoRecebidoRegistradoUsuario=@usuario" +
                    " WHERE pedido=@pedido";

            _cmCommandPedidoRecebidoParaSim = FMain.contextoBD.AmbienteBase.BD.criaSqlCommand();
            _cmCommandPedidoRecebidoParaSim.CommandText = strSql;
            _cmCommandPedidoRecebidoParaSim.Parameters.Add("@PedidoRecebidoRegistradoStatus", SqlDbType.TinyInt);
            _cmCommandPedidoRecebidoParaSim.Parameters.Add("@usuario", SqlDbType.VarChar);
            _cmCommandPedidoRecebidoParaSim.Parameters.Add("@pedido", SqlDbType.VarChar);
        }
        #endregion

        #region [ setPedidoRecebidoParaSim ]
        public bool setPedidoRecebidoParaSim(string numPedido, out string msg_erro)
        {
            #region [ Declarações ]
            bool blnRetorno = false;
            #endregion

            msg_erro = "";
            if (numPedido.Trim().Length == 0) return false;

            try
            {
                _cmCommandPedidoRecebidoParaSim.Parameters["@PedidoRecebidoRegistradoStatus"].Value = COD_ST_PEDIDO_RECEBIDO_SIM;
                _cmCommandPedidoRecebidoParaSim.Parameters["@usuario"].Value = Global.Usuario.usuario;
                _cmCommandPedidoRecebidoParaSim.Parameters["@pedido"].Value = numPedido;

                blnRetorno = _cmCommandPedidoRecebidoParaSim.ExecuteNonQuery() == 1;
            }
            catch (Exception ex)
            {
                msg_erro = ex.ToString();
                return false;
            }

            return blnRetorno;
        }
        #endregion

        #endregion

        #region [ Eventos ]

        #region [ FIntegracaoMarketplace ]

        #region [ FIntegracaoMarketplace_Load ]
        private void FIntegracaoMarketplace_Load(object sender, EventArgs e)
        {
            bool blnSucesso = false;

            try
            {
                limpaCampos();

                #region [ Combo Transportadora ]
                DataTable dtbTransportadora = FMain.contextoBD.AmbienteBase.comboDAO.criaDtbTransportadoraCombo();
                DataRow rowTransportadora = dtbTransportadora.NewRow();
                rowTransportadora["id"] = "";
                rowTransportadora["id_razao_social"] = "";
                dtbTransportadora.Rows.InsertAt(rowTransportadora, 0);
                cbTransportadora.DataSource = dtbTransportadora;
                cbTransportadora.ValueMember = "id";
                cbTransportadora.DisplayMember = "id_razao_social";
                cbTransportadora.SelectedIndex = -1;
                #endregion

                #region [ Combo Origem do Pedido (Grupo) ]
                DataTable dtbOrigemPedidoGrupo = FMain.contextoBD.AmbienteBase.comboDAO.criaDtbOrigemPedidoGrupoCombo(ComboDAO.eFiltraStAtivo.TODOS);
                DataRow rowOrigemPedidoGrupo = dtbOrigemPedidoGrupo.NewRow();
                rowOrigemPedidoGrupo["codigo"] = "";
                rowOrigemPedidoGrupo["descricao"] = "";
                dtbOrigemPedidoGrupo.Rows.InsertAt(rowOrigemPedidoGrupo, 0);
                cbOrigemPedidoGrupo.DataSource = dtbOrigemPedidoGrupo;
                cbOrigemPedidoGrupo.ValueMember = "codigo";
                cbOrigemPedidoGrupo.DisplayMember = "descricao";
                cbOrigemPedidoGrupo.SelectedIndex = -1;
                #endregion

                #region [ Combo Origem do Pedido ]
                DataTable dtbOrigemPedido = FMain.contextoBD.AmbienteBase.comboDAO.criaDtbOrigemPedidoCombo(ComboDAO.eFiltraStAtivo.TODOS);
                DataRow rowOrigemPedido = dtbOrigemPedido.NewRow();
                rowOrigemPedido["codigo"] = "";
                rowOrigemPedido["descricao"] = "";
                dtbOrigemPedido.Rows.InsertAt(rowOrigemPedido, 0);
                cbOrigemPedido.DataSource = dtbOrigemPedido;
                cbOrigemPedido.ValueMember = "codigo";
                cbOrigemPedido.DisplayMember = "descricao";
                cbOrigemPedido.SelectedIndex = -1;
                #endregion

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

        #region [ FIntegracaoMarketplace_Shown ]
        private void FIntegracaoMarketplace_Shown(object sender, EventArgs e)
        {

        }
        #endregion

        #region [ FIntegracaoMarketplace_FormClosing ]
        private void FIntegracaoMarketplace_FormClosing(object sender, FormClosingEventArgs e)
        {
            FMain.fMain.Location = this.Location;
            FMain.fMain.Visible = true;
            this.Visible = false;
        }
        #endregion

        #endregion

        #region [ Botões ]

        #region [ btnPesquisar ]
        private void btnPesquisar_Click(object sender, EventArgs e)
        {
            executaPesquisa();
        }
        #endregion

        #region [ btnLimpar ]
        private void btnLimpar_Click(object sender, EventArgs e)
        {
            limpaCampos();
        }
        #endregion

        #region [ btnMarcarTodos ]
        private void btnMarcarTodos_Click(object sender, EventArgs e)
        {
            trataBotaoMarcarTodos();
        }
        #endregion

        #region [ btnDesmarcarTodos ]
        private void btnDesmarcarTodos_Click(object sender, EventArgs e)
        {
            trataBotaoDesmarcarTodos();
        }
        #endregion

        #region [ btnConfirma_Click ]
        private void btnConfirma_Click(object sender, EventArgs e)
        {
            trataBotaoConfirma();
        }
        #endregion

        #endregion

        #endregion
    }
}
