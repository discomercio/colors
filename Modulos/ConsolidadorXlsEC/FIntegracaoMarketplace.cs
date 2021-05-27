#region [ using ]
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
#endregion

namespace ConsolidadorXlsEC 
{
    public partial class FIntegracaoMarketplace : FModelo
    {
		// ATENÇÃO: Sempre que for adicionada uma nova coluna ao grid, deve-se verificar se o seguinte comando continua presente em FIntegracaoMarketplace.Designer.cs
		// ========        this.grdDados.AutoGenerateColumns = false;
		// Percebeu-se que ao incluir uma nova coluna no grid, essa configuração AutoGenerateColumns pode ser removida automaticamente pelo Visual Studio, causando
		// um problema quando o grid é limpo (várias colunas desaparecem) e carregado novamente posteriormente (as colunas são exibidas com os nomes dos campos da
		// consulta SQL no header).
		// Por precaução, a propriedade AutoGenerateColumns passou a ser configurada também na inicialização do form no evento Shown.
		// Além disso, verificar se há necessidade de incluir essa mesma coluna no grid existente no form FConfirmaPedidoStatus

		#region [ Constantes ]
        public const string COD_ST_PEDIDO_RECEBIDO_NAO = "0";
        public const string COD_ST_PEDIDO_RECEBIDO_SIM = "1";
        public const string COD_ST_PEDIDO_RECEBIDO_NAO_DEFINIDO = "10";
		#endregion

		#region [ Parâmetros ]
		public static readonly string PEDIDO_MAGENTO_V1_STATUS_VALIDOS;
		public static readonly string PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_SKYHUB;
		public static readonly string PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_INTEGRACOMMERCE;
		public static readonly string PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_ANYMARKET;
		public static readonly string PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_VENDA_DIRETA;
		public static readonly string PEDIDO_MAGENTO_V2_STATUS_VALIDOS;
		public static readonly string PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_SKYHUB;
		public static readonly string PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_INTEGRACOMMERCE;
		public static readonly string PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_ANYMARKET;
		public static readonly string PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_VENDA_DIRETA;
		public static readonly string ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_SKYHUB;
		public static readonly string ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_INTEGRACOMMERCE;
		public static readonly string ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET;
		public static readonly string ECOMMERCE_PEDIDO_ORIGEM_VENDA_DIRETA;
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
        int _flagPedidoUsarMemorizacaoCompletaEnderecos;
		#endregion

		#region [ Construtor estático ]
		static FIntegracaoMarketplace()
		{
			PEDIDO_MAGENTO_V1_STATUS_VALIDOS = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V1_STATUS_VALIDOS);
			PEDIDO_MAGENTO_V2_STATUS_VALIDOS = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V2_STATUS_VALIDOS);
			ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_SKYHUB = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_SKYHUB);
			ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_INTEGRACOMMERCE = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_INTEGRACOMMERCE);
			ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET);
			ECOMMERCE_PEDIDO_ORIGEM_VENDA_DIRETA = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_ECOMMERCE_PEDIDO_ORIGEM_VENDA_DIRETA);
			PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_SKYHUB = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_SKYHUB);
			PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_INTEGRACOMMERCE = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_INTEGRACOMMERCE);
			PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_ANYMARKET = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_ANYMARKET);
			PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_VENDA_DIRETA = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_VENDA_DIRETA);
			PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_SKYHUB = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_SKYHUB);
			PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_INTEGRACOMMERCE = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_INTEGRACOMMERCE);
			PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_ANYMARKET = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_ANYMARKET);
			PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_VENDA_DIRETA = FMain.contextoBD.AmbienteBase.geralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_CXLSEC_INTEGRACAOMKTP_PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_VENDA_DIRETA);
		}
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
			  " AND (t_PEDIDO__BASE.marketplace_codigo_origem IS NOT NULL) AND (LEN(Coalesce(t_PEDIDO__BASE.marketplace_codigo_origem,'')) > 0)";
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
                strAux = " t_PEDIDO__BASE.marketplace_codigo_origem IN(" + strAux + ")";
                if (sbWhere.Length > 0) sbWhere.Append(" AND");
                sbWhere.Append(strAux);
            }
            #endregion

            #region [ Origem do Pedido ]
            if ((cbOrigemPedido.SelectedIndex > -1) && (cbOrigemPedido.SelectedValue.ToString().Length > 0))
            {
                strAux = " (t_PEDIDO__BASE.marketplace_codigo_origem = '" + cbOrigemPedido.SelectedValue.ToString() + "')";
                if (sbWhere.Length > 0) sbWhere.Append(" AND");
                sbWhere.Append(strAux);
            }
            #endregion

            #region [ Loja ]
            if (txtLoja.Text.Trim().Length > 0)
            {
                strAux = " (t_PEDIDO__BASE.numero_loja = " + txtLoja.Text + ")";
                if (sbWhere.Length > 0) sbWhere.Append(" AND");
                sbWhere.Append(strAux);
            }
			#endregion

			#region [ Plataforma ]
			if (cbPlataforma.SelectedIndex > -1)
			{
				if (cbPlataforma.SelectedValue.ToString().Equals(Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON.ToString()))
				{
					strAux = " (t_PEDIDO__BASE.pedido_bs_x_ac LIKE '" + FMain.lojaLoginParameters.magento_api_rest_prefixo_num_magento + "%')";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
				else
				{
					strAux = " (t_PEDIDO__BASE.pedido_bs_x_ac NOT LIKE '" + FMain.lojaLoginParameters.magento_api_rest_prefixo_num_magento + "%')";
					if (sbWhere.Length > 0) sbWhere.Append(" AND");
					sbWhere.Append(strAux);
				}
			}
			#endregion

			return sbWhere.ToString();
        }
        #endregion

        #region [ montaSqlConsulta ]
        private string montaSqlConsulta()
        {
            #region [ Declarações ]
            string strCidade;
            string strUf;
            string strNome;
            string strWhere;
            string strSql;
            #endregion

            #region [ Monta cláusula Where ]
            strWhere = montaClausulaWhere();
            if (strWhere.Length > 0) strWhere = " WHERE " + strWhere;
            #endregion

            #region [ Monta Select ]
            if (_flagPedidoUsarMemorizacaoCompletaEnderecos == 0)
            {
                strCidade = " c.cidade";
                strUf = " c.uf";
                strNome = " c.nome_iniciais_em_maiusculas";
            }
            else
            {
                strCidade = " (CASE p.st_memorizacao_completa_enderecos WHEN 0 THEN c.cidade ELSE p.endereco_cidade END)";
                strUf = " (CASE p.st_memorizacao_completa_enderecos WHEN 0 THEN c.uf ELSE p.endereco_uf END)";
                strNome = " (CASE p.st_memorizacao_completa_enderecos WHEN 0 THEN c.nome_iniciais_em_maiusculas ELSE dbo.SqlClrUtilIniciaisEmMaiusculas(p.endereco_nome) END)";
            }

            strSql = "SELECT" +
                        " p.transportadora_id," +
                        " p.pedido," +
						" t_PEDIDO__BASE.pedido_bs_x_ac," +
						" t_PEDIDO__BASE.pedido_bs_x_marketplace," +
						" t_PEDIDO__BASE.marketplace_codigo_origem," +
						" t_PEDIDO__BASE.loja," +
                        " p.MarketplacePedidoRecebidoRegistrarDataRecebido," +
                        strCidade + " AS cidade," +
                        strUf + " AS uf," +
                        strNome + " AS nome_iniciais_em_maiusculas," +
                        " Sum(tPI.qtde*tPI.preco_venda) AS vl_pedido," +
						" (SELECT descricao FROM t_CODIGO_DESCRICAO WHERE grupo = 'PedidoECommerce_Origem' AND codigo = t_PEDIDO__BASE.marketplace_codigo_origem) AS marketplace_codigo_origem_descricao," +
						" (SELECT codigo_pai FROM t_CODIGO_DESCRICAO WHERE grupo = 'PedidoECommerce_Origem' AND codigo = t_PEDIDO__BASE.marketplace_codigo_origem) AS marketplace_codigo_origem_pai" +
                    " FROM t_PEDIDO p" +
					" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (p.pedido_base=t_PEDIDO__BASE.pedido)"+
					" INNER JOIN t_PEDIDO_ITEM tPI ON (p.pedido=tPI.pedido)" +
                    " INNER JOIN t_CLIENTE c ON (p.id_cliente=c.id)" +
                        strWhere +
                    " GROUP BY p.transportadora_id" +
                       " ,p.pedido" +
					   " ,t_PEDIDO__BASE.pedido_bs_x_ac" +
					   " ,t_PEDIDO__BASE.pedido_bs_x_marketplace" +
					   " ,t_PEDIDO__BASE.marketplace_codigo_origem" +
					   " ,t_PEDIDO__BASE.loja" +
                       " ,p.MarketplacePedidoRecebidoRegistrarDataRecebido" +
                       " ," + strCidade +
                       " ," + strUf +
                       " ," + strNome +
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
				#region [ Consistências ]
				if (cbPlataforma.SelectedIndex==-1)
				{
					avisoErro("Selecione a Plataforma!");
					return;
				}
				#endregion

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
			bool blnRequisicaoApiOk;
			bool blnSalesOrderStatusValido;
			int intCounter = 0;
			int versaoPlataforma;
			string xmlReqSoap;
			string xmlRespSoap;
			string sessionId = "";
			string strMsgErro;
			string msg_erro;
			string msg_erro_aux;
			string msg_erro_api;
			string strOrigemPedidoAux;
			string strIncrementId = "";
			string strEntityId = "";
			string strStatus = "";
			string strComment = "";
			string sJson = null;
			string urlParamReqRest;
			string respJson;
			string urlBaseAddress = "";
			List<DataGridViewRow> linhasSelecionadas = new List<DataGridViewRow>();
			List<DataGridViewRow> salesOrderInfoComStatusOk = new List<DataGridViewRow>();
			List<DataGridViewRow> salesOrderInfoComStatusInvalido = new List<DataGridViewRow>();
			SalesOrderInfo salesOrderInfoAux = new SalesOrderInfo();
			SalesOrderAddCommentRequest addCommentRequest;
			SalesOrderAddCommentResponse addCommentResponse;
			Magento2AddComment mage2AddCommentRequest;
			FConfirmaPedidoStatus fConfirmaPedidoStatus;
			DialogResult drConfirmaPedidoStatus;
			Magento2SalesOrderInfo mage2SalesOrderInfo;
			HttpResponseMessage response;
			#endregion

			#region [ Inicialização ]
			if (FMain.lojaLoginParameters == null)
			{
				strMsgErro = "Falha ao tentar recuperar os parâmetros de login da API do Magento para a loja " + FMain.contextoBD.AmbienteBase.NumeroLojaArclube + "!";
				avisoErro(strMsgErro);
				return;
			}

			versaoPlataforma = (int)Global.converteInteiro(cbPlataforma.SelectedValue.ToString());
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
						(ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) == -1) &&
						(ECOMMERCE_PEDIDO_ORIGEM_VENDA_DIRETA.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) == -1))
					{
						avisoErro("Não é possível selecionar pedidos que não sejam SkyHub (" + ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_SKYHUB.Trim('|').Replace("|", ", ") + ") ou IntegraCommerce " +
							"(" + ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_INTEGRACOMMERCE.Trim('|').Replace(" | ", ", ") +
							") ou AnyMarket (" + ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET.Trim('|').Replace(" | ", ", ") +
							") ou Venda direta (" + ECOMMERCE_PEDIDO_ORIGEM_VENDA_DIRETA.Trim('|').Replace(" | ", ", ") +
							")!");
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

			if (versaoPlataforma == Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML)
			{
				#region [ Requisição para efetuar login ]
				xmlReqSoap = Magento.montaRequisicaoLogin(FMain.lojaLoginParameters.magento_api_username, FMain.lojaLoginParameters.magento_api_password);
				blnRequisicaoApiOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.login, out xmlRespSoap, out msg_erro_aux);
				if (!blnRequisicaoApiOk)
				{
					info(ModoExibicaoMensagemRodape.Normal);
					avisoErro("Erro ao efetuar logon na API do Magento: \n\n" + msg_erro_aux);
					return;
				}
				#endregion

				#region [ Obtém a SessionId ]
				sessionId = Magento.obtemSessionIdFromLoginResponse(xmlRespSoap, out msg_erro_aux);

				if ((sessionId ?? "").Length == 0)
				{
					info(ModoExibicaoMensagemRodape.Normal);
					avisoErro("Falha ao tentar obter o SessionId!!");
					return;
				}
				#endregion
			}

			try // Finally: encerra sessão Magento (SOAP API)
			{
				foreach (DataGridViewRow item in linhasSelecionadas)
				{
					intCounter++;
					info(ModoExibicaoMensagemRodape.EmExecucao, "verificando status dos pedidos no magento: " + intCounter + " de " + linhasSelecionadas.Count);

					#region [ Recupera o status dos pedidos ]
					if (versaoPlataforma == Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML)
					{
						xmlReqSoap = Magento.montaRequisicaoCallSalesOrderInfo(sessionId, item.Cells[colGrdDadosNumMagento.Name].Value.ToString());
						blnRequisicaoApiOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.call, out xmlRespSoap, out msg_erro_api);
						if (blnRequisicaoApiOk)
						{
							salesOrderInfoAux = Magento.decodificaXmlSalesOrderInfoResponse(xmlRespSoap, out msg_erro_aux);
						}
					}
					else
					{
						mage2SalesOrderInfo = Magento2.getSalesOrderInfo(item.Cells[colGrdDadosNumMagento.Name].Value.ToString(), FMain.lojaLoginParameters, out sJson, out msg_erro_api);
						blnRequisicaoApiOk = (mage2SalesOrderInfo != null);
						if (blnRequisicaoApiOk)
						{
							salesOrderInfoAux = Magento2.decodificaSalesOrderInfoMage2ParaMage1(mage2SalesOrderInfo, out msg_erro);
						}
					}
					#endregion

					#region [ Processa o status do pedido ]
					if (blnRequisicaoApiOk)
					{
						if (!salesOrderInfoAux.faultResponse.isFaultResponse)
						{
							if (versaoPlataforma == Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML)
							{
								blnSalesOrderStatusValido = isSalesOrderStatusValido(salesOrderInfoAux.status);
							}
							else
							{
								blnSalesOrderStatusValido = isSalesOrderMagento2StatusValido(salesOrderInfoAux.status);
							}

							if (blnSalesOrderStatusValido)
							{
								item.Cells[colGrdDadosStatus.Name].Value = salesOrderInfoAux.status;
								item.Cells[colGrdDadosStatusDescricao.Name].Value = getDescricaoStatusMagento(salesOrderInfoAux.status);
								// Obs: na conversão dos dados do Magento 2 para Magento 1, o campo 'entity_id' do Magento 2 é armazenado no campo 'order_id' do Magento 1
								item.Cells[colGrdDadosOrderEntityId.Name].Value = salesOrderInfoAux.order_id;
								salesOrderInfoComStatusOk.Add(item);
							}
							else
							{
								item.Cells[colGrdDadosStatus.Name].Value = salesOrderInfoAux.status;
								item.Cells[colGrdDadosStatusDescricao.Name].Value = getDescricaoStatusMagento(salesOrderInfoAux.status);
								item.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
								item.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: status inválido";
								// Obs: na conversão dos dados do Magento 2 para Magento 1, o campo 'entity_id' do Magento 2 é armazenado no campo 'order_id' do Magento 1
								item.Cells[colGrdDadosOrderEntityId.Name].Value = salesOrderInfoAux.order_id;
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
						avisoErro("Erro ao enviar requisição de um dos pedidos!!\n\nA operação será cancelada!!\n\n" + msg_erro_api);
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
						strEntityId = row.Cells[colGrdDadosOrderEntityId.Name].Value.ToString();
						switch (versaoPlataforma)
						{
							case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML:
								strStatus = PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_SKYHUB;
								break;
							case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON:
								strStatus = PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_SKYHUB;
								break;
							default:
								strStatus = "";
								break;
						}
						strComment = "";
						#endregion
					}
					else if (ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_INTEGRACOMMERCE.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) != -1)
					{
						#region [ Tratamento dos pedidos Integra Commerce ]
						strIncrementId = row.Cells[colGrdDadosNumMagento.Name].Value.ToString();
						strEntityId = row.Cells[colGrdDadosOrderEntityId.Name].Value.ToString();
						switch (versaoPlataforma)
						{
							case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML:
								strStatus = PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_INTEGRACOMMERCE;
								break;
							case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON:
								strStatus = PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_INTEGRACOMMERCE;
								break;
							default:
								strStatus = "";
								break;
						}
						strComment = Global.formataDataDdMmYyyyComSeparador(Convert.ToDateTime(row.Cells[colGrdDadosRecebido.Name].Value));
						#endregion
					}
					else if (ECOMMERCE_PEDIDO_ORIGEM_INTEGRACAO_ANYMARKET.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) != -1)
					{
						#region [ Tratamento dos pedidos AnyMarket ]
						strIncrementId = row.Cells[colGrdDadosNumMagento.Name].Value.ToString();
						strEntityId = row.Cells[colGrdDadosOrderEntityId.Name].Value.ToString();
						switch (versaoPlataforma)
						{
							case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML:
								strStatus = PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_ANYMARKET;
								break;
							case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON:
								strStatus = PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_ANYMARKET;
								break;
							default:
								strStatus = "";
								break;
						}
						strComment = "";
						#endregion
					}
					else if (ECOMMERCE_PEDIDO_ORIGEM_VENDA_DIRETA.ToUpper().IndexOf(strOrigemPedidoAux.ToUpper()) != -1)
					{
						#region [ Tratamento dos pedidos de venda direta ]
						strIncrementId = row.Cells[colGrdDadosNumMagento.Name].Value.ToString();
						strEntityId = row.Cells[colGrdDadosOrderEntityId.Name].Value.ToString();
						switch (versaoPlataforma)
						{
							case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML:
								strStatus = PEDIDO_MAGENTO_V1_STATUS_FINALIZACAO_VENDA_DIRETA;
								break;
							case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON:
								strStatus = PEDIDO_MAGENTO_V2_STATUS_FINALIZACAO_VENDA_DIRETA;
								break;
							default:
								strStatus = "";
								break;
						}
						strComment = "";
						#endregion
					}
					else
						continue;

					if (versaoPlataforma == Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML)
					{
						#region [ Enviar requisição AddComment (API SOAP) ]
						addCommentRequest = new SalesOrderAddCommentRequest();
						addCommentRequest.orderIncrementId = strIncrementId;
						addCommentRequest.status = strStatus;
						addCommentRequest.comment = strComment;

						xmlReqSoap = Magento.montaRequisicaoSalesOrderAddComment(sessionId, addCommentRequest);
						blnRequisicaoApiOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.call, out xmlRespSoap, out msg_erro_api);
						if (blnRequisicaoApiOk)
						{
							msg_erro_aux = "";
							addCommentResponse = Magento.decodificaXmlSalesOrderAddCommentResponse(xmlRespSoap, out msg_erro_api);
							if (addCommentResponse.faultResponse.isFaultResponse)
							{
								blnRequisicaoApiOk = false;
								row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
								row.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: falha ao atualizar status no magento \t" + addCommentResponse.faultResponse.faultstring;
							}
						}
						else
						{
							row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
							row.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: falha ao atualizar status no magento \t" + msg_erro_api;
						}
						#endregion
					}
					else
					{
						#region [ Enviar requisição AddComment (API REST) ]
						mage2AddCommentRequest = new Magento2AddComment();
						// No Magento 2, o campo 'parent_id' deve informar o valor do campo 'entity_id' do pedido e não o 'increment_id'
						// Obs: na conversão dos dados do Magento 2 para Magento 1, o campo 'entity_id' do Magento 2 é armazenado no campo 'order_id' do Magento 1
						mage2AddCommentRequest.statusHistory.parent_id = strEntityId;
						mage2AddCommentRequest.statusHistory.status = strStatus;
						mage2AddCommentRequest.statusHistory.is_customer_notified = "0";
						mage2AddCommentRequest.statusHistory.entity_name = "order";
						mage2AddCommentRequest.statusHistory.comment = strComment.Replace("\r\n", "\n");

						urlParamReqRest = Magento2.montaRequisicaoPostSalesOrderAddComment(strEntityId, FMain.lojaLoginParameters.magento_api_rest_endpoint, out urlBaseAddress);
						blnRequisicaoApiOk = Magento2.enviaRequisicaoPost(urlParamReqRest, mage2AddCommentRequest, FMain.lojaLoginParameters.magento_api_rest_access_token, urlBaseAddress, out respJson, out response, out msg_erro_api);
						if (!blnRequisicaoApiOk)
						{
							row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
							row.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: falha ao atualizar status no magento \t" + msg_erro_api;
						}
						#endregion
					}

					#region [ Processa status de retorno da requisição AddComment ]
					if (blnRequisicaoApiOk)
					{
						if (msg_erro_api.Length > 0)
						{
							row.Cells[colGrdDadosMensagemStatus.Name].Style.ForeColor = Color.Red;
							row.Cells[colGrdDadosMensagemStatus.Name].Value += "Falha: falha ao tentar atualizar status no magento \t" + msg_erro_api;
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

				if (versaoPlataforma == Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML)
				{
					#region [ Encerra sessão ]
					xmlReqSoap = Magento.montaRequisicaoEndSession(sessionId);
					blnRequisicaoApiOk = Magento.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Magento.Transacao.endSession, out xmlRespSoap, out msg_erro_aux);
					if (!blnRequisicaoApiOk)
					{
						avisoErro(msg_erro_aux);
					}
					#endregion
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

            if (PEDIDO_MAGENTO_V1_STATUS_VALIDOS.IndexOf(status ?? "") != -1)
                return true;
            return false;
        }
		#endregion

		#region [ isSalesOrderMagento2StatusValido ]
		private bool isSalesOrderMagento2StatusValido(string status)
		{
			if ((status ?? "").Trim().Length > 0)
				status = "|" + status + "|";

			if (PEDIDO_MAGENTO_V2_STATUS_VALIDOS.IndexOf(status ?? "") != -1)
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
			#region [ Declarações ]
			bool blnSucesso = false;
			string strMsgErro;
			#endregion

            try
            {
                limpaCampos();

				#region [ Dados de login no Magento ]
				if (FMain.lojaLoginParameters == null)
				{
					strMsgErro = "Falha ao tentar recuperar os parâmetros de login da API do Magento para a loja " + FMain.contextoBD.AmbienteBase.NumeroLojaArclube + "!";
					throw new Exception(strMsgErro);
				}
				#endregion

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

				#region [ Combo Plataforma ]
				DataTable dtbPlataforma = FMain.contextoBD.AmbienteBase.comboDAO.criaDtbPlataforma();
				cbPlataforma.DataSource = dtbPlataforma;
				cbPlataforma.ValueMember = "codigo";
				cbPlataforma.DisplayMember = "descricao";
				switch (FMain.lojaLoginParameters.magento_api_versao)
				{
					case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML:
						cbPlataforma.SelectedIndex = 0;
						break;
					case Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON:
						cbPlataforma.SelectedIndex = 1;
						break;
					default:
						cbPlataforma.SelectedIndex = -1;
						break;
				}
				#endregion

				_flagPedidoUsarMemorizacaoCompletaEnderecos = FMain.contextoBD.AmbienteBase.geralDAO.getCampoInteiroTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_FLAG_PEDIDO_MEMORIZACAOCOMPLETAENDERECOS, 0);

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
			try
			{
				#region [ Executa rotinas de inicialização ]
				if (!_InicializacaoOk)
				{
					#region [ Posiciona foco ]
					btnDummy.Focus();
					#endregion

					grdDados.AutoGenerateColumns = false;

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
