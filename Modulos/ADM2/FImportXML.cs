using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Media;
using System.Text;
using System.Windows.Forms;

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
            const string NOME_DESTA_ROTINA = "executaPesquisa()";
            String strSql;
            String strFrom;
            String strWhere = "";
            //BancoDados bd;
            //SqlConnection cnConexao;
            SqlCommand cmCommand;
            SqlDataAdapter daAdapter;
            DataTable dtbConsulta = new DataTable();
            DataRow rowConsulta;
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
                    "(t_ESTOQUE.data_entrada < '2021-01-31')" +
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
                " t_ESTOQUE.*," +
                " t_ESTOQUE_XML.xml_conteudo," +
                " t_ESTOQUE_XML.xml_prioridade ";

                strSql +=
                strFrom +
                strWhere +
                " ORDER BY t_ESTOQUE.id_estoque, t_ESTOQUE.data_entrada";
                //aviso(strSql);
                #endregion

                #region [ Executa a consulta no BD ]
                cmCommand.CommandText = strSql;
                daAdapter.SelectCommand = cmCommand;
                daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
                daAdapter.Fill(dtbConsulta);
                aviso("dtbConsulta.Rows.Count = " + dtbConsulta.Rows.Count.ToString());
                #endregion

                #region [ Carrega dados no grid ]
                try
                {
                    info(ModoExibicaoMensagemRodape.EmExecucao, "carregando dados no grid para exibição");
                    grdDados.SuspendLayout();

                    grdDados.Rows.Clear();
                    if (dtbConsulta.Rows.Count > 0) grdDados.Rows.Add(dtbConsulta.Rows.Count);

                    for (int i = 0; i < dtbConsulta.Rows.Count; i++)
                    {
                        rowConsulta = dtbConsulta.Rows[i];
                        grdDados.Rows[i].Cells[GRID_COL_ID_ESTOQUE].Value = BD.readToString(rowConsulta["id_estoque"]);
                        grdDados.Rows[i].Cells[GRID_COL_ID_ESTOQUE].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        grdDados.Rows[i].Cells[GRID_COL_DATA_ENTRADA].Value = Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["data_entrada"]));
                        grdDados.Rows[i].Cells[GRID_COL_DATA_ENTRADA].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        grdDados.Rows[i].Cells[GRID_COL_CD].Value = BD.readToString(rowConsulta["id_nfe_emitente"]);
                        grdDados.Rows[i].Cells[GRID_COL_CD].Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        grdDados.Rows[i].Cells[GRID_COL_DOCUMENTO].Value = BD.readToString(rowConsulta["documento"]);
                        grdDados.Rows[i].Cells[GRID_COL_DOCUMENTO].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        grdDados.Rows[i].Cells[GRID_COL_FABRICANTE].Value = BD.readToString(rowConsulta["fabricante"]);
                        grdDados.Rows[i].Cells[GRID_COL_FABRICANTE].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        grdDados.Rows[i].Cells[GRID_COL_DESCRICAO].Value = BD.readToString(rowConsulta["obs"]);
                        grdDados.Rows[i].Cells[GRID_COL_DESCRICAO].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    //#region [ Exibe o grid sem nenhuma linha pré-selecionada ]
                    //for (int i = 0; i < grdDados.Rows.Count; i++)
                    //{
                    //	if (grdDados.Rows[i].Selected) grdDados.Rows[i].Selected = false;
                    //}
                    //#endregion
                }
                finally
                {
                    grdDados.ResumeLayout();
                }
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

        #endregion
    }
}
