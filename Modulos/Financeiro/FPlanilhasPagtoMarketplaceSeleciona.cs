#region [ using ]
using System;
using System.Windows.Forms;
#endregion

namespace Financeiro
{
    public partial class FPlanilhasPagtoMarketplaceSeleciona : Form
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

        FPlanilhaPagamentoMarketplaceB2W fPlanilhaPagamentoMarketplaceB2W;
        #endregion

        #region [ Construtor ]
        public FPlanilhasPagtoMarketplaceSeleciona()
        {
            InitializeComponent();

            #region [ Define a cor de fundo de acordo com o ambiente acessado ]
            BackColor = Global.BackColorPainelPadrao;
            #endregion
        }
        #endregion

        #region [ Métodos ]

        #region[ avisoErro ]
        public void avisoErro(string mensagem)
        {
            MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        #endregion

        #region [ limpaCampos ]
        private void limpaCampos()
        {
            cbPlanilhaPagamentoEmpresas.SelectedIndex = -1;
        }
        #endregion

        #region [ trataBotaoAvancar ]
        private void trataBotaoAvancar()
        {
            byte selectedValue;

            if ((cbPlanilhaPagamentoEmpresas.SelectedIndex == -1) || (cbPlanilhaPagamentoEmpresas.SelectedValue.ToString().Trim().Length == 0))
            {
                avisoErro("Selecione a empresa Marketplace!");
                cbPlanilhaPagamentoEmpresas.Focus();
                return;
            }

            selectedValue = (byte)cbPlanilhaPagamentoEmpresas.SelectedValue;

            #region [ B2W ]
            if (selectedValue == Global.Cte.Marketplace.COD_PLANILHA_PAGAMENTO_B2W)
            {
                fPlanilhaPagamentoMarketplaceB2W = new FPlanilhaPagamentoMarketplaceB2W();
                fPlanilhaPagamentoMarketplaceB2W.Location = FMain.fMain.Location;
                fPlanilhaPagamentoMarketplaceB2W.Show();

                if (!fPlanilhaPagamentoMarketplaceB2W.ocorreuExceptionNaInicializacao)
                {
                    this.DialogResult = DialogResult.OK;
                    Close();
                }
            }
            #endregion
        } 
        #endregion

        #endregion

        #region [ Eventos ]

        #region [ FPlanilhasPagtoMarketplaceSeleciona_Load ]
        private void FPlanilhasPagtoMarketplaceSeleciona_Load(object sender, EventArgs e)
        {
            bool blnSucesso = false;

            try
            {
                limpaCampos();

                #region [ Combo Planilha Pagamento Empresas ]
                cbPlanilhaPagamentoEmpresas.DataSource = Global.montaOpcaoPlanilhaPagamentosMarketplace();
                cbPlanilhaPagamentoEmpresas.DisplayMember = "descricao";
                cbPlanilhaPagamentoEmpresas.ValueMember = "codigo";
                cbPlanilhaPagamentoEmpresas.SelectedIndex = -1;
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

        #region [ FPlanilhasPagtoMarketplaceSeleciona_Shown ]
        private void FPlanilhasPagtoMarketplaceSeleciona_Shown(object sender, EventArgs e)
        {
            try
            {
                #region [ Executa rotinas de inicialização ]
                if (!_InicializacaoOk)
                {

                    #region [ Permissão de acesso ao módulo ]

                    #endregion

                    #region [ Combo Planilha Pagamento Empresas ]
                    // Se houver apenas 1 opção, então seleciona
                    if (cbPlanilhaPagamentoEmpresas.Items.Count == 1)
                    {
                        cbPlanilhaPagamentoEmpresas.SelectedIndex = 0;
                        trataBotaoAvancar();
                    }
                    #endregion

                    #region [ Posiciona foco ]
                    cbPlanilhaPagamentoEmpresas.Focus();
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
                if (!_InicializacaoOk) Close();
            }
        }

        #endregion

        #region [ FPlanilhasPagtoMarketplaceSeleciona_FormClosing ]
        private void FPlanilhasPagtoMarketplaceSeleciona_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK)
            {
                FMain.fMain.Visible = true;
            }
            this.Visible = false;
        }
        #endregion

        #region [ btnCancela_Click ]
        private void btnCancela_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            Close();
        }
        #endregion

        #region [ btnAvancar_Click ]
        private void btnAvancar_Click(object sender, EventArgs e)
        {
            trataBotaoAvancar();
        }
        #endregion

        #endregion
    }
}
