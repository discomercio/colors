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
#endregion

namespace EtqFinanceiro
{
    public partial class FModelo : Form
    {
        #region[ Constantes ]
        public const string FmtRelogioData = Global.Cte.DataHora.FmtDia + "." + Global.Cte.DataHora.FmtMes + "." + Global.Cte.DataHora.FmtAno;
        public const string FmtRelogioHora = Global.Cte.DataHora.FmtHora + ":" + Global.Cte.DataHora.FmtMin + ":" + Global.Cte.DataHora.FmtSeg;
        #endregion

        #region [ Enum ]
        public enum ModoExibicaoMensagemRodape
        {
            Normal,
            EmExecucao
        }
        #endregion

        #region [ Métodos públicos ]

        #region[ info ]
        public void info(ModoExibicaoMensagemRodape modo)
        {
            info(modo, "");
        }

        public void info(ModoExibicaoMensagemRodape modo, string mensagem)
        {
            if (mensagem.Trim().Length == 0) mensagem = Global.Cte.Aplicativo.M_ID;
            lblMensagem.Text = mensagem;
            switch (modo)
            {
                case ModoExibicaoMensagemRodape.Normal:
                    lblMensagem.BackColor = System.Drawing.SystemColors.Control;
                    lblMensagem.Font = new Font(this.Font, FontStyle.Regular);
                    break;
                case ModoExibicaoMensagemRodape.EmExecucao:
                    lblMensagem.BackColor = System.Drawing.Color.Yellow;
                    lblMensagem.Font = new Font(this.Font, FontStyle.Bold);
                    break;
                default:
                    lblMensagem.BackColor = System.Drawing.SystemColors.Control;
                    lblMensagem.Font = new Font(this.Font, FontStyle.Regular);
                    break;
            }
            this.Refresh();
        }
        #endregion

        #region[ aviso ]
        public void aviso(string mensagem)
        {
            MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion

        #region[ avisoErro ]
        public void avisoErro(string mensagem)
        {
            MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        #endregion

        #region [ confirma ]
        public bool confirma(string mensagem)
        {
            return (MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes);
        }
        #endregion

        #endregion

        #region [ Construtor ]
        public FModelo()
        {
            InitializeComponent();

            #region [ Define a cor de fundo de acordo com o ambiente acessado ]
            // IMPORTANTE: todos os forms derivados de FModelo automaticamente terão a cor de background ajustadas por esta rotina
            BackColor = Global.BackColorPainelPadrao;
            pnStatus.BackColor = System.Drawing.SystemColors.Control;
            pnBotoes.BackColor = System.Drawing.SystemColors.Control;
            #endregion
        }
        #endregion

        #region [ Eventos ]

        #region [ FModelo ]

        #region [ FModelo_Load ]
        private void FModelo_Load(object sender, EventArgs e)
        {
            btnDummy.Top = -200;
            Text = Global.Cte.Aplicativo.M_ID;
            lblData.Text = DateTime.Now.ToString(FmtRelogioData);
            lblHora.Text = DateTime.Now.ToString(FmtRelogioHora);
            if (Height > Screen.PrimaryScreen.WorkingArea.Height) Height = Screen.PrimaryScreen.WorkingArea.Height;
            info(ModoExibicaoMensagemRodape.Normal);
        }
        #endregion

        #region [ FModelo_FormClosing ]
        private void FModelo_FormClosing(object sender, FormClosingEventArgs e)
        {
            #region [ Se houver um combo-box aberto, não fecha o form em caso de ESC ]
            foreach (Control c in Controls)
            {
                if (c.GetType() == typeof(ComboBox))
                {
                    if(((ComboBox)c).DroppedDown)
                    {
                        ((ComboBox)c).DroppedDown = false;
                        e.Cancel = true;
                        return;
                    }
                }
            }
            #endregion
        }
        #endregion

        #endregion

        #region [ Relógio ]

        #region [ tmrRelogio_Tick ]
        private void tmrRelogio_Tick(object sender, EventArgs e)
        {
            #region [ Data e hora ]
            lblData.Text = DateTime.Now.ToString(FmtRelogioData);
            lblHora.Text = DateTime.Now.ToString(FmtRelogioHora);
            #endregion
        }
        #endregion

        #endregion

        #region [ Botões/Menu ]

        #region [ Fechar ]

        #region [ btnFechar_Click ]
        private void btnFechar_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        #region [ menuArquivoFechar_Click ]
        private void menuArquivoFechar_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        #endregion

        #region [ Sobre ]

        #region [ btnSobre_Click ]
        private void btnSobre_Click(object sender, EventArgs e)
        {
            new PainelSobre().Show();
        }
        #endregion

        #region [ menuAjudaSobre_Click ]
        private void menuAjudaSobre_Click(object sender, EventArgs e)
        {
            new PainelSobre().Show();
        }
        #endregion

        #endregion

        #endregion

        #endregion
    }
}
