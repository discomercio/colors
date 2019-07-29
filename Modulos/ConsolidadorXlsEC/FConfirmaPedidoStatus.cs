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

namespace ConsolidadorXlsEC
{
    public partial class FConfirmaPedidoStatus : Form
    {
        #region [ Atributos ]
        public List<DataGridViewRow> PedidosStatusInvalidos { get; set; }
        public List<DataGridViewRow> LinhasSelecionadas { get; set; }
        #endregion

        #region [ Construtor ]
        public FConfirmaPedidoStatus()
        {
            InitializeComponent();

            LinhasSelecionadas = new List<DataGridViewRow>();
        }

        public FConfirmaPedidoStatus(List<DataGridViewRow> PedidosStatusInvalidos)
        {
            InitializeComponent();
            this.PedidosStatusInvalidos = PedidosStatusInvalidos;
            BackColor = Global.BackColorPainelPadrao;
            LinhasSelecionadas = PedidosStatusInvalidos;
        }
        #endregion

        #region [ Métodos ]

        #region[ confirma ]
        public bool confirma(string mensagem)
        {
            return (MessageBox.Show(mensagem, Global.Cte.Aplicativo.NOME_SISTEMA, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes);
        }
        #endregion

        #region [ trataBotaoMarcarTodos ]
        private void trataBotaoMarcarTodos()
        {
            if (grdPedidosConfirma.Rows.Count == 0) return;

            foreach (DataGridViewRow row in grdPedidosConfirma.Rows)
            {
                row.Cells[colGrdDadosCheckBoxConfirma.Name].Value = true;
            }
        }
        #endregion

        #region [ trataBotaoDesmarcarTodos ]
        private void trataBotaoDesmarcarTodos()
        {
            if (grdPedidosConfirma.Rows.Count == 0) return;

            foreach (DataGridViewRow row in grdPedidosConfirma.Rows)
            {
                row.Cells[colGrdDadosCheckBoxConfirma.Name].Value = false;
            }
        }
        #endregion

        #endregion

        #region [ Eventos ]

        #region [ FConfirmaPedidoStatus_Shown ]

        #region [ FIntegracaoMarketplace_Shown ]
        private void FIntegracaoMarketplace_Shown(object sender, EventArgs e)
        {
            int qtdeCells;
            qtdeCells = PedidosStatusInvalidos[0].Cells.Count;

            foreach (DataGridViewRow row in PedidosStatusInvalidos)
            {
                DataGridViewRow novaLinha = (DataGridViewRow)row.Clone();

                for (int i = 0; i < qtdeCells; i++)
                {
                    novaLinha.Cells[i].Value = row.Cells[i].Value;
                }
                grdPedidosConfirma.Rows.Add(novaLinha);
            }

            grdPedidosConfirma.ClearSelection();
        }
        #endregion

        #region [ FConfirmaPedidoStatus_Closing ]
        private void FConfirmaPedidoStatus_Closing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK)
            {
                if (!confirma("Se esta etapa for ignorada, a operação não será concluída\nTem certeza de que deseja sair e cancelar a operação?"))
                {
                    e.Cancel = true;
                    return;
                } 
            }
        }
        #endregion

        #endregion

        #region [ btnOk_Click ]
        private void btnOk_Click(object sender, EventArgs e)
        {
            foreach(DataGridViewRow row in grdPedidosConfirma.Rows)
            {
                if (Convert.ToBoolean(row.Cells[colGrdDadosCheckBoxConfirma.Name].Value) == false)
                {
                    LinhasSelecionadas.RemoveAll(x => x.Cells[colGrdDadosNumPedido.Name].Value == row.Cells[colGrdDadosNumPedido.Name].Value);
                }
            }

            this.DialogResult = DialogResult.OK;
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

        #endregion
    }
}
