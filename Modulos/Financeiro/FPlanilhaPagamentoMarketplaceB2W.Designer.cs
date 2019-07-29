namespace Financeiro
{
    partial class FPlanilhaPagamentoMarketplaceB2W
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FPlanilhaPagamentoMarketplaceB2W));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblTitulo = new System.Windows.Forms.Label();
            this.btnSelecionaArqPlanilha = new System.Windows.Forms.Button();
            this.txtPlanilha = new System.Windows.Forms.TextBox();
            this.lblTitPlanilhaPagamentos = new System.Windows.Forms.Label();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnConfirmar = new System.Windows.Forms.Button();
            this.pnDados = new System.Windows.Forms.Panel();
            this.gridDados = new System.Windows.Forms.DataGridView();
            this.colLinha = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colPedido = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDataPedido = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colValorTotalPedido = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colTipo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colValor = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colObservacao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pnBotoes.SuspendLayout();
            this.pnCampos.SuspendLayout();
            this.pnDados.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridDados)).BeginInit();
            this.SuspendLayout();
            // 
            // pnBotoes
            // 
            this.pnBotoes.Controls.Add(this.btnConfirmar);
            this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnConfirmar, 0);
            // 
            // pnCampos
            // 
            this.pnCampos.Controls.Add(this.pnDados);
            this.pnCampos.Controls.Add(this.btnSelecionaArqPlanilha);
            this.pnCampos.Controls.Add(this.txtPlanilha);
            this.pnCampos.Controls.Add(this.lblTitPlanilhaPagamentos);
            this.pnCampos.Controls.Add(this.lblTitulo);
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
            // 
            // lblTitulo
            // 
            this.lblTitulo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTitulo.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitulo.ForeColor = System.Drawing.Color.Black;
            this.lblTitulo.Image = ((System.Drawing.Image)(resources.GetObject("lblTitulo.Image")));
            this.lblTitulo.Location = new System.Drawing.Point(0, 0);
            this.lblTitulo.Name = "lblTitulo";
            this.lblTitulo.Size = new System.Drawing.Size(1014, 40);
            this.lblTitulo.TabIndex = 2;
            this.lblTitulo.Text = "Planilha de Pagamentos B2W";
            this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnSelecionaArqPlanilha
            // 
            this.btnSelecionaArqPlanilha.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelecionaArqPlanilha.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaArqPlanilha.Image")));
            this.btnSelecionaArqPlanilha.Location = new System.Drawing.Point(965, 50);
            this.btnSelecionaArqPlanilha.Name = "btnSelecionaArqPlanilha";
            this.btnSelecionaArqPlanilha.Size = new System.Drawing.Size(39, 25);
            this.btnSelecionaArqPlanilha.TabIndex = 6;
            this.btnSelecionaArqPlanilha.UseVisualStyleBackColor = true;
            this.btnSelecionaArqPlanilha.Click += new System.EventHandler(this.btnSelecionaArqPlanilha_Click);
            // 
            // txtPlanilha
            // 
            this.txtPlanilha.BackColor = System.Drawing.SystemColors.Window;
            this.txtPlanilha.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPlanilha.Location = new System.Drawing.Point(111, 53);
            this.txtPlanilha.Name = "txtPlanilha";
            this.txtPlanilha.ReadOnly = true;
            this.txtPlanilha.Size = new System.Drawing.Size(848, 20);
            this.txtPlanilha.TabIndex = 5;
            // 
            // lblTitPlanilhaPagamentos
            // 
            this.lblTitPlanilhaPagamentos.AutoSize = true;
            this.lblTitPlanilhaPagamentos.Location = new System.Drawing.Point(8, 56);
            this.lblTitPlanilhaPagamentos.Name = "lblTitPlanilhaPagamentos";
            this.lblTitPlanilhaPagamentos.Size = new System.Drawing.Size(97, 13);
            this.lblTitPlanilhaPagamentos.TabIndex = 4;
            this.lblTitPlanilhaPagamentos.Text = "Selecionar Planilha";
            // 
            // openFileDialog
            // 
            this.openFileDialog.InitialDirectory = "\\";
            // 
            // btnConfirmar
            // 
            this.btnConfirmar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnConfirmar.Image = ((System.Drawing.Image)(resources.GetObject("btnConfirmar.Image")));
            this.btnConfirmar.Location = new System.Drawing.Point(878, 4);
            this.btnConfirmar.Name = "btnConfirmar";
            this.btnConfirmar.Size = new System.Drawing.Size(40, 44);
            this.btnConfirmar.TabIndex = 7;
            this.btnConfirmar.TabStop = false;
            this.btnConfirmar.UseVisualStyleBackColor = true;
            this.btnConfirmar.Click += new System.EventHandler(this.btnConfirmar_Click);
            // 
            // pnDados
            // 
            this.pnDados.Controls.Add(this.gridDados);
            this.pnDados.Location = new System.Drawing.Point(10, 81);
            this.pnDados.Name = "pnDados";
            this.pnDados.Size = new System.Drawing.Size(999, 520);
            this.pnDados.TabIndex = 7;
            // 
            // gridDados
            // 
            this.gridDados.AllowUserToAddRows = false;
            this.gridDados.AllowUserToDeleteRows = false;
            this.gridDados.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridDados.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gridDados.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.gridDados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridDados.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colLinha,
            this.colPedido,
            this.colDataPedido,
            this.colValorTotalPedido,
            this.colTipo,
            this.colValor,
            this.colObservacao});
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gridDados.DefaultCellStyle = dataGridViewCellStyle9;
            this.gridDados.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridDados.Location = new System.Drawing.Point(0, 0);
            this.gridDados.MultiSelect = false;
            this.gridDados.Name = "gridDados";
            this.gridDados.ReadOnly = true;
            this.gridDados.RowHeadersVisible = false;
            this.gridDados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.gridDados.Size = new System.Drawing.Size(999, 520);
            this.gridDados.StandardTab = true;
            this.gridDados.TabIndex = 0;
            // 
            // colLinha
            // 
            this.colLinha.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.colLinha.DataPropertyName = "Linha";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.colLinha.DefaultCellStyle = dataGridViewCellStyle2;
            this.colLinha.HeaderText = "Linha";
            this.colLinha.MinimumWidth = 60;
            this.colLinha.Name = "colLinha";
            this.colLinha.ReadOnly = true;
            this.colLinha.Width = 60;
            // 
            // colPedido
            // 
            this.colPedido.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.colPedido.DataPropertyName = "Pedido";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.colPedido.DefaultCellStyle = dataGridViewCellStyle3;
            this.colPedido.HeaderText = "Pedido";
            this.colPedido.MinimumWidth = 120;
            this.colPedido.Name = "colPedido";
            this.colPedido.ReadOnly = true;
            this.colPedido.Width = 120;
            // 
            // colDataPedido
            // 
            this.colDataPedido.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.colDataPedido.DataPropertyName = "DataPedido";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.colDataPedido.DefaultCellStyle = dataGridViewCellStyle4;
            this.colDataPedido.HeaderText = "Data Pedido";
            this.colDataPedido.MinimumWidth = 100;
            this.colDataPedido.Name = "colDataPedido";
            this.colDataPedido.ReadOnly = true;
            // 
            // colValorTotalPedido
            // 
            this.colValorTotalPedido.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.colValorTotalPedido.DataPropertyName = "ValorTotalPedido";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.colValorTotalPedido.DefaultCellStyle = dataGridViewCellStyle5;
            this.colValorTotalPedido.HeaderText = "Valor Pedido";
            this.colValorTotalPedido.MinimumWidth = 100;
            this.colValorTotalPedido.Name = "colValorTotalPedido";
            this.colValorTotalPedido.ReadOnly = true;
            // 
            // colTipo
            // 
            this.colTipo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.colTipo.DataPropertyName = "Tipo";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.colTipo.DefaultCellStyle = dataGridViewCellStyle6;
            this.colTipo.HeaderText = "Tipo transação";
            this.colTipo.MinimumWidth = 180;
            this.colTipo.Name = "colTipo";
            this.colTipo.ReadOnly = true;
            this.colTipo.Width = 180;
            // 
            // colValor
            // 
            this.colValor.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.colValor.DataPropertyName = "Valor";
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.colValor.DefaultCellStyle = dataGridViewCellStyle7;
            this.colValor.HeaderText = "Valor";
            this.colValor.MinimumWidth = 100;
            this.colValor.Name = "colValor";
            this.colValor.ReadOnly = true;
            // 
            // colObservacao
            // 
            this.colObservacao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colObservacao.DataPropertyName = "Observacao";
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.colObservacao.DefaultCellStyle = dataGridViewCellStyle8;
            this.colObservacao.HeaderText = "Observação";
            this.colObservacao.Name = "colObservacao";
            this.colObservacao.ReadOnly = true;
            // 
            // FPlanilhaPagamentoMarketplaceB2W
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1018, 706);
            this.Name = "FPlanilhaPagamentoMarketplaceB2W";
            this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FPlanilhaPagamentoMarketplaceB2W_FormClosing);
            this.Load += new System.EventHandler(this.FPlanilhaPagamentoMarketplaceB2W_Load);
            this.Shown += new System.EventHandler(this.FPlanilhaPagamentoMarketplaceB2W_Shown);
            this.pnBotoes.ResumeLayout(false);
            this.pnCampos.ResumeLayout(false);
            this.pnCampos.PerformLayout();
            this.pnDados.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridDados)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.Button btnSelecionaArqPlanilha;
        private System.Windows.Forms.TextBox txtPlanilha;
        private System.Windows.Forms.Label lblTitPlanilhaPagamentos;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button btnConfirmar;
        private System.Windows.Forms.Panel pnDados;
        private System.Windows.Forms.DataGridView gridDados;
        private System.Windows.Forms.DataGridViewTextBoxColumn colLinha;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPedido;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDataPedido;
        private System.Windows.Forms.DataGridViewTextBoxColumn colValorTotalPedido;
        private System.Windows.Forms.DataGridViewTextBoxColumn colTipo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colValor;
        private System.Windows.Forms.DataGridViewTextBoxColumn colObservacao;
    }
}