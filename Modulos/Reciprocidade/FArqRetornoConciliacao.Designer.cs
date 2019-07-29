namespace Reciprocidade
{
    partial class FArqRetornoConciliacao
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FArqRetornoConciliacao));
			this.label1 = new System.Windows.Forms.Label();
			this.btnSelecionaArqRetorno = new System.Windows.Forms.Button();
			this.txtArqRetorno = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.gboxBoletos = new System.Windows.Forms.GroupBox();
			this.lblTitTotalRegistros = new System.Windows.Forms.Label();
			this.grdBoletos = new System.Windows.Forms.DataGridView();
			this.numero_boleto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.data_emissao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.data_vencimento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.valor = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.data_pagamento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.lblTotalRegistros = new System.Windows.Forms.Label();
			this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
			this.lbMensagem = new System.Windows.Forms.ListBox();
			this.gboxMsgErro = new System.Windows.Forms.GroupBox();
			this.lbErro = new System.Windows.Forms.ListBox();
			this.btnCarregaArqRetorno = new System.Windows.Forms.Button();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxBoletos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).BeginInit();
			this.gboxMensagensInformativas.SuspendLayout();
			this.gboxMsgErro.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnCarregaArqRetorno);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnCarregaArqRetorno, 0);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxMsgErro);
			this.pnCampos.Controls.Add(this.gboxMensagensInformativas);
			this.pnCampos.Controls.Add(this.gboxBoletos);
			this.pnCampos.Controls.Add(this.btnSelecionaArqRetorno);
			this.pnCampos.Controls.Add(this.txtArqRetorno);
			this.pnCampos.Controls.Add(this.label2);
			this.pnCampos.Controls.Add(this.label1);
			this.pnCampos.Size = new System.Drawing.Size(1008, 626);
			// 
			// label1
			// 
			this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Image = ((System.Drawing.Image)(resources.GetObject("label1.Image")));
			this.label1.Location = new System.Drawing.Point(-2, 1);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(1008, 40);
			this.label1.TabIndex = 10;
			this.label1.Text = "CONCILIAÇÃO: Carga do Arquivo ";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnSelecionaArqRetorno
			// 
			this.btnSelecionaArqRetorno.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaArqRetorno.Image")));
			this.btnSelecionaArqRetorno.Location = new System.Drawing.Point(814, 47);
			this.btnSelecionaArqRetorno.Name = "btnSelecionaArqRetorno";
			this.btnSelecionaArqRetorno.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaArqRetorno.TabIndex = 13;
			this.btnSelecionaArqRetorno.UseVisualStyleBackColor = true;
			this.btnSelecionaArqRetorno.Click += new System.EventHandler(this.btnSelecionaArqRetorno_Click);
			// 
			// txtArqRetorno
			// 
			this.txtArqRetorno.BackColor = System.Drawing.Color.White;
			this.txtArqRetorno.Location = new System.Drawing.Point(113, 50);
			this.txtArqRetorno.Name = "txtArqRetorno";
			this.txtArqRetorno.ReadOnly = true;
			this.txtArqRetorno.Size = new System.Drawing.Size(695, 20);
			this.txtArqRetorno.TabIndex = 12;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(9, 53);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(99, 13);
			this.label2.TabIndex = 11;
			this.label2.Text = "Arquivo de Retorno";
			// 
			// gboxBoletos
			// 
			this.gboxBoletos.Controls.Add(this.lblTitTotalRegistros);
			this.gboxBoletos.Controls.Add(this.grdBoletos);
			this.gboxBoletos.Controls.Add(this.lblTotalRegistros);
			this.gboxBoletos.Location = new System.Drawing.Point(12, 90);
			this.gboxBoletos.Name = "gboxBoletos";
			this.gboxBoletos.Size = new System.Drawing.Size(987, 262);
			this.gboxBoletos.TabIndex = 15;
			this.gboxBoletos.TabStop = false;
			this.gboxBoletos.Text = "Dados do Arquivo de Retorno";
			// 
			// lblTitTotalRegistros
			// 
			this.lblTitTotalRegistros.AutoSize = true;
			this.lblTitTotalRegistros.Location = new System.Drawing.Point(12, 244);
			this.lblTitTotalRegistros.Name = "lblTitTotalRegistros";
			this.lblTitTotalRegistros.Size = new System.Drawing.Size(88, 13);
			this.lblTitTotalRegistros.TabIndex = 5;
			this.lblTitTotalRegistros.Text = "Total de registros";
			// 
			// grdBoletos
			// 
			this.grdBoletos.AllowUserToAddRows = false;
			this.grdBoletos.AllowUserToDeleteRows = false;
			this.grdBoletos.AllowUserToResizeColumns = false;
			this.grdBoletos.AllowUserToResizeRows = false;
			this.grdBoletos.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.grdBoletos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.grdBoletos.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.grdBoletos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grdBoletos.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.numero_boleto,
            this.data_emissao,
            this.data_vencimento,
            this.valor,
            this.data_pagamento});
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.grdBoletos.DefaultCellStyle = dataGridViewCellStyle7;
			this.grdBoletos.Location = new System.Drawing.Point(15, 19);
			this.grdBoletos.MultiSelect = false;
			this.grdBoletos.Name = "grdBoletos";
			this.grdBoletos.ReadOnly = true;
			this.grdBoletos.RowHeadersVisible = false;
			this.grdBoletos.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
			this.grdBoletos.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdBoletos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grdBoletos.ShowEditingIcon = false;
			this.grdBoletos.Size = new System.Drawing.Size(965, 223);
			this.grdBoletos.StandardTab = true;
			this.grdBoletos.TabIndex = 0;
			// 
			// numero_boleto
			// 
			this.numero_boleto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			this.numero_boleto.DefaultCellStyle = dataGridViewCellStyle2;
			this.numero_boleto.HeaderText = "Nº Título";
			this.numero_boleto.MinimumWidth = 180;
			this.numero_boleto.Name = "numero_boleto";
			this.numero_boleto.ReadOnly = true;
			this.numero_boleto.Width = 200;
			// 
			// data_emissao
			// 
			this.data_emissao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.data_emissao.DefaultCellStyle = dataGridViewCellStyle3;
			this.data_emissao.HeaderText = "Data Emissão";
			this.data_emissao.MinimumWidth = 140;
			this.data_emissao.Name = "data_emissao";
			this.data_emissao.ReadOnly = true;
			this.data_emissao.Width = 200;
			// 
			// data_vencimento
			// 
			this.data_vencimento.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.data_vencimento.DefaultCellStyle = dataGridViewCellStyle4;
			this.data_vencimento.HeaderText = "Data Vencimento";
			this.data_vencimento.MinimumWidth = 140;
			this.data_vencimento.Name = "data_vencimento";
			this.data_vencimento.ReadOnly = true;
			this.data_vencimento.Width = 200;
			// 
			// valor
			// 
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			this.valor.DefaultCellStyle = dataGridViewCellStyle5;
			this.valor.HeaderText = "Valor";
			this.valor.Name = "valor";
			this.valor.ReadOnly = true;
			this.valor.Width = 165;
			// 
			// data_pagamento
			// 
			this.data_pagamento.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.data_pagamento.DefaultCellStyle = dataGridViewCellStyle6;
			this.data_pagamento.HeaderText = "Data Pagamento";
			this.data_pagamento.Name = "data_pagamento";
			this.data_pagamento.ReadOnly = true;
			// 
			// lblTotalRegistros
			// 
			this.lblTotalRegistros.AutoSize = true;
			this.lblTotalRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalRegistros.Location = new System.Drawing.Point(100, 244);
			this.lblTotalRegistros.Name = "lblTotalRegistros";
			this.lblTotalRegistros.Size = new System.Drawing.Size(28, 13);
			this.lblTotalRegistros.TabIndex = 6;
			this.lblTotalRegistros.Text = "999";
			// 
			// gboxMensagensInformativas
			// 
			this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
			this.gboxMensagensInformativas.Location = new System.Drawing.Point(12, 370);
			this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
			this.gboxMensagensInformativas.Size = new System.Drawing.Size(987, 95);
			this.gboxMensagensInformativas.TabIndex = 16;
			this.gboxMensagensInformativas.TabStop = false;
			this.gboxMensagensInformativas.Text = "Mensagens Informativas";
			// 
			// lbMensagem
			// 
			this.lbMensagem.FormattingEnabled = true;
			this.lbMensagem.Location = new System.Drawing.Point(15, 19);
			this.lbMensagem.Name = "lbMensagem";
			this.lbMensagem.ScrollAlwaysVisible = true;
			this.lbMensagem.Size = new System.Drawing.Size(965, 69);
			this.lbMensagem.TabIndex = 0;
			// 
			// gboxMsgErro
			// 
			this.gboxMsgErro.Controls.Add(this.lbErro);
			this.gboxMsgErro.Location = new System.Drawing.Point(12, 480);
			this.gboxMsgErro.Name = "gboxMsgErro";
			this.gboxMsgErro.Size = new System.Drawing.Size(982, 95);
			this.gboxMsgErro.TabIndex = 17;
			this.gboxMsgErro.TabStop = false;
			this.gboxMsgErro.Text = "Mensagens de Erro";
			// 
			// lbErro
			// 
			this.lbErro.ForeColor = System.Drawing.Color.Red;
			this.lbErro.FormattingEnabled = true;
			this.lbErro.Location = new System.Drawing.Point(15, 19);
			this.lbErro.Name = "lbErro";
			this.lbErro.ScrollAlwaysVisible = true;
			this.lbErro.Size = new System.Drawing.Size(965, 69);
			this.lbErro.TabIndex = 0;
			// 
			// btnCarregaArqRetorno
			// 
			this.btnCarregaArqRetorno.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnCarregaArqRetorno.Image = ((System.Drawing.Image)(resources.GetObject("btnCarregaArqRetorno.Image")));
			this.btnCarregaArqRetorno.Location = new System.Drawing.Point(868, 4);
			this.btnCarregaArqRetorno.Name = "btnCarregaArqRetorno";
			this.btnCarregaArqRetorno.Size = new System.Drawing.Size(40, 44);
			this.btnCarregaArqRetorno.TabIndex = 9;
			this.btnCarregaArqRetorno.TabStop = false;
			this.btnCarregaArqRetorno.UseVisualStyleBackColor = true;
			this.btnCarregaArqRetorno.Click += new System.EventHandler(this.btnCarregaArqRetorno_Click);
			// 
			// openFileDialog
			// 
			this.openFileDialog.Filter = "Todos os arquivos|*.*|Arquivo texto|*.txt";
			// 
			// FArqRetornoConciliacao
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1008, 723);
			this.Name = "FArqRetornoConciliacao";
			this.Text = "FArqRetornoConciliacao";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FArqRetornoConciliacao_FormClosing);
			this.Load += new System.EventHandler(this.FArqRetornoConciliacao_Load);
			this.Shown += new System.EventHandler(this.FArqRetornoConciliacao_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxBoletos.ResumeLayout(false);
			this.gboxBoletos.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).EndInit();
			this.gboxMensagensInformativas.ResumeLayout(false);
			this.gboxMsgErro.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSelecionaArqRetorno;
        private System.Windows.Forms.TextBox txtArqRetorno;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox gboxBoletos;
        private System.Windows.Forms.Label lblTotalRegistros;
        private System.Windows.Forms.Label lblTitTotalRegistros;
        private System.Windows.Forms.DataGridView grdBoletos;
        private System.Windows.Forms.GroupBox gboxMensagensInformativas;
        private System.Windows.Forms.ListBox lbMensagem;
        private System.Windows.Forms.GroupBox gboxMsgErro;
        private System.Windows.Forms.ListBox lbErro;
        private System.Windows.Forms.Button btnCarregaArqRetorno;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.DataGridViewTextBoxColumn numero_boleto;
		private System.Windows.Forms.DataGridViewTextBoxColumn data_emissao;
		private System.Windows.Forms.DataGridViewTextBoxColumn data_vencimento;
		private System.Windows.Forms.DataGridViewTextBoxColumn valor;
		private System.Windows.Forms.DataGridViewTextBoxColumn data_pagamento;
    }
}