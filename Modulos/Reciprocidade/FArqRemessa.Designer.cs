namespace Reciprocidade
{
    partial class FArqRemessa
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FArqRemessa));
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtDiretorio = new System.Windows.Forms.TextBox();
			this.btnSelecionaDiretorio = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.lblTitTotalGridBoletos = new System.Windows.Forms.Label();
			this.grdBoletos = new System.Windows.Forms.DataGridView();
			this.id_registro = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.cnpj = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.num_titulo = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.data_emissao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.data_vencimento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.valor = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.data_pagamento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.tipo_ocorrencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.lblTotalGridBoletos = new System.Windows.Forms.Label();
			this.lblTotalRegistros = new System.Windows.Forms.Label();
			this.btnExecutaConsulta = new System.Windows.Forms.Button();
			this.btnGravaArqRemessa = new System.Windows.Forms.Button();
			this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
			this.btnCancelaEnvioTitulo = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).BeginInit();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnCancelaEnvioTitulo);
			this.pnBotoes.Controls.Add(this.btnGravaArqRemessa);
			this.pnBotoes.Controls.Add(this.btnExecutaConsulta);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnExecutaConsulta, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnGravaArqRemessa, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnCancelaEnvioTitulo, 0);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.groupBox1);
			this.pnCampos.Controls.Add(this.btnSelecionaDiretorio);
			this.pnCampos.Controls.Add(this.txtDiretorio);
			this.pnCampos.Controls.Add(this.label3);
			this.pnCampos.Controls.Add(this.label1);
			this.pnCampos.Size = new System.Drawing.Size(1008, 576);
			// 
			// label1
			// 
			this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Image = ((System.Drawing.Image)(resources.GetObject("label1.Image")));
			this.label1.Location = new System.Drawing.Point(-2, 1);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(1008, 40);
			this.label1.TabIndex = 0;
			this.label1.Text = "Geração do Arquivo de Remessa";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(7, 56);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(46, 13);
			this.label3.TabIndex = 2;
			this.label3.Text = "Diretório";
			// 
			// txtDiretorio
			// 
			this.txtDiretorio.Location = new System.Drawing.Point(62, 53);
			this.txtDiretorio.Name = "txtDiretorio";
			this.txtDiretorio.Size = new System.Drawing.Size(615, 20);
			this.txtDiretorio.TabIndex = 4;
			// 
			// btnSelecionaDiretorio
			// 
			this.btnSelecionaDiretorio.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaDiretorio.Image")));
			this.btnSelecionaDiretorio.Location = new System.Drawing.Point(683, 50);
			this.btnSelecionaDiretorio.Name = "btnSelecionaDiretorio";
			this.btnSelecionaDiretorio.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaDiretorio.TabIndex = 5;
			this.btnSelecionaDiretorio.UseVisualStyleBackColor = true;
			this.btnSelecionaDiretorio.Click += new System.EventHandler(this.btnSelecionaDiretorio_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.lblTitTotalGridBoletos);
			this.groupBox1.Controls.Add(this.grdBoletos);
			this.groupBox1.Controls.Add(this.lblTotalGridBoletos);
			this.groupBox1.Controls.Add(this.lblTotalRegistros);
			this.groupBox1.Location = new System.Drawing.Point(10, 87);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(983, 472);
			this.groupBox1.TabIndex = 6;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "SerasaRecipr";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(157, 454);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(51, 13);
			this.label2.TabIndex = 12;
			this.label2.Text = "Registros";
			this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblTitTotalGridBoletos
			// 
			this.lblTitTotalGridBoletos.Location = new System.Drawing.Point(465, 454);
			this.lblTitTotalGridBoletos.Name = "lblTitTotalGridBoletos";
			this.lblTitTotalGridBoletos.Size = new System.Drawing.Size(46, 13);
			this.lblTitTotalGridBoletos.TabIndex = 10;
			this.lblTitTotalGridBoletos.Text = "Total";
			this.lblTitTotalGridBoletos.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// grdBoletos
			// 
			this.grdBoletos.AllowUserToAddRows = false;
			this.grdBoletos.AllowUserToDeleteRows = false;
			this.grdBoletos.AllowUserToResizeColumns = false;
			this.grdBoletos.AllowUserToResizeRows = false;
			this.grdBoletos.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.grdBoletos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.grdBoletos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grdBoletos.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.id_registro,
            this.cnpj,
            this.num_titulo,
            this.data_emissao,
            this.data_vencimento,
            this.valor,
            this.data_pagamento,
            this.tipo_ocorrencia});
			this.grdBoletos.Location = new System.Drawing.Point(15, 19);
			this.grdBoletos.Name = "grdBoletos";
			this.grdBoletos.RowHeadersVisible = false;
			this.grdBoletos.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
			this.grdBoletos.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdBoletos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grdBoletos.Size = new System.Drawing.Size(952, 432);
			this.grdBoletos.TabIndex = 0;
			// 
			// id_registro
			// 
			this.id_registro.HeaderText = "Id";
			this.id_registro.Name = "id_registro";
			this.id_registro.Visible = false;
			// 
			// cnpj
			// 
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			this.cnpj.DefaultCellStyle = dataGridViewCellStyle1;
			this.cnpj.HeaderText = "CNPJ";
			this.cnpj.Name = "cnpj";
			this.cnpj.Width = 140;
			// 
			// num_titulo
			// 
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			this.num_titulo.DefaultCellStyle = dataGridViewCellStyle2;
			this.num_titulo.HeaderText = "Nº Título";
			this.num_titulo.Name = "num_titulo";
			this.num_titulo.Width = 130;
			// 
			// data_emissao
			// 
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.data_emissao.DefaultCellStyle = dataGridViewCellStyle3;
			this.data_emissao.HeaderText = "Data Emissão";
			this.data_emissao.Name = "data_emissao";
			this.data_emissao.Width = 110;
			// 
			// data_vencimento
			// 
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.data_vencimento.DefaultCellStyle = dataGridViewCellStyle4;
			this.data_vencimento.HeaderText = "Data Vencto";
			this.data_vencimento.Name = "data_vencimento";
			this.data_vencimento.Width = 110;
			// 
			// valor
			// 
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			this.valor.DefaultCellStyle = dataGridViewCellStyle5;
			this.valor.HeaderText = "Valor";
			this.valor.Name = "valor";
			this.valor.Width = 130;
			// 
			// data_pagamento
			// 
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.data_pagamento.DefaultCellStyle = dataGridViewCellStyle6;
			this.data_pagamento.HeaderText = "Data Pagto";
			this.data_pagamento.Name = "data_pagamento";
			this.data_pagamento.Width = 110;
			// 
			// tipo_ocorrencia
			// 
			this.tipo_ocorrencia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			this.tipo_ocorrencia.HeaderText = "Ocorrência";
			this.tipo_ocorrencia.Name = "tipo_ocorrencia";
			// 
			// lblTotalGridBoletos
			// 
			this.lblTotalGridBoletos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalGridBoletos.Location = new System.Drawing.Point(517, 454);
			this.lblTotalGridBoletos.Name = "lblTotalGridBoletos";
			this.lblTotalGridBoletos.Size = new System.Drawing.Size(119, 13);
			this.lblTotalGridBoletos.TabIndex = 11;
			this.lblTotalGridBoletos.Text = "123.456,99";
			this.lblTotalGridBoletos.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblTotalRegistros
			// 
			this.lblTotalRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalRegistros.Location = new System.Drawing.Point(214, 454);
			this.lblTotalRegistros.Name = "lblTotalRegistros";
			this.lblTotalRegistros.Size = new System.Drawing.Size(75, 13);
			this.lblTotalRegistros.TabIndex = 13;
			this.lblTotalRegistros.Text = "123.456";
			// 
			// btnExecutaConsulta
			// 
			this.btnExecutaConsulta.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnExecutaConsulta.Image = ((System.Drawing.Image)(resources.GetObject("btnExecutaConsulta.Image")));
			this.btnExecutaConsulta.Location = new System.Drawing.Point(776, 4);
			this.btnExecutaConsulta.Name = "btnExecutaConsulta";
			this.btnExecutaConsulta.Size = new System.Drawing.Size(40, 44);
			this.btnExecutaConsulta.TabIndex = 8;
			this.btnExecutaConsulta.TabStop = false;
			this.btnExecutaConsulta.UseVisualStyleBackColor = true;
			this.btnExecutaConsulta.Click += new System.EventHandler(this.btnExecutaConsulta_Click);
			// 
			// btnGravaArqRemessa
			// 
			this.btnGravaArqRemessa.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnGravaArqRemessa.Image = ((System.Drawing.Image)(resources.GetObject("btnGravaArqRemessa.Image")));
			this.btnGravaArqRemessa.Location = new System.Drawing.Point(868, 4);
			this.btnGravaArqRemessa.Name = "btnGravaArqRemessa";
			this.btnGravaArqRemessa.Size = new System.Drawing.Size(40, 44);
			this.btnGravaArqRemessa.TabIndex = 9;
			this.btnGravaArqRemessa.TabStop = false;
			this.btnGravaArqRemessa.UseVisualStyleBackColor = true;
			this.btnGravaArqRemessa.Click += new System.EventHandler(this.btnGravaArqRemessa_Click);
			// 
			// folderBrowserDialog
			// 
			this.folderBrowserDialog.RootFolder = System.Environment.SpecialFolder.MyComputer;
			// 
			// btnCancelaEnvioTitulo
			// 
			this.btnCancelaEnvioTitulo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnCancelaEnvioTitulo.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelaEnvioTitulo.Image")));
			this.btnCancelaEnvioTitulo.Location = new System.Drawing.Point(822, 4);
			this.btnCancelaEnvioTitulo.Name = "btnCancelaEnvioTitulo";
			this.btnCancelaEnvioTitulo.Size = new System.Drawing.Size(40, 44);
			this.btnCancelaEnvioTitulo.TabIndex = 11;
			this.btnCancelaEnvioTitulo.TabStop = false;
			this.btnCancelaEnvioTitulo.UseVisualStyleBackColor = true;
			this.btnCancelaEnvioTitulo.Click += new System.EventHandler(this.btnCancelaEnvioTitulo_Click);
			// 
			// FArqRemessa
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1008, 673);
			this.Name = "FArqRemessa";
			this.Text = "FArqRemessa";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FArqRemessa_FormClosing);
			this.Load += new System.EventHandler(this.FArqRemessa_Load);
			this.Shown += new System.EventHandler(this.FArqRemessa_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelecionaDiretorio;
        private System.Windows.Forms.TextBox txtDiretorio;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView grdBoletos;
        private System.Windows.Forms.Button btnGravaArqRemessa;
        private System.Windows.Forms.Button btnExecutaConsulta;
        private System.Windows.Forms.Label lblTotalGridBoletos;
        private System.Windows.Forms.Label lblTitTotalGridBoletos;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
		private System.Windows.Forms.Button btnCancelaEnvioTitulo;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblTotalRegistros;
		private System.Windows.Forms.DataGridViewTextBoxColumn id_registro;
		private System.Windows.Forms.DataGridViewTextBoxColumn cnpj;
		private System.Windows.Forms.DataGridViewTextBoxColumn num_titulo;
		private System.Windows.Forms.DataGridViewTextBoxColumn data_emissao;
		private System.Windows.Forms.DataGridViewTextBoxColumn data_vencimento;
		private System.Windows.Forms.DataGridViewTextBoxColumn valor;
		private System.Windows.Forms.DataGridViewTextBoxColumn data_pagamento;
		private System.Windows.Forms.DataGridViewTextBoxColumn tipo_ocorrencia;
    }
}