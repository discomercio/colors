namespace Reciprocidade
{
    partial class FTitulosConciliacao
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FTitulosConciliacao));
			this.lblTitulo = new System.Windows.Forms.Label();
			this.gridDados = new System.Windows.Forms.DataGridView();
			this.id = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.nosso_numero = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dt_emissao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dt_vencto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dt_pagto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.vl_titulo = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.btnOcorrenciaTratar = new System.Windows.Forms.Button();
			this.btnPesquisar = new System.Windows.Forms.Button();
			this.lblTitTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.lblTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnLocalizar = new System.Windows.Forms.Button();
			this.txtIdSearch = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridDados)).BeginInit();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnExcel);
			this.pnBotoes.Controls.Add(this.btnPesquisar);
			this.pnBotoes.Controls.Add(this.btnOcorrenciaTratar);
			this.pnBotoes.TabIndex = 0;
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnOcorrenciaTratar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPesquisar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnExcel, 0);
			// 
			// btnSobre
			// 
			this.btnSobre.TabIndex = 3;
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.TabIndex = 4;
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.btnLocalizar);
			this.pnCampos.Controls.Add(this.txtIdSearch);
			this.pnCampos.Controls.Add(this.label1);
			this.pnCampos.Controls.Add(this.lblTitTotalizacaoRegistros);
			this.pnCampos.Controls.Add(this.gridDados);
			this.pnCampos.Controls.Add(this.lblTitulo);
			this.pnCampos.Controls.Add(this.lblTotalizacaoRegistros);
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
			this.lblTitulo.Size = new System.Drawing.Size(1004, 40);
			this.lblTitulo.TabIndex = 0;
			this.lblTitulo.Text = "CONCILIAÇÃO: Títulos";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.gridDados.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.gridDados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.gridDados.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.id,
            this.nosso_numero,
            this.dt_emissao,
            this.dt_vencto,
            this.dt_pagto,
            this.vl_titulo});
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.gridDados.DefaultCellStyle = dataGridViewCellStyle7;
			this.gridDados.Dock = System.Windows.Forms.DockStyle.Top;
			this.gridDados.Location = new System.Drawing.Point(0, 40);
			this.gridDados.MultiSelect = false;
			this.gridDados.Name = "gridDados";
			this.gridDados.ReadOnly = true;
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.gridDados.RowHeadersDefaultCellStyle = dataGridViewCellStyle8;
			this.gridDados.RowHeadersVisible = false;
			this.gridDados.RowHeadersWidth = 15;
			this.gridDados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.gridDados.Size = new System.Drawing.Size(1004, 391);
			this.gridDados.StandardTab = true;
			this.gridDados.TabIndex = 1;
			this.gridDados.DoubleClick += new System.EventHandler(this.gridDados_DoubleClick);
			this.gridDados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.gridDados_KeyDown);
			// 
			// id
			// 
			this.id.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			this.id.HeaderText = "ID";
			this.id.Name = "id";
			this.id.ReadOnly = true;
			this.id.Width = 90;
			// 
			// nosso_numero
			// 
			this.nosso_numero.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			this.nosso_numero.DefaultCellStyle = dataGridViewCellStyle2;
			this.nosso_numero.HeaderText = "Nº Título";
			this.nosso_numero.MinimumWidth = 75;
			this.nosso_numero.Name = "nosso_numero";
			this.nosso_numero.ReadOnly = true;
			this.nosso_numero.Width = 231;
			// 
			// dt_emissao
			// 
			this.dt_emissao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.dt_emissao.DefaultCellStyle = dataGridViewCellStyle3;
			this.dt_emissao.HeaderText = "Data Emissão";
			this.dt_emissao.MinimumWidth = 75;
			this.dt_emissao.Name = "dt_emissao";
			this.dt_emissao.ReadOnly = true;
			this.dt_emissao.Width = 150;
			// 
			// dt_vencto
			// 
			this.dt_vencto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.dt_vencto.DefaultCellStyle = dataGridViewCellStyle4;
			this.dt_vencto.HeaderText = "Data Vencimento";
			this.dt_vencto.MinimumWidth = 75;
			this.dt_vencto.Name = "dt_vencto";
			this.dt_vencto.ReadOnly = true;
			this.dt_vencto.Width = 150;
			// 
			// dt_pagto
			// 
			this.dt_pagto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle5.ForeColor = System.Drawing.Color.Green;
			this.dt_pagto.DefaultCellStyle = dataGridViewCellStyle5;
			this.dt_pagto.HeaderText = "Data Pagto";
			this.dt_pagto.MinimumWidth = 75;
			this.dt_pagto.Name = "dt_pagto";
			this.dt_pagto.ReadOnly = true;
			this.dt_pagto.Width = 150;
			// 
			// vl_titulo
			// 
			this.vl_titulo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			this.vl_titulo.DefaultCellStyle = dataGridViewCellStyle6;
			this.vl_titulo.HeaderText = "Valor";
			this.vl_titulo.MinimumWidth = 90;
			this.vl_titulo.Name = "vl_titulo";
			this.vl_titulo.ReadOnly = true;
			// 
			// btnOcorrenciaTratar
			// 
			this.btnOcorrenciaTratar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnOcorrenciaTratar.Image = ((System.Drawing.Image)(resources.GetObject("btnOcorrenciaTratar.Image")));
			this.btnOcorrenciaTratar.Location = new System.Drawing.Point(868, 3);
			this.btnOcorrenciaTratar.Name = "btnOcorrenciaTratar";
			this.btnOcorrenciaTratar.Size = new System.Drawing.Size(40, 44);
			this.btnOcorrenciaTratar.TabIndex = 2;
			this.btnOcorrenciaTratar.TabStop = false;
			this.btnOcorrenciaTratar.UseVisualStyleBackColor = true;
			this.btnOcorrenciaTratar.Click += new System.EventHandler(this.btnOcorrenciaTratar_Click);
			// 
			// btnPesquisar
			// 
			this.btnPesquisar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPesquisar.Image = ((System.Drawing.Image)(resources.GetObject("btnPesquisar.Image")));
			this.btnPesquisar.Location = new System.Drawing.Point(776, 3);
			this.btnPesquisar.Name = "btnPesquisar";
			this.btnPesquisar.Size = new System.Drawing.Size(40, 44);
			this.btnPesquisar.TabIndex = 0;
			this.btnPesquisar.TabStop = false;
			this.btnPesquisar.UseVisualStyleBackColor = true;
			this.btnPesquisar.Click += new System.EventHandler(this.btnPesquisar_Click);
			// 
			// lblTitTotalizacaoRegistros
			// 
			this.lblTitTotalizacaoRegistros.AutoSize = true;
			this.lblTitTotalizacaoRegistros.Location = new System.Drawing.Point(794, 441);
			this.lblTitTotalizacaoRegistros.Name = "lblTitTotalizacaoRegistros";
			this.lblTitTotalizacaoRegistros.Size = new System.Drawing.Size(54, 13);
			this.lblTitTotalizacaoRegistros.TabIndex = 5;
			this.lblTitTotalizacaoRegistros.Text = "Registros:";
			// 
			// lblTotalizacaoRegistros
			// 
			this.lblTotalizacaoRegistros.AutoSize = true;
			this.lblTotalizacaoRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalizacaoRegistros.Location = new System.Drawing.Point(854, 441);
			this.lblTotalizacaoRegistros.Name = "lblTotalizacaoRegistros";
			this.lblTotalizacaoRegistros.Size = new System.Drawing.Size(21, 13);
			this.lblTotalizacaoRegistros.TabIndex = 6;
			this.lblTotalizacaoRegistros.Text = "99";
			// 
			// btnExcel
			// 
			this.btnExcel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.Location = new System.Drawing.Point(822, 3);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(40, 44);
			this.btnExcel.TabIndex = 1;
			this.btnExcel.TabStop = false;
			this.btnExcel.UseVisualStyleBackColor = true;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnLocalizar
			// 
			this.btnLocalizar.Location = new System.Drawing.Point(143, 437);
			this.btnLocalizar.Name = "btnLocalizar";
			this.btnLocalizar.Size = new System.Drawing.Size(64, 23);
			this.btnLocalizar.TabIndex = 4;
			this.btnLocalizar.Text = "Localizar";
			this.btnLocalizar.UseVisualStyleBackColor = true;
			this.btnLocalizar.Click += new System.EventHandler(this.btnLocalizar_Click);
			// 
			// txtIdSearch
			// 
			this.txtIdSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtIdSearch.Location = new System.Drawing.Point(37, 438);
			this.txtIdSearch.Name = "txtIdSearch";
			this.txtIdSearch.Size = new System.Drawing.Size(100, 20);
			this.txtIdSearch.TabIndex = 3;
			this.txtIdSearch.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtIdSearch.Enter += new System.EventHandler(this.txtIdSearch_Enter);
			this.txtIdSearch.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIdSearch_KeyDown);
			this.txtIdSearch.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIdSearch_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(11, 442);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(20, 13);
			this.label1.TabIndex = 2;
			this.label1.Text = "ID";
			// 
			// FTitulosConciliacao
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1008, 562);
			this.Name = "FTitulosConciliacao";
			this.Text = "FTitulosConciliacao";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FTitulosConciliacao_FormClosing);
			this.Shown += new System.EventHandler(this.FTitulosConciliacao_Shown);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FTitulosConciliacao_KeyDown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridDados)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.DataGridView gridDados;
        private System.Windows.Forms.Button btnOcorrenciaTratar;
        private System.Windows.Forms.Button btnPesquisar;
        private System.Windows.Forms.Label lblTitTotalizacaoRegistros;
		private System.Windows.Forms.Label lblTotalizacaoRegistros;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnLocalizar;
		private System.Windows.Forms.TextBox txtIdSearch;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.DataGridViewTextBoxColumn id;
		private System.Windows.Forms.DataGridViewTextBoxColumn nosso_numero;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_emissao;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_vencto;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_pagto;
		private System.Windows.Forms.DataGridViewTextBoxColumn vl_titulo;
    }
}