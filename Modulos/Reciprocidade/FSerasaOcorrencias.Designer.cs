namespace Reciprocidade
{
    partial class FSerasaOcorrencias
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FSerasaOcorrencias));
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.gridDados = new System.Windows.Forms.DataGridView();
			this.lblTitTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.lblTotalizacaoRegistros = new System.Windows.Forms.Label();
			this.btnPesquisar = new System.Windows.Forms.Button();
			this.btnOcorrenciaTratar = new System.Windows.Forms.Button();
			this.id = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.nosso_numero = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dt_emissao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dt_vencto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.vl_titulo = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dt_pagto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.vl_pago = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.primeira_ocorrencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.todas_ocorrencias = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridDados)).BeginInit();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnOcorrenciaTratar);
			this.pnBotoes.Controls.Add(this.btnPesquisar);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPesquisar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnOcorrenciaTratar, 0);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// pnCampos
			// 
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
			this.lblTitulo.TabIndex = 2;
			this.lblTitulo.Text = "Ocorrências";
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
            this.vl_titulo,
            this.dt_pagto,
            this.vl_pago,
            this.primeira_ocorrencia,
            this.todas_ocorrencias});
			dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.gridDados.DefaultCellStyle = dataGridViewCellStyle9;
			this.gridDados.Dock = System.Windows.Forms.DockStyle.Top;
			this.gridDados.Location = new System.Drawing.Point(0, 40);
			this.gridDados.MultiSelect = false;
			this.gridDados.Name = "gridDados";
			this.gridDados.ReadOnly = true;
			dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.gridDados.RowHeadersDefaultCellStyle = dataGridViewCellStyle10;
			this.gridDados.RowHeadersVisible = false;
			this.gridDados.RowHeadersWidth = 15;
			this.gridDados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.gridDados.Size = new System.Drawing.Size(1004, 401);
			this.gridDados.StandardTab = true;
			this.gridDados.TabIndex = 3;
			this.gridDados.DoubleClick += new System.EventHandler(this.gridDados_DoubleClick);
			this.gridDados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.gridDados_KeyDown);
			// 
			// lblTitTotalizacaoRegistros
			// 
			this.lblTitTotalizacaoRegistros.AutoSize = true;
			this.lblTitTotalizacaoRegistros.Location = new System.Drawing.Point(794, 444);
			this.lblTitTotalizacaoRegistros.Name = "lblTitTotalizacaoRegistros";
			this.lblTitTotalizacaoRegistros.Size = new System.Drawing.Size(54, 13);
			this.lblTitTotalizacaoRegistros.TabIndex = 6;
			this.lblTitTotalizacaoRegistros.Text = "Registros:";
			// 
			// lblTotalizacaoRegistros
			// 
			this.lblTotalizacaoRegistros.AutoSize = true;
			this.lblTotalizacaoRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalizacaoRegistros.Location = new System.Drawing.Point(854, 444);
			this.lblTotalizacaoRegistros.Name = "lblTotalizacaoRegistros";
			this.lblTotalizacaoRegistros.Size = new System.Drawing.Size(21, 13);
			this.lblTotalizacaoRegistros.TabIndex = 7;
			this.lblTotalizacaoRegistros.Text = "99";
			// 
			// btnPesquisar
			// 
			this.btnPesquisar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPesquisar.Image = ((System.Drawing.Image)(resources.GetObject("btnPesquisar.Image")));
			this.btnPesquisar.Location = new System.Drawing.Point(822, 3);
			this.btnPesquisar.Name = "btnPesquisar";
			this.btnPesquisar.Size = new System.Drawing.Size(40, 44);
			this.btnPesquisar.TabIndex = 8;
			this.btnPesquisar.TabStop = false;
			this.btnPesquisar.UseVisualStyleBackColor = true;
			this.btnPesquisar.Click += new System.EventHandler(this.btnPesquisar_Click);
			// 
			// btnOcorrenciaTratar
			// 
			this.btnOcorrenciaTratar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnOcorrenciaTratar.Image = ((System.Drawing.Image)(resources.GetObject("btnOcorrenciaTratar.Image")));
			this.btnOcorrenciaTratar.Location = new System.Drawing.Point(868, 3);
			this.btnOcorrenciaTratar.Name = "btnOcorrenciaTratar";
			this.btnOcorrenciaTratar.Size = new System.Drawing.Size(40, 44);
			this.btnOcorrenciaTratar.TabIndex = 9;
			this.btnOcorrenciaTratar.TabStop = false;
			this.btnOcorrenciaTratar.UseVisualStyleBackColor = true;
			this.btnOcorrenciaTratar.Click += new System.EventHandler(this.btnOcorrenciaTratar_Click);
			// 
			// id
			// 
			this.id.HeaderText = "id";
			this.id.Name = "id";
			this.id.ReadOnly = true;
			this.id.Visible = false;
			// 
			// nosso_numero
			// 
			this.nosso_numero.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			this.nosso_numero.DefaultCellStyle = dataGridViewCellStyle2;
			this.nosso_numero.HeaderText = "Nº Título";
			this.nosso_numero.MinimumWidth = 90;
			this.nosso_numero.Name = "nosso_numero";
			this.nosso_numero.ReadOnly = true;
			this.nosso_numero.Width = 120;
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
			// 
			// dt_vencto
			// 
			this.dt_vencto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			this.dt_vencto.DefaultCellStyle = dataGridViewCellStyle4;
			this.dt_vencto.HeaderText = "Data Vencto";
			this.dt_vencto.MinimumWidth = 75;
			this.dt_vencto.Name = "dt_vencto";
			this.dt_vencto.ReadOnly = true;
			// 
			// vl_titulo
			// 
			this.vl_titulo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			this.vl_titulo.DefaultCellStyle = dataGridViewCellStyle5;
			this.vl_titulo.HeaderText = "Valor";
			this.vl_titulo.MinimumWidth = 90;
			this.vl_titulo.Name = "vl_titulo";
			this.vl_titulo.ReadOnly = true;
			this.vl_titulo.Width = 110;
			// 
			// dt_pagto
			// 
			this.dt_pagto.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.dt_pagto.DefaultCellStyle = dataGridViewCellStyle6;
			this.dt_pagto.HeaderText = "Data Pagto";
			this.dt_pagto.MinimumWidth = 75;
			this.dt_pagto.Name = "dt_pagto";
			this.dt_pagto.ReadOnly = true;
			// 
			// vl_pago
			// 
			this.vl_pago.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.vl_pago.DefaultCellStyle = dataGridViewCellStyle7;
			this.vl_pago.HeaderText = "Valor Pago";
			this.vl_pago.MinimumWidth = 75;
			this.vl_pago.Name = "vl_pago";
			this.vl_pago.ReadOnly = true;
			this.vl_pago.Width = 110;
			// 
			// primeira_ocorrencia
			// 
			this.primeira_ocorrencia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			this.primeira_ocorrencia.DefaultCellStyle = dataGridViewCellStyle8;
			this.primeira_ocorrencia.HeaderText = "Ocorrência";
			this.primeira_ocorrencia.Name = "primeira_ocorrencia";
			this.primeira_ocorrencia.ReadOnly = true;
			// 
			// todas_ocorrencias
			// 
			this.todas_ocorrencias.HeaderText = "Todas Ocorrencias";
			this.todas_ocorrencias.Name = "todas_ocorrencias";
			this.todas_ocorrencias.ReadOnly = true;
			this.todas_ocorrencias.Visible = false;
			// 
			// FSerasaOcorrencias
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1008, 562);
			this.Name = "FSerasaOcorrencias";
			this.Text = "FSerasaOcorrencias";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FSerasaOcorrencias_FormClosing);
			this.Shown += new System.EventHandler(this.FSerasaOcorrencias_Shown);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FSerasaOcorrencias_KeyDown);
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
        private System.Windows.Forms.Label lblTotalizacaoRegistros;
		private System.Windows.Forms.Label lblTitTotalizacaoRegistros;
		private System.Windows.Forms.DataGridViewTextBoxColumn id;
		private System.Windows.Forms.DataGridViewTextBoxColumn nosso_numero;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_emissao;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_vencto;
		private System.Windows.Forms.DataGridViewTextBoxColumn vl_titulo;
		private System.Windows.Forms.DataGridViewTextBoxColumn dt_pagto;
		private System.Windows.Forms.DataGridViewTextBoxColumn vl_pago;
		private System.Windows.Forms.DataGridViewTextBoxColumn primeira_ocorrencia;
		private System.Windows.Forms.DataGridViewTextBoxColumn todas_ocorrencias;
    }
}