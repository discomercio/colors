namespace Financeiro
{
	partial class FBoletoParcelaEdita
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoParcelaEdita));
			this.gboxVencimento = new System.Windows.Forms.GroupBox();
			this.txtVencto = new System.Windows.Forms.TextBox();
			this.lblTitVencto = new System.Windows.Forms.Label();
			this.gboxRateio = new System.Windows.Forms.GroupBox();
			this.lblTotalGridParcelas = new System.Windows.Forms.Label();
			this.lblTitTotalGridParcelas = new System.Windows.Forms.Label();
			this.grdRateio = new System.Windows.Forms.DataGridView();
			this.grdRateio_pedido = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.grdRateio_valor = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.btnCadastrar = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxVencimento.SuspendLayout();
			this.gboxRateio.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdRateio)).BeginInit();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnCadastrar);
			this.pnBotoes.Size = new System.Drawing.Size(500, 55);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnCadastrar, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxRateio);
			this.pnCampos.Controls.Add(this.gboxVencimento);
			this.pnCampos.Size = new System.Drawing.Size(500, 378);
			this.pnCampos.Click += new System.EventHandler(this.pnCampos_Click);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.Location = new System.Drawing.Point(451, 4);
			this.btnFechar.TabIndex = 2;
			// 
			// btnSobre
			// 
			this.btnSobre.Location = new System.Drawing.Point(406, 4);
			this.btnSobre.TabIndex = 1;
			// 
			// gboxVencimento
			// 
			this.gboxVencimento.Controls.Add(this.txtVencto);
			this.gboxVencimento.Controls.Add(this.lblTitVencto);
			this.gboxVencimento.Location = new System.Drawing.Point(77, 21);
			this.gboxVencimento.Name = "gboxVencimento";
			this.gboxVencimento.Size = new System.Drawing.Size(323, 77);
			this.gboxVencimento.TabIndex = 0;
			this.gboxVencimento.TabStop = false;
			this.gboxVencimento.Text = "Data de Vencimento";
			// 
			// txtVencto
			// 
			this.txtVencto.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtVencto.Location = new System.Drawing.Point(141, 34);
			this.txtVencto.Name = "txtVencto";
			this.txtVencto.Size = new System.Drawing.Size(125, 22);
			this.txtVencto.TabIndex = 0;
			this.txtVencto.Text = "01/01/2000";
			this.txtVencto.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtVencto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtVencto_KeyDown);
			this.txtVencto.Leave += new System.EventHandler(this.txtVencto_Leave);
			this.txtVencto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtVencto_KeyPress);
			this.txtVencto.Enter += new System.EventHandler(this.txtVencto_Enter);
			// 
			// lblTitVencto
			// 
			this.lblTitVencto.AutoSize = true;
			this.lblTitVencto.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitVencto.Location = new System.Drawing.Point(56, 37);
			this.lblTitVencto.Name = "lblTitVencto";
			this.lblTitVencto.Size = new System.Drawing.Size(79, 16);
			this.lblTitVencto.TabIndex = 0;
			this.lblTitVencto.Text = "Vencimento";
			// 
			// gboxRateio
			// 
			this.gboxRateio.Controls.Add(this.lblTitTotalGridParcelas);
			this.gboxRateio.Controls.Add(this.lblTotalGridParcelas);
			this.gboxRateio.Controls.Add(this.grdRateio);
			this.gboxRateio.Location = new System.Drawing.Point(77, 117);
			this.gboxRateio.Name = "gboxRateio";
			this.gboxRateio.Size = new System.Drawing.Size(323, 236);
			this.gboxRateio.TabIndex = 1;
			this.gboxRateio.TabStop = false;
			this.gboxRateio.Text = "Valor da Parcela (Rateio)";
			// 
			// lblTotalGridParcelas
			// 
			this.lblTotalGridParcelas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalGridParcelas.Location = new System.Drawing.Point(181, 210);
			this.lblTotalGridParcelas.Name = "lblTotalGridParcelas";
			this.lblTotalGridParcelas.Size = new System.Drawing.Size(120, 13);
			this.lblTotalGridParcelas.TabIndex = 4;
			this.lblTotalGridParcelas.Text = "123.456,99";
			this.lblTotalGridParcelas.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblTitTotalGridParcelas
			// 
			this.lblTitTotalGridParcelas.Location = new System.Drawing.Point(124, 210);
			this.lblTitTotalGridParcelas.Name = "lblTitTotalGridParcelas";
			this.lblTitTotalGridParcelas.Size = new System.Drawing.Size(46, 13);
			this.lblTitTotalGridParcelas.TabIndex = 3;
			this.lblTitTotalGridParcelas.Text = "Total";
			this.lblTitTotalGridParcelas.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// grdRateio
			// 
			this.grdRateio.AllowUserToAddRows = false;
			this.grdRateio.AllowUserToDeleteRows = false;
			this.grdRateio.AllowUserToResizeColumns = false;
			this.grdRateio.AllowUserToResizeRows = false;
			this.grdRateio.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.grdRateio.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.grdRateio.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.grdRateio.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grdRateio.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.grdRateio_pedido,
            this.grdRateio_valor});
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.grdRateio.DefaultCellStyle = dataGridViewCellStyle4;
			this.grdRateio.Location = new System.Drawing.Point(19, 32);
			this.grdRateio.MultiSelect = false;
			this.grdRateio.Name = "grdRateio";
			this.grdRateio.RowHeadersVisible = false;
			this.grdRateio.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
			this.grdRateio.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdRateio.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
			this.grdRateio.ShowEditingIcon = false;
			this.grdRateio.Size = new System.Drawing.Size(285, 175);
			this.grdRateio.TabIndex = 0;
			this.grdRateio.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.grdRateio_CellEndEdit);
			// 
			// grdRateio_pedido
			// 
			this.grdRateio_pedido.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdRateio_pedido.DefaultCellStyle = dataGridViewCellStyle2;
			this.grdRateio_pedido.HeaderText = "Pedido";
			this.grdRateio_pedido.MinimumWidth = 140;
			this.grdRateio_pedido.Name = "grdRateio_pedido";
			this.grdRateio_pedido.ReadOnly = true;
			this.grdRateio_pedido.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.grdRateio_pedido.Width = 140;
			// 
			// grdRateio_valor
			// 
			this.grdRateio_valor.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
			dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.grdRateio_valor.DefaultCellStyle = dataGridViewCellStyle3;
			this.grdRateio_valor.HeaderText = "Valor";
			this.grdRateio_valor.MinimumWidth = 140;
			this.grdRateio_valor.Name = "grdRateio_valor";
			this.grdRateio_valor.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			// 
			// btnCadastrar
			// 
			this.btnCadastrar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnCadastrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCadastrar.Image")));
			this.btnCadastrar.Location = new System.Drawing.Point(361, 4);
			this.btnCadastrar.Name = "btnCadastrar";
			this.btnCadastrar.Size = new System.Drawing.Size(40, 44);
			this.btnCadastrar.TabIndex = 0;
			this.btnCadastrar.TabStop = false;
			this.btnCadastrar.UseVisualStyleBackColor = true;
			this.btnCadastrar.Click += new System.EventHandler(this.btnCadastrar_Click);
			// 
			// FBoletoParcelaEdita
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(500, 475);
			this.KeyPreview = true;
			this.Name = "FBoletoParcelaEdita";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.Shown += new System.EventHandler(this.FBoletoParcelaEdita_Shown);
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.FBoletoParcelaEdita_KeyPress);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FBoletoParcelaEdita_KeyDown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.gboxVencimento.ResumeLayout(false);
			this.gboxVencimento.PerformLayout();
			this.gboxRateio.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdRateio)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxVencimento;
		private System.Windows.Forms.GroupBox gboxRateio;
		private System.Windows.Forms.TextBox txtVencto;
		private System.Windows.Forms.Label lblTitVencto;
		private System.Windows.Forms.DataGridView grdRateio;
		private System.Windows.Forms.DataGridViewTextBoxColumn grdRateio_pedido;
		private System.Windows.Forms.DataGridViewTextBoxColumn grdRateio_valor;
		private System.Windows.Forms.Label lblTotalGridParcelas;
		private System.Windows.Forms.Label lblTitTotalGridParcelas;
		private System.Windows.Forms.Button btnCadastrar;
	}
}
