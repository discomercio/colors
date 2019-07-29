namespace Financeiro
{
	partial class FBoletoArqRemessa
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoArqRemessa));
			this.cbBoletoCedente = new System.Windows.Forms.ComboBox();
			this.lblTitBoletoCedente = new System.Windows.Forms.Label();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.btnExecutaConsulta = new System.Windows.Forms.Button();
			this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
			this.lblTitDiretorio = new System.Windows.Forms.Label();
			this.txtDiretorio = new System.Windows.Forms.TextBox();
			this.btnSelecionaDiretorio = new System.Windows.Forms.Button();
			this.gboxBoletos = new System.Windows.Forms.GroupBox();
			this.lblTitTotalGridBoletos = new System.Windows.Forms.Label();
			this.lblTotalParcelas = new System.Windows.Forms.Label();
			this.lblTitTotalParcelas = new System.Windows.Forms.Label();
			this.lblTotalSerieBoletos = new System.Windows.Forms.Label();
			this.lblTitTotalSerieBoletos = new System.Windows.Forms.Label();
			this.lblTotalGridBoletos = new System.Windows.Forms.Label();
			this.grdBoletos = new System.Windows.Forms.DataGridView();
			this.cliente = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.num_documento = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.parcelas = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.subtotal = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.id_boleto = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.btnGravaArqRemessa = new System.Windows.Forms.Button();
			this.btnCancela = new System.Windows.Forms.Button();
			this.btnDesfazerBoleto = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxBoletos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).BeginInit();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnDesfazerBoleto);
			this.pnBotoes.Controls.Add(this.btnCancela);
			this.pnBotoes.Controls.Add(this.btnGravaArqRemessa);
			this.pnBotoes.Controls.Add(this.btnExecutaConsulta);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnExecutaConsulta, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnGravaArqRemessa, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnCancela, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDesfazerBoleto, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxBoletos);
			this.pnCampos.Controls.Add(this.btnSelecionaDiretorio);
			this.pnCampos.Controls.Add(this.txtDiretorio);
			this.pnCampos.Controls.Add(this.lblTitDiretorio);
			this.pnCampos.Controls.Add(this.lblTitulo);
			this.pnCampos.Controls.Add(this.cbBoletoCedente);
			this.pnCampos.Controls.Add(this.lblTitBoletoCedente);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnSobre
			// 
			this.btnSobre.TabIndex = 4;
			// 
			// cbBoletoCedente
			// 
			this.cbBoletoCedente.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbBoletoCedente.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbBoletoCedente.BackColor = System.Drawing.SystemColors.Window;
			this.cbBoletoCedente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbBoletoCedente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbBoletoCedente.FormattingEnabled = true;
			this.cbBoletoCedente.Location = new System.Drawing.Point(61, 52);
			this.cbBoletoCedente.MaxDropDownItems = 12;
			this.cbBoletoCedente.Name = "cbBoletoCedente";
			this.cbBoletoCedente.Size = new System.Drawing.Size(659, 21);
			this.cbBoletoCedente.TabIndex = 0;
			this.cbBoletoCedente.SelectionChangeCommitted += new System.EventHandler(this.cbBoletoCedente_SelectionChangeCommitted);
			// 
			// lblTitBoletoCedente
			// 
			this.lblTitBoletoCedente.AutoSize = true;
			this.lblTitBoletoCedente.Location = new System.Drawing.Point(8, 55);
			this.lblTitBoletoCedente.Name = "lblTitBoletoCedente";
			this.lblTitBoletoCedente.Size = new System.Drawing.Size(47, 13);
			this.lblTitBoletoCedente.TabIndex = 7;
			this.lblTitBoletoCedente.Text = "Cedente";
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
			this.lblTitulo.TabIndex = 8;
			this.lblTitulo.Text = "Boleto: Geração do Arquivo de Remessa";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExecutaConsulta
			// 
			this.btnExecutaConsulta.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnExecutaConsulta.Image = ((System.Drawing.Image)(resources.GetObject("btnExecutaConsulta.Image")));
			this.btnExecutaConsulta.Location = new System.Drawing.Point(744, 4);
			this.btnExecutaConsulta.Name = "btnExecutaConsulta";
			this.btnExecutaConsulta.Size = new System.Drawing.Size(40, 44);
			this.btnExecutaConsulta.TabIndex = 0;
			this.btnExecutaConsulta.TabStop = false;
			this.btnExecutaConsulta.UseVisualStyleBackColor = true;
			this.btnExecutaConsulta.Click += new System.EventHandler(this.btnExecutaConsulta_Click);
			// 
			// folderBrowserDialog
			// 
			this.folderBrowserDialog.RootFolder = System.Environment.SpecialFolder.MyComputer;
			// 
			// lblTitDiretorio
			// 
			this.lblTitDiretorio.AutoSize = true;
			this.lblTitDiretorio.Location = new System.Drawing.Point(9, 88);
			this.lblTitDiretorio.Name = "lblTitDiretorio";
			this.lblTitDiretorio.Size = new System.Drawing.Size(46, 13);
			this.lblTitDiretorio.TabIndex = 9;
			this.lblTitDiretorio.Text = "Diretório";
			// 
			// txtDiretorio
			// 
			this.txtDiretorio.BackColor = System.Drawing.SystemColors.Window;
			this.txtDiretorio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDiretorio.Location = new System.Drawing.Point(61, 85);
			this.txtDiretorio.Name = "txtDiretorio";
			this.txtDiretorio.ReadOnly = true;
			this.txtDiretorio.Size = new System.Drawing.Size(615, 20);
			this.txtDiretorio.TabIndex = 1;
			this.txtDiretorio.DoubleClick += new System.EventHandler(this.txtDiretorio_DoubleClick);
			this.txtDiretorio.Enter += new System.EventHandler(this.txtDiretorio_Enter);
			// 
			// btnSelecionaDiretorio
			// 
			this.btnSelecionaDiretorio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSelecionaDiretorio.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaDiretorio.Image")));
			this.btnSelecionaDiretorio.Location = new System.Drawing.Point(683, 82);
			this.btnSelecionaDiretorio.Name = "btnSelecionaDiretorio";
			this.btnSelecionaDiretorio.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaDiretorio.TabIndex = 2;
			this.btnSelecionaDiretorio.UseVisualStyleBackColor = true;
			this.btnSelecionaDiretorio.Click += new System.EventHandler(this.btnSelecionaDiretorio_Click);
			// 
			// gboxBoletos
			// 
			this.gboxBoletos.Controls.Add(this.lblTitTotalGridBoletos);
			this.gboxBoletos.Controls.Add(this.lblTotalParcelas);
			this.gboxBoletos.Controls.Add(this.lblTitTotalParcelas);
			this.gboxBoletos.Controls.Add(this.lblTotalSerieBoletos);
			this.gboxBoletos.Controls.Add(this.lblTitTotalSerieBoletos);
			this.gboxBoletos.Controls.Add(this.lblTotalGridBoletos);
			this.gboxBoletos.Controls.Add(this.grdBoletos);
			this.gboxBoletos.Location = new System.Drawing.Point(10, 126);
			this.gboxBoletos.Name = "gboxBoletos";
			this.gboxBoletos.Size = new System.Drawing.Size(995, 469);
			this.gboxBoletos.TabIndex = 3;
			this.gboxBoletos.TabStop = false;
			this.gboxBoletos.Text = "Boletos";
			// 
			// lblTitTotalGridBoletos
			// 
			this.lblTitTotalGridBoletos.Location = new System.Drawing.Point(802, 450);
			this.lblTitTotalGridBoletos.Name = "lblTitTotalGridBoletos";
			this.lblTitTotalGridBoletos.Size = new System.Drawing.Size(46, 13);
			this.lblTitTotalGridBoletos.TabIndex = 1;
			this.lblTitTotalGridBoletos.Text = "Total";
			this.lblTitTotalGridBoletos.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblTotalParcelas
			// 
			this.lblTotalParcelas.AutoSize = true;
			this.lblTotalParcelas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalParcelas.Location = new System.Drawing.Point(497, 450);
			this.lblTotalParcelas.Name = "lblTotalParcelas";
			this.lblTotalParcelas.Size = new System.Drawing.Size(39, 13);
			this.lblTotalParcelas.TabIndex = 6;
			this.lblTotalParcelas.Text = "9.999";
			// 
			// lblTitTotalParcelas
			// 
			this.lblTitTotalParcelas.AutoSize = true;
			this.lblTitTotalParcelas.Location = new System.Drawing.Point(356, 450);
			this.lblTitTotalParcelas.Name = "lblTitTotalParcelas";
			this.lblTitTotalParcelas.Size = new System.Drawing.Size(141, 13);
			this.lblTitTotalParcelas.TabIndex = 5;
			this.lblTitTotalParcelas.Text = "Total de parcelas de boletos";
			// 
			// lblTotalSerieBoletos
			// 
			this.lblTotalSerieBoletos.AutoSize = true;
			this.lblTotalSerieBoletos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalSerieBoletos.Location = new System.Drawing.Point(135, 450);
			this.lblTotalSerieBoletos.Name = "lblTotalSerieBoletos";
			this.lblTotalSerieBoletos.Size = new System.Drawing.Size(28, 13);
			this.lblTotalSerieBoletos.TabIndex = 4;
			this.lblTotalSerieBoletos.Text = "999";
			// 
			// lblTitTotalSerieBoletos
			// 
			this.lblTitTotalSerieBoletos.AutoSize = true;
			this.lblTitTotalSerieBoletos.Location = new System.Drawing.Point(12, 450);
			this.lblTitTotalSerieBoletos.Name = "lblTitTotalSerieBoletos";
			this.lblTitTotalSerieBoletos.Size = new System.Drawing.Size(123, 13);
			this.lblTitTotalSerieBoletos.TabIndex = 3;
			this.lblTitTotalSerieBoletos.Text = "Total de série de boletos";
			// 
			// lblTotalGridBoletos
			// 
			this.lblTotalGridBoletos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalGridBoletos.Location = new System.Drawing.Point(857, 450);
			this.lblTotalGridBoletos.Name = "lblTotalGridBoletos";
			this.lblTotalGridBoletos.Size = new System.Drawing.Size(120, 13);
			this.lblTotalGridBoletos.TabIndex = 2;
			this.lblTotalGridBoletos.Text = "123.456,99";
			this.lblTotalGridBoletos.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// grdBoletos
			// 
			this.grdBoletos.AllowUserToAddRows = false;
			this.grdBoletos.AllowUserToDeleteRows = false;
			this.grdBoletos.AllowUserToResizeColumns = false;
			this.grdBoletos.AllowUserToResizeRows = false;
			this.grdBoletos.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.grdBoletos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.grdBoletos.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
			this.grdBoletos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grdBoletos.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cliente,
            this.num_documento,
            this.parcelas,
            this.subtotal,
            this.id_boleto});
			dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.grdBoletos.DefaultCellStyle = dataGridViewCellStyle12;
			this.grdBoletos.Location = new System.Drawing.Point(15, 19);
			this.grdBoletos.MultiSelect = false;
			this.grdBoletos.Name = "grdBoletos";
			this.grdBoletos.ReadOnly = true;
			this.grdBoletos.RowHeadersVisible = false;
			this.grdBoletos.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
			this.grdBoletos.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdBoletos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grdBoletos.ShowEditingIcon = false;
			this.grdBoletos.Size = new System.Drawing.Size(965, 428);
			this.grdBoletos.StandardTab = true;
			this.grdBoletos.TabIndex = 0;
			// 
			// cliente
			// 
			this.cliente.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle8.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.cliente.DefaultCellStyle = dataGridViewCellStyle8;
			this.cliente.HeaderText = "Cliente";
			this.cliente.MinimumWidth = 120;
			this.cliente.Name = "cliente";
			this.cliente.ReadOnly = true;
			this.cliente.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			// 
			// num_documento
			// 
			this.num_documento.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			dataGridViewCellStyle9.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.num_documento.DefaultCellStyle = dataGridViewCellStyle9;
			this.num_documento.HeaderText = "Nº Documento";
			this.num_documento.MinimumWidth = 120;
			this.num_documento.Name = "num_documento";
			this.num_documento.ReadOnly = true;
			this.num_documento.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.num_documento.Width = 120;
			// 
			// parcelas
			// 
			this.parcelas.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			dataGridViewCellStyle10.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.parcelas.DefaultCellStyle = dataGridViewCellStyle10;
			this.parcelas.HeaderText = "Parcelas";
			this.parcelas.MinimumWidth = 220;
			this.parcelas.Name = "parcelas";
			this.parcelas.ReadOnly = true;
			this.parcelas.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.parcelas.Width = 220;
			// 
			// subtotal
			// 
			this.subtotal.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomRight;
			dataGridViewCellStyle11.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.subtotal.DefaultCellStyle = dataGridViewCellStyle11;
			this.subtotal.HeaderText = "Subtotal";
			this.subtotal.MinimumWidth = 130;
			this.subtotal.Name = "subtotal";
			this.subtotal.ReadOnly = true;
			this.subtotal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.subtotal.Visible = false;
			this.subtotal.Width = 130;
			// 
			// id_boleto
			// 
			this.id_boleto.HeaderText = "id_boleto";
			this.id_boleto.Name = "id_boleto";
			this.id_boleto.ReadOnly = true;
			this.id_boleto.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.id_boleto.Visible = false;
			// 
			// btnGravaArqRemessa
			// 
			this.btnGravaArqRemessa.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnGravaArqRemessa.Image = ((System.Drawing.Image)(resources.GetObject("btnGravaArqRemessa.Image")));
			this.btnGravaArqRemessa.Location = new System.Drawing.Point(789, 4);
			this.btnGravaArqRemessa.Name = "btnGravaArqRemessa";
			this.btnGravaArqRemessa.Size = new System.Drawing.Size(40, 44);
			this.btnGravaArqRemessa.TabIndex = 1;
			this.btnGravaArqRemessa.TabStop = false;
			this.btnGravaArqRemessa.UseVisualStyleBackColor = true;
			this.btnGravaArqRemessa.Click += new System.EventHandler(this.btnGravaArqRemessa_Click);
			// 
			// btnCancela
			// 
			this.btnCancela.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnCancela.Image = ((System.Drawing.Image)(resources.GetObject("btnCancela.Image")));
			this.btnCancela.Location = new System.Drawing.Point(879, 4);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(40, 44);
			this.btnCancela.TabIndex = 3;
			this.btnCancela.TabStop = false;
			this.btnCancela.UseVisualStyleBackColor = true;
			this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
			// 
			// btnDesfazerBoleto
			// 
			this.btnDesfazerBoleto.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnDesfazerBoleto.Image = ((System.Drawing.Image)(resources.GetObject("btnDesfazerBoleto.Image")));
			this.btnDesfazerBoleto.Location = new System.Drawing.Point(834, 4);
			this.btnDesfazerBoleto.Name = "btnDesfazerBoleto";
			this.btnDesfazerBoleto.Size = new System.Drawing.Size(40, 44);
			this.btnDesfazerBoleto.TabIndex = 2;
			this.btnDesfazerBoleto.TabStop = false;
			this.btnDesfazerBoleto.UseVisualStyleBackColor = true;
			this.btnDesfazerBoleto.Click += new System.EventHandler(this.btnDesfazerBoleto_Click);
			// 
			// FBoletoArqRemessa
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1018, 706);
			this.KeyPreview = true;
			this.Name = "FBoletoArqRemessa";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.Load += new System.EventHandler(this.FBoletoArqRemessa_Load);
			this.Shown += new System.EventHandler(this.FBoletoArqRemessa_Shown);
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FBoletoArqRemessa_FormClosing);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FBoletoArqRemessa_KeyDown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxBoletos.ResumeLayout(false);
			this.gboxBoletos.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdBoletos)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.ComboBox cbBoletoCedente;
		private System.Windows.Forms.Label lblTitBoletoCedente;
		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Button btnExecutaConsulta;
		private System.Windows.Forms.TextBox txtDiretorio;
		private System.Windows.Forms.Label lblTitDiretorio;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
		private System.Windows.Forms.Button btnSelecionaDiretorio;
		private System.Windows.Forms.GroupBox gboxBoletos;
		private System.Windows.Forms.Label lblTotalGridBoletos;
		private System.Windows.Forms.Label lblTitTotalGridBoletos;
		private System.Windows.Forms.DataGridView grdBoletos;
		private System.Windows.Forms.Button btnGravaArqRemessa;
		private System.Windows.Forms.Label lblTotalSerieBoletos;
		private System.Windows.Forms.Label lblTitTotalSerieBoletos;
		private System.Windows.Forms.Label lblTotalParcelas;
		private System.Windows.Forms.Label lblTitTotalParcelas;
		private System.Windows.Forms.DataGridViewTextBoxColumn cliente;
		private System.Windows.Forms.DataGridViewTextBoxColumn num_documento;
		private System.Windows.Forms.DataGridViewTextBoxColumn parcelas;
		private System.Windows.Forms.DataGridViewTextBoxColumn subtotal;
		private System.Windows.Forms.DataGridViewTextBoxColumn id_boleto;
		private System.Windows.Forms.Button btnCancela;
		private System.Windows.Forms.Button btnDesfazerBoleto;

	}
}
