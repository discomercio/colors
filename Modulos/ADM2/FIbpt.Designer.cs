namespace ADM2
{
	partial class FIbpt
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
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FIbpt));
			this.lblTitulo = new System.Windows.Forms.Label();
			this.openFileDialogIbpt = new System.Windows.Forms.OpenFileDialog();
			this.btnSelecionaArquivo = new System.Windows.Forms.Button();
			this.txtArquivo = new System.Windows.Forms.TextBox();
			this.lblTitArqIbpt = new System.Windows.Forms.Label();
			this.gboxDados = new System.Windows.Forms.GroupBox();
			this.lblTotalRegistros = new System.Windows.Forms.Label();
			this.lblTitTotalRegistros = new System.Windows.Forms.Label();
			this.grdDados = new System.Windows.Forms.DataGridView();
			this.colCodigo = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colEX = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colTabela = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colAliqNac = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colAliqImp = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.colDescricao = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.gboxDetalhesArquivo = new System.Windows.Forms.GroupBox();
			this.lblQtdeRegsLC116 = new System.Windows.Forms.Label();
			this.lblTitQtdeRegsLC116 = new System.Windows.Forms.Label();
			this.lblUltVersaoArqCarregadaBd = new System.Windows.Forms.Label();
			this.lblQtdeRegsTotal = new System.Windows.Forms.Label();
			this.lblQtdeRegsNbs = new System.Windows.Forms.Label();
			this.lblQtdeRegsNcm = new System.Windows.Forms.Label();
			this.lblVersaoArquivo = new System.Windows.Forms.Label();
			this.lblTitQtdeRegsTotal = new System.Windows.Forms.Label();
			this.lblTitUltVersaoArqCarregadaBd = new System.Windows.Forms.Label();
			this.lblTitQtdeRegsNbs = new System.Windows.Forms.Label();
			this.lblTitQtdeRegsNcm = new System.Windows.Forms.Label();
			this.lblTitVersaoArquivo = new System.Windows.Forms.Label();
			this.gboxMsgErro = new System.Windows.Forms.GroupBox();
			this.lbErro = new System.Windows.Forms.ListBox();
			this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
			this.lbMensagem = new System.Windows.Forms.ListBox();
			this.btnCarregaArquivo = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxDados.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdDados)).BeginInit();
			this.gboxDetalhesArquivo.SuspendLayout();
			this.gboxMsgErro.SuspendLayout();
			this.gboxMensagensInformativas.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnCarregaArquivo);
			this.pnBotoes.TabIndex = 0;
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnCarregaArquivo, 0);
			// 
			// btnSobre
			// 
			this.btnSobre.TabIndex = 1;
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.TabIndex = 2;
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxDados);
			this.pnCampos.Controls.Add(this.gboxDetalhesArquivo);
			this.pnCampos.Controls.Add(this.gboxMsgErro);
			this.pnCampos.Controls.Add(this.gboxMensagensInformativas);
			this.pnCampos.Controls.Add(this.btnSelecionaArquivo);
			this.pnCampos.Controls.Add(this.txtArquivo);
			this.pnCampos.Controls.Add(this.lblTitArqIbpt);
			this.pnCampos.Controls.Add(this.lblTitulo);
			this.pnCampos.Size = new System.Drawing.Size(1008, 599);
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
			this.lblTitulo.Text = "IBPT: Carga do Arquivo";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// openFileDialogIbpt
			// 
			this.openFileDialogIbpt.AddExtension = false;
			this.openFileDialogIbpt.Filter = "Arquivo do IBPT|*.CSV|Todos os arquivos|*.*";
			this.openFileDialogIbpt.InitialDirectory = "\\";
			this.openFileDialogIbpt.Title = "Selecionar arquivo do IBPT";
			// 
			// btnSelecionaArquivo
			// 
			this.btnSelecionaArquivo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSelecionaArquivo.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaArquivo.Image")));
			this.btnSelecionaArquivo.Location = new System.Drawing.Point(811, 45);
			this.btnSelecionaArquivo.Name = "btnSelecionaArquivo";
			this.btnSelecionaArquivo.Size = new System.Drawing.Size(39, 25);
			this.btnSelecionaArquivo.TabIndex = 3;
			this.btnSelecionaArquivo.UseVisualStyleBackColor = true;
			this.btnSelecionaArquivo.Click += new System.EventHandler(this.btnSelecionaArquivo_Click);
			// 
			// txtArquivo
			// 
			this.txtArquivo.BackColor = System.Drawing.SystemColors.Window;
			this.txtArquivo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtArquivo.Location = new System.Drawing.Point(109, 48);
			this.txtArquivo.Name = "txtArquivo";
			this.txtArquivo.ReadOnly = true;
			this.txtArquivo.Size = new System.Drawing.Size(695, 20);
			this.txtArquivo.TabIndex = 2;
			this.txtArquivo.DoubleClick += new System.EventHandler(this.txtArquivo_DoubleClick);
			this.txtArquivo.Enter += new System.EventHandler(this.txtArquivo_Enter);
			// 
			// lblTitArqIbpt
			// 
			this.lblTitArqIbpt.AutoSize = true;
			this.lblTitArqIbpt.Location = new System.Drawing.Point(18, 51);
			this.lblTitArqIbpt.Name = "lblTitArqIbpt";
			this.lblTitArqIbpt.Size = new System.Drawing.Size(85, 13);
			this.lblTitArqIbpt.TabIndex = 1;
			this.lblTitArqIbpt.Text = "Arquivo do IBPT";
			// 
			// gboxDados
			// 
			this.gboxDados.Controls.Add(this.lblTotalRegistros);
			this.gboxDados.Controls.Add(this.lblTitTotalRegistros);
			this.gboxDados.Controls.Add(this.grdDados);
			this.gboxDados.Location = new System.Drawing.Point(5, 140);
			this.gboxDados.Name = "gboxDados";
			this.gboxDados.Size = new System.Drawing.Size(994, 246);
			this.gboxDados.TabIndex = 5;
			this.gboxDados.TabStop = false;
			this.gboxDados.Text = "Dados do Arquivo";
			// 
			// lblTotalRegistros
			// 
			this.lblTotalRegistros.AutoSize = true;
			this.lblTotalRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTotalRegistros.Location = new System.Drawing.Point(100, 228);
			this.lblTotalRegistros.Name = "lblTotalRegistros";
			this.lblTotalRegistros.Size = new System.Drawing.Size(28, 13);
			this.lblTotalRegistros.TabIndex = 2;
			this.lblTotalRegistros.Text = "999";
			// 
			// lblTitTotalRegistros
			// 
			this.lblTitTotalRegistros.AutoSize = true;
			this.lblTitTotalRegistros.Location = new System.Drawing.Point(12, 228);
			this.lblTitTotalRegistros.Name = "lblTitTotalRegistros";
			this.lblTitTotalRegistros.Size = new System.Drawing.Size(88, 13);
			this.lblTitTotalRegistros.TabIndex = 1;
			this.lblTitTotalRegistros.Text = "Total de registros";
			// 
			// grdDados
			// 
			this.grdDados.AllowUserToAddRows = false;
			this.grdDados.AllowUserToDeleteRows = false;
			this.grdDados.AllowUserToResizeColumns = false;
			this.grdDados.AllowUserToResizeRows = false;
			this.grdDados.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.grdDados.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter;
			dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.grdDados.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this.grdDados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.grdDados.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colCodigo,
            this.colEX,
            this.colTabela,
            this.colAliqNac,
            this.colAliqImp,
            this.colDescricao});
			dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
			dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window;
			dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
			dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
			this.grdDados.DefaultCellStyle = dataGridViewCellStyle8;
			this.grdDados.Location = new System.Drawing.Point(15, 19);
			this.grdDados.MultiSelect = false;
			this.grdDados.Name = "grdDados";
			this.grdDados.ReadOnly = true;
			this.grdDados.RowHeadersVisible = false;
			this.grdDados.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
			this.grdDados.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.grdDados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.grdDados.ShowEditingIcon = false;
			this.grdDados.Size = new System.Drawing.Size(965, 207);
			this.grdDados.StandardTab = true;
			this.grdDados.TabIndex = 0;
			this.grdDados.SortCompare += new System.Windows.Forms.DataGridViewSortCompareEventHandler(this.grdDados_SortCompare);
			// 
			// colCodigo
			// 
			this.colCodigo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			this.colCodigo.DefaultCellStyle = dataGridViewCellStyle2;
			this.colCodigo.HeaderText = "Código";
			this.colCodigo.MinimumWidth = 100;
			this.colCodigo.Name = "colCodigo";
			this.colCodigo.ReadOnly = true;
			// 
			// colEX
			// 
			this.colEX.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.colEX.DefaultCellStyle = dataGridViewCellStyle3;
			this.colEX.HeaderText = "EX";
			this.colEX.MinimumWidth = 60;
			this.colEX.Name = "colEX";
			this.colEX.ReadOnly = true;
			this.colEX.Width = 60;
			// 
			// colTabela
			// 
			this.colTabela.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
			this.colTabela.DefaultCellStyle = dataGridViewCellStyle4;
			this.colTabela.HeaderText = "Tabela";
			this.colTabela.MinimumWidth = 80;
			this.colTabela.Name = "colTabela";
			this.colTabela.ReadOnly = true;
			this.colTabela.Width = 80;
			// 
			// colAliqNac
			// 
			this.colAliqNac.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
			dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colAliqNac.DefaultCellStyle = dataGridViewCellStyle5;
			this.colAliqNac.HeaderText = "Alíquota Nac";
			this.colAliqNac.MinimumWidth = 110;
			this.colAliqNac.Name = "colAliqNac";
			this.colAliqNac.ReadOnly = true;
			this.colAliqNac.Width = 110;
			// 
			// colAliqImp
			// 
			this.colAliqImp.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
			dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
			dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colAliqImp.DefaultCellStyle = dataGridViewCellStyle6;
			this.colAliqImp.HeaderText = "Alíquota Imp";
			this.colAliqImp.MinimumWidth = 110;
			this.colAliqImp.Name = "colAliqImp";
			this.colAliqImp.ReadOnly = true;
			this.colAliqImp.Width = 110;
			// 
			// colDescricao
			// 
			this.colDescricao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
			dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
			dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.colDescricao.DefaultCellStyle = dataGridViewCellStyle7;
			this.colDescricao.HeaderText = "Descrição";
			this.colDescricao.MinimumWidth = 150;
			this.colDescricao.Name = "colDescricao";
			this.colDescricao.ReadOnly = true;
			// 
			// gboxDetalhesArquivo
			// 
			this.gboxDetalhesArquivo.Controls.Add(this.lblQtdeRegsLC116);
			this.gboxDetalhesArquivo.Controls.Add(this.lblTitQtdeRegsLC116);
			this.gboxDetalhesArquivo.Controls.Add(this.lblUltVersaoArqCarregadaBd);
			this.gboxDetalhesArquivo.Controls.Add(this.lblQtdeRegsTotal);
			this.gboxDetalhesArquivo.Controls.Add(this.lblQtdeRegsNbs);
			this.gboxDetalhesArquivo.Controls.Add(this.lblQtdeRegsNcm);
			this.gboxDetalhesArquivo.Controls.Add(this.lblVersaoArquivo);
			this.gboxDetalhesArquivo.Controls.Add(this.lblTitQtdeRegsTotal);
			this.gboxDetalhesArquivo.Controls.Add(this.lblTitUltVersaoArqCarregadaBd);
			this.gboxDetalhesArquivo.Controls.Add(this.lblTitQtdeRegsNbs);
			this.gboxDetalhesArquivo.Controls.Add(this.lblTitQtdeRegsNcm);
			this.gboxDetalhesArquivo.Controls.Add(this.lblTitVersaoArquivo);
			this.gboxDetalhesArquivo.Location = new System.Drawing.Point(5, 82);
			this.gboxDetalhesArquivo.Name = "gboxDetalhesArquivo";
			this.gboxDetalhesArquivo.Size = new System.Drawing.Size(994, 50);
			this.gboxDetalhesArquivo.TabIndex = 4;
			this.gboxDetalhesArquivo.TabStop = false;
			this.gboxDetalhesArquivo.Text = "Informações do Arquivo";
			// 
			// lblQtdeRegsLC116
			// 
			this.lblQtdeRegsLC116.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblQtdeRegsLC116.Location = new System.Drawing.Point(705, 16);
			this.lblQtdeRegsLC116.Name = "lblQtdeRegsLC116";
			this.lblQtdeRegsLC116.Size = new System.Drawing.Size(77, 13);
			this.lblQtdeRegsLC116.TabIndex = 7;
			this.lblQtdeRegsLC116.Text = "000.000";
			// 
			// lblTitQtdeRegsLC116
			// 
			this.lblTitQtdeRegsLC116.AutoSize = true;
			this.lblTitQtdeRegsLC116.Location = new System.Drawing.Point(591, 16);
			this.lblTitQtdeRegsLC116.Name = "lblTitQtdeRegsLC116";
			this.lblTitQtdeRegsLC116.Size = new System.Drawing.Size(112, 13);
			this.lblTitQtdeRegsLC116.TabIndex = 6;
			this.lblTitQtdeRegsLC116.Text = "Qtde registros LC 116:";
			// 
			// lblUltVersaoArqCarregadaBd
			// 
			this.lblUltVersaoArqCarregadaBd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblUltVersaoArqCarregadaBd.Location = new System.Drawing.Point(218, 32);
			this.lblUltVersaoArqCarregadaBd.Name = "lblUltVersaoArqCarregadaBd";
			this.lblUltVersaoArqCarregadaBd.Size = new System.Drawing.Size(442, 13);
			this.lblUltVersaoArqCarregadaBd.TabIndex = 11;
			this.lblUltVersaoArqCarregadaBd.Text = "0.0.1  (por XXXXXXXXXX em 99/99/9999 99:99)";
			// 
			// lblQtdeRegsTotal
			// 
			this.lblQtdeRegsTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblQtdeRegsTotal.Location = new System.Drawing.Point(907, 16);
			this.lblQtdeRegsTotal.Name = "lblQtdeRegsTotal";
			this.lblQtdeRegsTotal.Size = new System.Drawing.Size(77, 13);
			this.lblQtdeRegsTotal.TabIndex = 9;
			this.lblQtdeRegsTotal.Text = "000.000";
			// 
			// lblQtdeRegsNbs
			// 
			this.lblQtdeRegsNbs.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblQtdeRegsNbs.Location = new System.Drawing.Point(491, 16);
			this.lblQtdeRegsNbs.Name = "lblQtdeRegsNbs";
			this.lblQtdeRegsNbs.Size = new System.Drawing.Size(77, 13);
			this.lblQtdeRegsNbs.TabIndex = 5;
			this.lblQtdeRegsNbs.Text = "000.000";
			// 
			// lblQtdeRegsNcm
			// 
			this.lblQtdeRegsNcm.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblQtdeRegsNcm.Location = new System.Drawing.Point(289, 16);
			this.lblQtdeRegsNcm.Name = "lblQtdeRegsNcm";
			this.lblQtdeRegsNcm.Size = new System.Drawing.Size(77, 13);
			this.lblQtdeRegsNcm.TabIndex = 3;
			this.lblQtdeRegsNcm.Text = "000.000";
			// 
			// lblVersaoArquivo
			// 
			this.lblVersaoArquivo.BackColor = System.Drawing.SystemColors.Control;
			this.lblVersaoArquivo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblVersaoArquivo.Location = new System.Drawing.Point(104, 16);
			this.lblVersaoArquivo.Name = "lblVersaoArquivo";
			this.lblVersaoArquivo.Size = new System.Drawing.Size(77, 13);
			this.lblVersaoArquivo.TabIndex = 1;
			this.lblVersaoArquivo.Text = "0.0.1";
			// 
			// lblTitQtdeRegsTotal
			// 
			this.lblTitQtdeRegsTotal.AutoSize = true;
			this.lblTitQtdeRegsTotal.Location = new System.Drawing.Point(807, 16);
			this.lblTitQtdeRegsTotal.Name = "lblTitQtdeRegsTotal";
			this.lblTitQtdeRegsTotal.Size = new System.Drawing.Size(98, 13);
			this.lblTitQtdeRegsTotal.TabIndex = 8;
			this.lblTitQtdeRegsTotal.Text = "Qtde registros total:";
			// 
			// lblTitUltVersaoArqCarregadaBd
			// 
			this.lblTitUltVersaoArqCarregadaBd.AutoSize = true;
			this.lblTitUltVersaoArqCarregadaBd.Location = new System.Drawing.Point(6, 32);
			this.lblTitUltVersaoArqCarregadaBd.Name = "lblTitUltVersaoArqCarregadaBd";
			this.lblTitUltVersaoArqCarregadaBd.Size = new System.Drawing.Size(210, 13);
			this.lblTitUltVersaoArqCarregadaBd.TabIndex = 10;
			this.lblTitUltVersaoArqCarregadaBd.Text = "Versão do último arquivo carregado no BD:";
			// 
			// lblTitQtdeRegsNbs
			// 
			this.lblTitQtdeRegsNbs.AutoSize = true;
			this.lblTitQtdeRegsNbs.Location = new System.Drawing.Point(389, 16);
			this.lblTitQtdeRegsNbs.Name = "lblTitQtdeRegsNbs";
			this.lblTitQtdeRegsNbs.Size = new System.Drawing.Size(100, 13);
			this.lblTitQtdeRegsNbs.TabIndex = 4;
			this.lblTitQtdeRegsNbs.Text = "Qtde registros NBS:";
			// 
			// lblTitQtdeRegsNcm
			// 
			this.lblTitQtdeRegsNcm.AutoSize = true;
			this.lblTitQtdeRegsNcm.Location = new System.Drawing.Point(185, 16);
			this.lblTitQtdeRegsNcm.Name = "lblTitQtdeRegsNcm";
			this.lblTitQtdeRegsNcm.Size = new System.Drawing.Size(102, 13);
			this.lblTitQtdeRegsNcm.TabIndex = 2;
			this.lblTitQtdeRegsNcm.Text = "Qtde registros NCM:";
			// 
			// lblTitVersaoArquivo
			// 
			this.lblTitVersaoArquivo.AutoSize = true;
			this.lblTitVersaoArquivo.Location = new System.Drawing.Point(6, 16);
			this.lblTitVersaoArquivo.Name = "lblTitVersaoArquivo";
			this.lblTitVersaoArquivo.Size = new System.Drawing.Size(96, 13);
			this.lblTitVersaoArquivo.TabIndex = 0;
			this.lblTitVersaoArquivo.Text = "Versão do arquivo:";
			// 
			// gboxMsgErro
			// 
			this.gboxMsgErro.Controls.Add(this.lbErro);
			this.gboxMsgErro.Location = new System.Drawing.Point(5, 493);
			this.gboxMsgErro.Name = "gboxMsgErro";
			this.gboxMsgErro.Size = new System.Drawing.Size(994, 95);
			this.gboxMsgErro.TabIndex = 7;
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
			this.lbErro.DoubleClick += new System.EventHandler(this.lbErro_DoubleClick);
			// 
			// gboxMensagensInformativas
			// 
			this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
			this.gboxMensagensInformativas.Location = new System.Drawing.Point(5, 392);
			this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
			this.gboxMensagensInformativas.Size = new System.Drawing.Size(994, 95);
			this.gboxMensagensInformativas.TabIndex = 6;
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
			this.lbMensagem.DoubleClick += new System.EventHandler(this.lbMensagem_DoubleClick);
			// 
			// btnCarregaArquivo
			// 
			this.btnCarregaArquivo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnCarregaArquivo.Image = ((System.Drawing.Image)(resources.GetObject("btnCarregaArquivo.Image")));
			this.btnCarregaArquivo.Location = new System.Drawing.Point(869, 4);
			this.btnCarregaArquivo.Name = "btnCarregaArquivo";
			this.btnCarregaArquivo.Size = new System.Drawing.Size(40, 44);
			this.btnCarregaArquivo.TabIndex = 0;
			this.btnCarregaArquivo.TabStop = false;
			this.btnCarregaArquivo.UseVisualStyleBackColor = true;
			this.btnCarregaArquivo.Click += new System.EventHandler(this.btnCarregaArquivo_Click);
			// 
			// FIbpt
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1008, 696);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "FIbpt";
			this.Text = "ADM2  -  1.00 - 01.JUN.2013";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FIbpt_FormClosing);
			this.Load += new System.EventHandler(this.FIbpt_Load);
			this.Shown += new System.EventHandler(this.FIbpt_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.gboxDados.ResumeLayout(false);
			this.gboxDados.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdDados)).EndInit();
			this.gboxDetalhesArquivo.ResumeLayout(false);
			this.gboxDetalhesArquivo.PerformLayout();
			this.gboxMsgErro.ResumeLayout(false);
			this.gboxMensagensInformativas.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.OpenFileDialog openFileDialogIbpt;
		private System.Windows.Forms.Button btnSelecionaArquivo;
		private System.Windows.Forms.TextBox txtArquivo;
		private System.Windows.Forms.Label lblTitArqIbpt;
		private System.Windows.Forms.GroupBox gboxDados;
		private System.Windows.Forms.Label lblTotalRegistros;
		private System.Windows.Forms.Label lblTitTotalRegistros;
		private System.Windows.Forms.DataGridView grdDados;
		private System.Windows.Forms.GroupBox gboxDetalhesArquivo;
		private System.Windows.Forms.Label lblUltVersaoArqCarregadaBd;
		private System.Windows.Forms.Label lblTitUltVersaoArqCarregadaBd;
		private System.Windows.Forms.Label lblQtdeRegsNbs;
		private System.Windows.Forms.Label lblTitQtdeRegsNbs;
		private System.Windows.Forms.Label lblQtdeRegsNcm;
		private System.Windows.Forms.Label lblTitQtdeRegsNcm;
		private System.Windows.Forms.Label lblVersaoArquivo;
		private System.Windows.Forms.Label lblTitVersaoArquivo;
		private System.Windows.Forms.GroupBox gboxMsgErro;
		private System.Windows.Forms.ListBox lbErro;
		private System.Windows.Forms.GroupBox gboxMensagensInformativas;
		private System.Windows.Forms.ListBox lbMensagem;
		private System.Windows.Forms.Label lblQtdeRegsTotal;
		private System.Windows.Forms.Label lblTitQtdeRegsTotal;
		private System.Windows.Forms.Button btnCarregaArquivo;
		private System.Windows.Forms.Label lblQtdeRegsLC116;
		private System.Windows.Forms.Label lblTitQtdeRegsLC116;
		private System.Windows.Forms.DataGridViewTextBoxColumn colCodigo;
		private System.Windows.Forms.DataGridViewTextBoxColumn colEX;
		private System.Windows.Forms.DataGridViewTextBoxColumn colTabela;
		private System.Windows.Forms.DataGridViewTextBoxColumn colAliqNac;
		private System.Windows.Forms.DataGridViewTextBoxColumn colAliqImp;
		private System.Windows.Forms.DataGridViewTextBoxColumn colDescricao;
	}
}
