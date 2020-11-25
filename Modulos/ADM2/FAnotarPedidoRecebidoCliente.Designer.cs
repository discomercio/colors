namespace ADM2
{
    partial class FAnotarPedidoRecebidoCliente
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FAnotarPedidoRecebidoCliente));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnSelecionaArquivoRastreio = new System.Windows.Forms.Button();
            this.txtArquivoRastreio = new System.Windows.Forms.TextBox();
            this.lblArquivoRastreio = new System.Windows.Forms.Label();
            this.lblTituloPainel = new System.Windows.Forms.Label();
            this.openFileDialogCtrl = new System.Windows.Forms.OpenFileDialog();
            this.btnConfirma = new System.Windows.Forms.Button();
            this.gboxMsgErro = new System.Windows.Forms.GroupBox();
            this.lbErro = new System.Windows.Forms.ListBox();
            this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
            this.lbMensagem = new System.Windows.Forms.ListBox();
            this.lblTotalRegistros = new System.Windows.Forms.Label();
            this.lblTitTotalRegistros = new System.Windows.Forms.Label();
            this.grid = new System.Windows.Forms.DataGridView();
            this.ColVisibleOrdenacaoPadrao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NF = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Destinatario = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Destino = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Situacao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Detalhe = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DataEntrega = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PrevisaoEntrega = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Status = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Mensagem = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColHiddenValorOrdenacaoPadrao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColHiddenNF = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColHiddenDataEntrega = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColHiddenPrevisaoEntrega = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColHiddenGuid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblQtdeRegErro = new System.Windows.Forms.Label();
            this.lblTitQtdeRegErro = new System.Windows.Forms.Label();
            this.lblQtdeRegApto = new System.Windows.Forms.Label();
            this.lblTitQtdeRegApto = new System.Windows.Forms.Label();
            this.lblQtdeAtualizSucesso = new System.Windows.Forms.Label();
            this.lblTitQtdeAtualizSucesso = new System.Windows.Forms.Label();
            this.lblQtdeAtualizFalha = new System.Windows.Forms.Label();
            this.lblTitQtdeAtualizFalha = new System.Windows.Forms.Label();
            this.pnBotoes.SuspendLayout();
            this.pnCampos.SuspendLayout();
            this.gboxMsgErro.SuspendLayout();
            this.gboxMensagensInformativas.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            this.SuspendLayout();
            // 
            // pnBotoes
            // 
            this.pnBotoes.Controls.Add(this.btnConfirma);
            this.pnBotoes.Size = new System.Drawing.Size(1314, 55);
            this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnConfirma, 0);
            // 
            // btnSobre
            // 
            this.btnSobre.Location = new System.Drawing.Point(1215, 4);
            this.btnSobre.TabIndex = 1;
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
            // 
            // btnFechar
            // 
            this.btnFechar.Location = new System.Drawing.Point(1260, 4);
            this.btnFechar.TabIndex = 2;
            // 
            // pnCampos
            // 
            this.pnCampos.Controls.Add(this.lblQtdeAtualizSucesso);
            this.pnCampos.Controls.Add(this.lblTitQtdeAtualizSucesso);
            this.pnCampos.Controls.Add(this.lblQtdeAtualizFalha);
            this.pnCampos.Controls.Add(this.lblTitQtdeAtualizFalha);
            this.pnCampos.Controls.Add(this.lblQtdeRegApto);
            this.pnCampos.Controls.Add(this.lblTitQtdeRegApto);
            this.pnCampos.Controls.Add(this.lblQtdeRegErro);
            this.pnCampos.Controls.Add(this.lblTitQtdeRegErro);
            this.pnCampos.Controls.Add(this.grid);
            this.pnCampos.Controls.Add(this.lblTotalRegistros);
            this.pnCampos.Controls.Add(this.lblTitTotalRegistros);
            this.pnCampos.Controls.Add(this.gboxMsgErro);
            this.pnCampos.Controls.Add(this.gboxMensagensInformativas);
            this.pnCampos.Controls.Add(this.btnSelecionaArquivoRastreio);
            this.pnCampos.Controls.Add(this.txtArquivoRastreio);
            this.pnCampos.Controls.Add(this.lblArquivoRastreio);
            this.pnCampos.Controls.Add(this.lblTituloPainel);
            this.pnCampos.Size = new System.Drawing.Size(1314, 609);
            // 
            // btnSelecionaArquivoRastreio
            // 
            this.btnSelecionaArquivoRastreio.Image = ((System.Drawing.Image)(resources.GetObject("btnSelecionaArquivoRastreio.Image")));
            this.btnSelecionaArquivoRastreio.Location = new System.Drawing.Point(1249, 47);
            this.btnSelecionaArquivoRastreio.Name = "btnSelecionaArquivoRastreio";
            this.btnSelecionaArquivoRastreio.Size = new System.Drawing.Size(39, 25);
            this.btnSelecionaArquivoRastreio.TabIndex = 1;
            this.btnSelecionaArquivoRastreio.UseVisualStyleBackColor = true;
            this.btnSelecionaArquivoRastreio.Click += new System.EventHandler(this.btnSelecionaArquivoRastreio_Click);
            // 
            // txtArquivoRastreio
            // 
            this.txtArquivoRastreio.BackColor = System.Drawing.Color.White;
            this.txtArquivoRastreio.Location = new System.Drawing.Point(94, 50);
            this.txtArquivoRastreio.Name = "txtArquivoRastreio";
            this.txtArquivoRastreio.ReadOnly = true;
            this.txtArquivoRastreio.Size = new System.Drawing.Size(1149, 20);
            this.txtArquivoRastreio.TabIndex = 0;
            // 
            // lblArquivoRastreio
            // 
            this.lblArquivoRastreio.AutoSize = true;
            this.lblArquivoRastreio.Location = new System.Drawing.Point(21, 53);
            this.lblArquivoRastreio.Name = "lblArquivoRastreio";
            this.lblArquivoRastreio.Size = new System.Drawing.Size(67, 13);
            this.lblArquivoRastreio.TabIndex = 22;
            this.lblArquivoRastreio.Text = "Arquivo CSV";
            // 
            // lblTituloPainel
            // 
            this.lblTituloPainel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTituloPainel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTituloPainel.Image = ((System.Drawing.Image)(resources.GetObject("lblTituloPainel.Image")));
            this.lblTituloPainel.Location = new System.Drawing.Point(-2, 1);
            this.lblTituloPainel.Name = "lblTituloPainel";
            this.lblTituloPainel.Size = new System.Drawing.Size(1314, 40);
            this.lblTituloPainel.TabIndex = 21;
            this.lblTituloPainel.Text = "Anotar Pedidos Recebidos pelo Cliente";
            this.lblTituloPainel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // openFileDialogCtrl
            // 
            this.openFileDialogCtrl.AddExtension = false;
            this.openFileDialogCtrl.Filter = "Arquivo CSV|*.csv|Todos os arquivos|*.*";
            this.openFileDialogCtrl.InitialDirectory = "\\";
            // 
            // btnConfirma
            // 
            this.btnConfirma.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnConfirma.Image = ((System.Drawing.Image)(resources.GetObject("btnConfirma.Image")));
            this.btnConfirma.Location = new System.Drawing.Point(1169, 4);
            this.btnConfirma.Name = "btnConfirma";
            this.btnConfirma.Size = new System.Drawing.Size(40, 44);
            this.btnConfirma.TabIndex = 0;
            this.btnConfirma.TabStop = false;
            this.btnConfirma.UseVisualStyleBackColor = true;
            this.btnConfirma.Click += new System.EventHandler(this.btnConfirma_Click);
            // 
            // gboxMsgErro
            // 
            this.gboxMsgErro.Controls.Add(this.lbErro);
            this.gboxMsgErro.Location = new System.Drawing.Point(10, 511);
            this.gboxMsgErro.Name = "gboxMsgErro";
            this.gboxMsgErro.Size = new System.Drawing.Size(1290, 88);
            this.gboxMsgErro.TabIndex = 24;
            this.gboxMsgErro.TabStop = false;
            this.gboxMsgErro.Text = "Mensagens de Erro";
            // 
            // lbErro
            // 
            this.lbErro.ForeColor = System.Drawing.Color.Red;
            this.lbErro.FormattingEnabled = true;
            this.lbErro.Items.AddRange(new object[] {
            "Linha 1",
            "Linha 2",
            "Linha 3",
            "Linha 4",
            "Linha 5",
            "Linha 6",
            "Linha 7",
            "Linha 8",
            "Linha 9",
            "Linha 10"});
            this.lbErro.Location = new System.Drawing.Point(15, 19);
            this.lbErro.Name = "lbErro";
            this.lbErro.ScrollAlwaysVisible = true;
            this.lbErro.Size = new System.Drawing.Size(1263, 56);
            this.lbErro.TabIndex = 0;
            // 
            // gboxMensagensInformativas
            // 
            this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
            this.gboxMensagensInformativas.Location = new System.Drawing.Point(10, 415);
            this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
            this.gboxMensagensInformativas.Size = new System.Drawing.Size(1290, 88);
            this.gboxMensagensInformativas.TabIndex = 25;
            this.gboxMensagensInformativas.TabStop = false;
            this.gboxMensagensInformativas.Text = "Mensagens Informativas";
            // 
            // lbMensagem
            // 
            this.lbMensagem.FormattingEnabled = true;
            this.lbMensagem.Items.AddRange(new object[] {
            "Linha 1",
            "Linha 2",
            "Linha 3",
            "Linha 4",
            "Linha 5",
            "Linha 6",
            "Linha 7",
            "Linha 8",
            "Linha 9",
            "Linha 10"});
            this.lbMensagem.Location = new System.Drawing.Point(15, 19);
            this.lbMensagem.Name = "lbMensagem";
            this.lbMensagem.ScrollAlwaysVisible = true;
            this.lbMensagem.Size = new System.Drawing.Size(1263, 56);
            this.lbMensagem.TabIndex = 0;
            // 
            // lblTotalRegistros
            // 
            this.lblTotalRegistros.AutoSize = true;
            this.lblTotalRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotalRegistros.Location = new System.Drawing.Point(112, 383);
            this.lblTotalRegistros.Name = "lblTotalRegistros";
            this.lblTotalRegistros.Size = new System.Drawing.Size(14, 13);
            this.lblTotalRegistros.TabIndex = 3;
            this.lblTotalRegistros.Text = "0";
            // 
            // lblTitTotalRegistros
            // 
            this.lblTitTotalRegistros.AutoSize = true;
            this.lblTitTotalRegistros.Location = new System.Drawing.Point(10, 383);
            this.lblTitTotalRegistros.Name = "lblTitTotalRegistros";
            this.lblTitTotalRegistros.Size = new System.Drawing.Size(96, 13);
            this.lblTitTotalRegistros.TabIndex = 27;
            this.lblTitTotalRegistros.Text = "Total de Registros:";
            // 
            // grid
            // 
            this.grid.AllowUserToAddRows = false;
            this.grid.AllowUserToDeleteRows = false;
            this.grid.AllowUserToResizeRows = false;
            this.grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.grid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColVisibleOrdenacaoPadrao,
            this.NF,
            this.Destinatario,
            this.Destino,
            this.Situacao,
            this.Detalhe,
            this.DataEntrega,
            this.PrevisaoEntrega,
            this.Status,
            this.Mensagem,
            this.ColHiddenValorOrdenacaoPadrao,
            this.ColHiddenNF,
            this.ColHiddenDataEntrega,
            this.ColHiddenPrevisaoEntrega,
            this.ColHiddenGuid});
            dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grid.DefaultCellStyle = dataGridViewCellStyle12;
            this.grid.Location = new System.Drawing.Point(10, 85);
            this.grid.MultiSelect = false;
            this.grid.Name = "grid";
            this.grid.ReadOnly = true;
            this.grid.RowHeadersVisible = false;
            this.grid.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.grid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grid.ShowEditingIcon = false;
            this.grid.Size = new System.Drawing.Size(1290, 294);
            this.grid.TabIndex = 2;
            this.grid.SortCompare += new System.Windows.Forms.DataGridViewSortCompareEventHandler(this.grid_SortCompare);
            // 
            // ColVisibleOrdenacaoPadrao
            // 
            this.ColVisibleOrdenacaoPadrao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.ColVisibleOrdenacaoPadrao.DefaultCellStyle = dataGridViewCellStyle2;
            this.ColVisibleOrdenacaoPadrao.Frozen = true;
            this.ColVisibleOrdenacaoPadrao.HeaderText = "";
            this.ColVisibleOrdenacaoPadrao.MinimumWidth = 40;
            this.ColVisibleOrdenacaoPadrao.Name = "ColVisibleOrdenacaoPadrao";
            this.ColVisibleOrdenacaoPadrao.ReadOnly = true;
            this.ColVisibleOrdenacaoPadrao.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.ColVisibleOrdenacaoPadrao.Width = 40;
            // 
            // NF
            // 
            this.NF.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NF.DefaultCellStyle = dataGridViewCellStyle3;
            this.NF.Frozen = true;
            this.NF.HeaderText = "NF";
            this.NF.MinimumWidth = 70;
            this.NF.Name = "NF";
            this.NF.ReadOnly = true;
            this.NF.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.NF.Width = 70;
            // 
            // Destinatario
            // 
            this.Destinatario.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.Destinatario.DefaultCellStyle = dataGridViewCellStyle4;
            this.Destinatario.HeaderText = "Destinatário";
            this.Destinatario.Name = "Destinatario";
            this.Destinatario.ReadOnly = true;
            // 
            // Destino
            // 
            this.Destino.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.Destino.DefaultCellStyle = dataGridViewCellStyle5;
            this.Destino.HeaderText = "Destino";
            this.Destino.MinimumWidth = 150;
            this.Destino.Name = "Destino";
            this.Destino.ReadOnly = true;
            this.Destino.Width = 150;
            // 
            // Situacao
            // 
            this.Situacao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.Situacao.DefaultCellStyle = dataGridViewCellStyle6;
            this.Situacao.HeaderText = "Situação";
            this.Situacao.MinimumWidth = 140;
            this.Situacao.Name = "Situacao";
            this.Situacao.ReadOnly = true;
            this.Situacao.Width = 140;
            // 
            // Detalhe
            // 
            this.Detalhe.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.Detalhe.DefaultCellStyle = dataGridViewCellStyle7;
            this.Detalhe.HeaderText = "Detalhe";
            this.Detalhe.MinimumWidth = 220;
            this.Detalhe.Name = "Detalhe";
            this.Detalhe.ReadOnly = true;
            this.Detalhe.Width = 220;
            // 
            // DataEntrega
            // 
            this.DataEntrega.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DataEntrega.DefaultCellStyle = dataGridViewCellStyle8;
            this.DataEntrega.HeaderText = "Data Entrega";
            this.DataEntrega.MinimumWidth = 85;
            this.DataEntrega.Name = "DataEntrega";
            this.DataEntrega.ReadOnly = true;
            this.DataEntrega.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.DataEntrega.Width = 85;
            // 
            // PrevisaoEntrega
            // 
            this.PrevisaoEntrega.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PrevisaoEntrega.DefaultCellStyle = dataGridViewCellStyle9;
            this.PrevisaoEntrega.HeaderText = "Previsão Entrega";
            this.PrevisaoEntrega.MinimumWidth = 85;
            this.PrevisaoEntrega.Name = "PrevisaoEntrega";
            this.PrevisaoEntrega.ReadOnly = true;
            this.PrevisaoEntrega.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.PrevisaoEntrega.Width = 85;
            // 
            // Status
            // 
            this.Status.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.Status.DefaultCellStyle = dataGridViewCellStyle10;
            this.Status.HeaderText = "Status";
            this.Status.MinimumWidth = 120;
            this.Status.Name = "Status";
            this.Status.ReadOnly = true;
            this.Status.Width = 120;
            // 
            // Mensagem
            // 
            this.Mensagem.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.Mensagem.DefaultCellStyle = dataGridViewCellStyle11;
            this.Mensagem.HeaderText = "Mensagem";
            this.Mensagem.MinimumWidth = 170;
            this.Mensagem.Name = "Mensagem";
            this.Mensagem.ReadOnly = true;
            this.Mensagem.Width = 170;
            // 
            // ColHiddenValorOrdenacaoPadrao
            // 
            this.ColHiddenValorOrdenacaoPadrao.HeaderText = "Campo Ordenação Padrão";
            this.ColHiddenValorOrdenacaoPadrao.Name = "ColHiddenValorOrdenacaoPadrao";
            this.ColHiddenValorOrdenacaoPadrao.ReadOnly = true;
            this.ColHiddenValorOrdenacaoPadrao.Visible = false;
            // 
            // ColHiddenNF
            // 
            this.ColHiddenNF.HeaderText = "Campo Ordenação NF";
            this.ColHiddenNF.Name = "ColHiddenNF";
            this.ColHiddenNF.ReadOnly = true;
            this.ColHiddenNF.Visible = false;
            // 
            // ColHiddenDataEntrega
            // 
            this.ColHiddenDataEntrega.HeaderText = "Campo Ordenação Data Entrega";
            this.ColHiddenDataEntrega.Name = "ColHiddenDataEntrega";
            this.ColHiddenDataEntrega.ReadOnly = true;
            this.ColHiddenDataEntrega.Visible = false;
            // 
            // ColHiddenPrevisaoEntrega
            // 
            this.ColHiddenPrevisaoEntrega.HeaderText = "Campo Ordenação Previsão Entrega";
            this.ColHiddenPrevisaoEntrega.Name = "ColHiddenPrevisaoEntrega";
            this.ColHiddenPrevisaoEntrega.ReadOnly = true;
            this.ColHiddenPrevisaoEntrega.Visible = false;
            // 
            // ColHiddenGuid
            // 
            this.ColHiddenGuid.HeaderText = "Guid";
            this.ColHiddenGuid.Name = "ColHiddenGuid";
            this.ColHiddenGuid.ReadOnly = true;
            this.ColHiddenGuid.Visible = false;
            // 
            // lblQtdeRegErro
            // 
            this.lblQtdeRegErro.AutoSize = true;
            this.lblQtdeRegErro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQtdeRegErro.ForeColor = System.Drawing.Color.Red;
            this.lblQtdeRegErro.Location = new System.Drawing.Point(380, 383);
            this.lblQtdeRegErro.Name = "lblQtdeRegErro";
            this.lblQtdeRegErro.Size = new System.Drawing.Size(14, 13);
            this.lblQtdeRegErro.TabIndex = 4;
            this.lblQtdeRegErro.Text = "0";
            // 
            // lblTitQtdeRegErro
            // 
            this.lblTitQtdeRegErro.AutoSize = true;
            this.lblTitQtdeRegErro.Location = new System.Drawing.Point(249, 383);
            this.lblTitQtdeRegErro.Name = "lblTitQtdeRegErro";
            this.lblTitQtdeRegErro.Size = new System.Drawing.Size(125, 13);
            this.lblTitQtdeRegErro.TabIndex = 30;
            this.lblTitQtdeRegErro.Text = "Qtde Registros com Erro:";
            // 
            // lblQtdeRegApto
            // 
            this.lblQtdeRegApto.AutoSize = true;
            this.lblQtdeRegApto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQtdeRegApto.ForeColor = System.Drawing.Color.Green;
            this.lblQtdeRegApto.Location = new System.Drawing.Point(637, 383);
            this.lblQtdeRegApto.Name = "lblQtdeRegApto";
            this.lblQtdeRegApto.Size = new System.Drawing.Size(14, 13);
            this.lblQtdeRegApto.TabIndex = 5;
            this.lblQtdeRegApto.Text = "0";
            // 
            // lblTitQtdeRegApto
            // 
            this.lblTitQtdeRegApto.AutoSize = true;
            this.lblTitQtdeRegApto.Location = new System.Drawing.Point(521, 383);
            this.lblTitQtdeRegApto.Name = "lblTitQtdeRegApto";
            this.lblTitQtdeRegApto.Size = new System.Drawing.Size(110, 13);
            this.lblTitQtdeRegApto.TabIndex = 32;
            this.lblTitQtdeRegApto.Text = "Qtde Registros Aptos:";
            // 
            // lblQtdeAtualizSucesso
            // 
            this.lblQtdeAtualizSucesso.AutoSize = true;
            this.lblQtdeAtualizSucesso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQtdeAtualizSucesso.ForeColor = System.Drawing.Color.Green;
            this.lblQtdeAtualizSucesso.Location = new System.Drawing.Point(1246, 383);
            this.lblQtdeAtualizSucesso.Name = "lblQtdeAtualizSucesso";
            this.lblQtdeAtualizSucesso.Size = new System.Drawing.Size(14, 13);
            this.lblQtdeAtualizSucesso.TabIndex = 7;
            this.lblQtdeAtualizSucesso.Text = "0";
            // 
            // lblTitQtdeAtualizSucesso
            // 
            this.lblTitQtdeAtualizSucesso.AutoSize = true;
            this.lblTitQtdeAtualizSucesso.Location = new System.Drawing.Point(1077, 383);
            this.lblTitQtdeAtualizSucesso.Name = "lblTitQtdeAtualizSucesso";
            this.lblTitQtdeAtualizSucesso.Size = new System.Drawing.Size(163, 13);
            this.lblTitQtdeAtualizSucesso.TabIndex = 36;
            this.lblTitQtdeAtualizSucesso.Text = "Qtde Atualizações com Sucesso:";
            // 
            // lblQtdeAtualizFalha
            // 
            this.lblQtdeAtualizFalha.AutoSize = true;
            this.lblQtdeAtualizFalha.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQtdeAtualizFalha.ForeColor = System.Drawing.Color.Red;
            this.lblQtdeAtualizFalha.Location = new System.Drawing.Point(943, 383);
            this.lblQtdeAtualizFalha.Name = "lblQtdeAtualizFalha";
            this.lblQtdeAtualizFalha.Size = new System.Drawing.Size(14, 13);
            this.lblQtdeAtualizFalha.TabIndex = 6;
            this.lblQtdeAtualizFalha.Text = "0";
            // 
            // lblTitQtdeAtualizFalha
            // 
            this.lblTitQtdeAtualizFalha.AutoSize = true;
            this.lblTitQtdeAtualizFalha.Location = new System.Drawing.Point(789, 383);
            this.lblTitQtdeAtualizFalha.Name = "lblTitQtdeAtualizFalha";
            this.lblTitQtdeAtualizFalha.Size = new System.Drawing.Size(148, 13);
            this.lblTitQtdeAtualizFalha.TabIndex = 34;
            this.lblTitQtdeAtualizFalha.Text = "Qtde Atualizações com Falha:";
            // 
            // FAnotarPedidoRecebidoCliente
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1314, 706);
            this.Name = "FAnotarPedidoRecebidoCliente";
            this.Text = "FAnotarPedidoRecebidoCliente";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FAnotarPedidoRecebidoCliente_FormClosing);
            this.Load += new System.EventHandler(this.FAnotarPedidoRecebidoCliente_Load);
            this.Shown += new System.EventHandler(this.FAnotarPedidoRecebidoCliente_Shown);
            this.pnBotoes.ResumeLayout(false);
            this.pnCampos.ResumeLayout(false);
            this.pnCampos.PerformLayout();
            this.gboxMsgErro.ResumeLayout(false);
            this.gboxMensagensInformativas.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelecionaArquivoRastreio;
        private System.Windows.Forms.TextBox txtArquivoRastreio;
        private System.Windows.Forms.Label lblArquivoRastreio;
        private System.Windows.Forms.Label lblTituloPainel;
        private System.Windows.Forms.OpenFileDialog openFileDialogCtrl;
        private System.Windows.Forms.Button btnConfirma;
        private System.Windows.Forms.GroupBox gboxMsgErro;
        private System.Windows.Forms.ListBox lbErro;
        private System.Windows.Forms.GroupBox gboxMensagensInformativas;
        private System.Windows.Forms.ListBox lbMensagem;
        private System.Windows.Forms.Label lblTotalRegistros;
        private System.Windows.Forms.Label lblTitTotalRegistros;
        private System.Windows.Forms.DataGridView grid;
        private System.Windows.Forms.Label lblQtdeRegApto;
        private System.Windows.Forms.Label lblTitQtdeRegApto;
        private System.Windows.Forms.Label lblQtdeRegErro;
        private System.Windows.Forms.Label lblTitQtdeRegErro;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColVisibleOrdenacaoPadrao;
        private System.Windows.Forms.DataGridViewTextBoxColumn NF;
        private System.Windows.Forms.DataGridViewTextBoxColumn Destinatario;
        private System.Windows.Forms.DataGridViewTextBoxColumn Destino;
        private System.Windows.Forms.DataGridViewTextBoxColumn Situacao;
        private System.Windows.Forms.DataGridViewTextBoxColumn Detalhe;
        private System.Windows.Forms.DataGridViewTextBoxColumn DataEntrega;
        private System.Windows.Forms.DataGridViewTextBoxColumn PrevisaoEntrega;
        private System.Windows.Forms.DataGridViewTextBoxColumn Status;
        private System.Windows.Forms.DataGridViewTextBoxColumn Mensagem;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColHiddenValorOrdenacaoPadrao;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColHiddenNF;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColHiddenDataEntrega;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColHiddenPrevisaoEntrega;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColHiddenGuid;
        private System.Windows.Forms.Label lblQtdeAtualizSucesso;
        private System.Windows.Forms.Label lblTitQtdeAtualizSucesso;
        private System.Windows.Forms.Label lblQtdeAtualizFalha;
        private System.Windows.Forms.Label lblTitQtdeAtualizFalha;
    }
}