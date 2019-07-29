namespace ConsolidadorXlsEC
{
    partial class FConfirmaPedidoStatus
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FConfirmaPedidoStatus));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle34 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle44 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle35 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle36 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle37 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle38 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle39 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle40 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle41 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle42 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle43 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle98 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle99 = new System.Windows.Forms.DataGridViewCellStyle();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lblMensagemAlertaPt1 = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.grdPedidosConfirma = new System.Windows.Forms.DataGridView();
            this.colGrdDadosCheckBox = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colGrdDadosCheckBoxConfirma = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colGrdDadosTransportadora = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosNumPedido = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosNumMagento = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosNumMktplace = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosMarketplaceDescricao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosMarketplace = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosCodigoOrigemPai = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosCliente = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosValor = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosRecebido = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosStatusDescricao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrdDadosMensagemStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblMensagemAlertaPt2 = new System.Windows.Forms.Label();
            this.btnMarcarTodos = new System.Windows.Forms.Button();
            this.btnDesmarcarTodos = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdPedidosConfirma)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(12, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(83, 83);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            // 
            // lblMensagemAlertaPt1
            // 
            this.lblMensagemAlertaPt1.AutoSize = true;
            this.lblMensagemAlertaPt1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMensagemAlertaPt1.Location = new System.Drawing.Point(101, 55);
            this.lblMensagemAlertaPt1.Name = "lblMensagemAlertaPt1";
            this.lblMensagemAlertaPt1.Size = new System.Drawing.Size(769, 20);
            this.lblMensagemAlertaPt1.TabIndex = 9;
            this.lblMensagemAlertaPt1.Text = "Os pedidos abaixo contém status não previstos para o tratamento automático no Mag" +
    "ento.";
            // 
            // btnOk
            // 
            this.btnOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.Image = ((System.Drawing.Image)(resources.GetObject("btnOk.Image")));
            this.btnOk.Location = new System.Drawing.Point(835, 369);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(170, 40);
            this.btnOk.TabIndex = 11;
            this.btnOk.Text = "Prosseguir";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // grdPedidosConfirma
            // 
            this.grdPedidosConfirma.AllowUserToAddRows = false;
            this.grdPedidosConfirma.AllowUserToDeleteRows = false;
            this.grdPedidosConfirma.AllowUserToResizeColumns = false;
            this.grdPedidosConfirma.AllowUserToResizeRows = false;
            this.grdPedidosConfirma.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            dataGridViewCellStyle34.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter;
            dataGridViewCellStyle34.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle34.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle34.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle34.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle34.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle34.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grdPedidosConfirma.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle34;
            this.grdPedidosConfirma.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.grdPedidosConfirma.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colGrdDadosCheckBox,
            this.colGrdDadosCheckBoxConfirma,
            this.colGrdDadosTransportadora,
            this.colGrdDadosNumPedido,
            this.colGrdDadosNumMagento,
            this.colGrdDadosNumMktplace,
            this.colGrdDadosMarketplaceDescricao,
            this.colGrdDadosMarketplace,
            this.colGrdDadosCodigoOrigemPai,
            this.colGrdDadosCliente,
            this.colGrdDadosValor,
            this.colGrdDadosRecebido,
            this.colGrdDadosStatus,
            this.colGrdDadosStatusDescricao,
            this.colGrdDadosMensagemStatus});
            this.grdPedidosConfirma.Location = new System.Drawing.Point(12, 101);
            this.grdPedidosConfirma.MultiSelect = false;
            this.grdPedidosConfirma.Name = "grdPedidosConfirma";
            this.grdPedidosConfirma.RowHeadersVisible = false;
            this.grdPedidosConfirma.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridViewCellStyle44.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            dataGridViewCellStyle44.SelectionForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle44.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grdPedidosConfirma.RowsDefaultCellStyle = dataGridViewCellStyle44;
            this.grdPedidosConfirma.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.grdPedidosConfirma.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grdPedidosConfirma.ShowEditingIcon = false;
            this.grdPedidosConfirma.Size = new System.Drawing.Size(993, 253);
            this.grdPedidosConfirma.StandardTab = true;
            this.grdPedidosConfirma.TabIndex = 12;
            // 
            // colGrdDadosCheckBox
            // 
            this.colGrdDadosCheckBox.HeaderText = "";
            this.colGrdDadosCheckBox.MinimumWidth = 30;
            this.colGrdDadosCheckBox.Name = "colGrdDadosCheckBox";
            this.colGrdDadosCheckBox.Visible = false;
            this.colGrdDadosCheckBox.Width = 30;
            // 
            // colGrdDadosCheckBoxConfirma
            // 
            this.colGrdDadosCheckBoxConfirma.HeaderText = "";
            this.colGrdDadosCheckBoxConfirma.MinimumWidth = 30;
            this.colGrdDadosCheckBoxConfirma.Name = "colGrdDadosCheckBoxConfirma";
            this.colGrdDadosCheckBoxConfirma.Width = 30;
            // 
            // colGrdDadosTransportadora
            // 
            this.colGrdDadosTransportadora.DataPropertyName = "transportadora_id";
            this.colGrdDadosTransportadora.HeaderText = "Transportadora";
            this.colGrdDadosTransportadora.MinimumWidth = 100;
            this.colGrdDadosTransportadora.Name = "colGrdDadosTransportadora";
            // 
            // colGrdDadosNumPedido
            // 
            this.colGrdDadosNumPedido.DataPropertyName = "pedido";
            dataGridViewCellStyle35.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            this.colGrdDadosNumPedido.DefaultCellStyle = dataGridViewCellStyle35;
            this.colGrdDadosNumPedido.HeaderText = "Nº Pedido";
            this.colGrdDadosNumPedido.MinimumWidth = 90;
            this.colGrdDadosNumPedido.Name = "colGrdDadosNumPedido";
            this.colGrdDadosNumPedido.ReadOnly = true;
            this.colGrdDadosNumPedido.Width = 90;
            // 
            // colGrdDadosNumMagento
            // 
            this.colGrdDadosNumMagento.DataPropertyName = "pedido_bs_x_ac";
            dataGridViewCellStyle36.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            this.colGrdDadosNumMagento.DefaultCellStyle = dataGridViewCellStyle36;
            this.colGrdDadosNumMagento.HeaderText = "Nº Magento";
            this.colGrdDadosNumMagento.MinimumWidth = 100;
            this.colGrdDadosNumMagento.Name = "colGrdDadosNumMagento";
            this.colGrdDadosNumMagento.ReadOnly = true;
            // 
            // colGrdDadosNumMktplace
            // 
            this.colGrdDadosNumMktplace.DataPropertyName = "pedido_bs_x_marketplace";
            dataGridViewCellStyle37.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            this.colGrdDadosNumMktplace.DefaultCellStyle = dataGridViewCellStyle37;
            this.colGrdDadosNumMktplace.HeaderText = "Nº Marketplace";
            this.colGrdDadosNumMktplace.MinimumWidth = 110;
            this.colGrdDadosNumMktplace.Name = "colGrdDadosNumMktplace";
            this.colGrdDadosNumMktplace.ReadOnly = true;
            this.colGrdDadosNumMktplace.Width = 110;
            // 
            // colGrdDadosMarketplaceDescricao
            // 
            this.colGrdDadosMarketplaceDescricao.DataPropertyName = "marketplace_codigo_origem_descricao";
            dataGridViewCellStyle98.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            this.colGrdDadosMarketplaceDescricao.DefaultCellStyle = dataGridViewCellStyle98;
            this.colGrdDadosMarketplaceDescricao.HeaderText = "Marketplace";
            this.colGrdDadosMarketplaceDescricao.MinimumWidth = 120;
            this.colGrdDadosMarketplaceDescricao.Name = "colGrdDadosMarketplaceDescricao";
            this.colGrdDadosMarketplaceDescricao.ReadOnly = true;
            this.colGrdDadosMarketplaceDescricao.Width = 120;
            // 
            // colGrdDadosMarketplace
            // 
            this.colGrdDadosMarketplace.DataPropertyName = "marketplace_codigo_origem";
            this.colGrdDadosMarketplace.HeaderText = "Marketplace Código";
            this.colGrdDadosMarketplace.Name = "colGrdDadosMarketplace";
            this.colGrdDadosMarketplace.Visible = false;
            // 
            // colGrdDadosCodigoOrigemPai
            // 
            this.colGrdDadosCodigoOrigemPai.DataPropertyName = "marketplace_codigo_origem_pai";
            this.colGrdDadosCodigoOrigemPai.HeaderText = "Marketplace Codigo Pai";
            this.colGrdDadosCodigoOrigemPai.Name = "colGrdDadosCodigoOrigemPai";
            this.colGrdDadosCodigoOrigemPai.Visible = false;
            // 
            // colGrdDadosCliente
            // 
            this.colGrdDadosCliente.DataPropertyName = "nome_iniciais_em_maiusculas";
            dataGridViewCellStyle39.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            this.colGrdDadosCliente.DefaultCellStyle = dataGridViewCellStyle39;
            this.colGrdDadosCliente.HeaderText = "Cliente";
            this.colGrdDadosCliente.MinimumWidth = 240;
            this.colGrdDadosCliente.Name = "colGrdDadosCliente";
            this.colGrdDadosCliente.ReadOnly = true;
            this.colGrdDadosCliente.Width = 240;
            // 
            // colGrdDadosValor
            // 
            this.colGrdDadosValor.DataPropertyName = "vl_pedido";
            dataGridViewCellStyle40.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.colGrdDadosValor.DefaultCellStyle = dataGridViewCellStyle40;
            this.colGrdDadosValor.HeaderText = "Valor";
            this.colGrdDadosValor.MinimumWidth = 80;
            this.colGrdDadosValor.Name = "colGrdDadosValor";
            this.colGrdDadosValor.ReadOnly = true;
            this.colGrdDadosValor.Width = 80;
            // 
            // colGrdDadosRecebido
            // 
            this.colGrdDadosRecebido.DataPropertyName = "MarketplacePedidoRecebidoRegistrarDataRecebido";
            dataGridViewCellStyle41.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.colGrdDadosRecebido.DefaultCellStyle = dataGridViewCellStyle41;
            this.colGrdDadosRecebido.HeaderText = "Recebido";
            this.colGrdDadosRecebido.MinimumWidth = 110;
            this.colGrdDadosRecebido.Name = "colGrdDadosRecebido";
            this.colGrdDadosRecebido.ReadOnly = true;
            this.colGrdDadosRecebido.Visible = false;
            this.colGrdDadosRecebido.Width = 110;
            // 
            // colGrdDadosStatus
            // 
            dataGridViewCellStyle99.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.colGrdDadosStatus.DefaultCellStyle = dataGridViewCellStyle99;
            this.colGrdDadosStatus.HeaderText = "Status";
            this.colGrdDadosStatus.MinimumWidth = 120;
            this.colGrdDadosStatus.Name = "colGrdDadosStatus";
            this.colGrdDadosStatus.ReadOnly = true;
            this.colGrdDadosStatus.Visible = false;
            this.colGrdDadosStatus.Width = 120;
            // 
            // colGrdDadosStatusDescricao
            // 
            this.colGrdDadosStatusDescricao.HeaderText = "Status";
            this.colGrdDadosStatusDescricao.MinimumWidth = 130;
            this.colGrdDadosStatusDescricao.Name = "colGrdDadosStatusDescricao";
            this.colGrdDadosStatusDescricao.ReadOnly = true;
            this.colGrdDadosStatusDescricao.Width = 130;
            // 
            // colGrdDadosMensagemStatus
            // 
            dataGridViewCellStyle43.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            this.colGrdDadosMensagemStatus.DefaultCellStyle = dataGridViewCellStyle43;
            this.colGrdDadosMensagemStatus.HeaderText = "Mensagem";
            this.colGrdDadosMensagemStatus.MinimumWidth = 196;
            this.colGrdDadosMensagemStatus.Name = "colGrdDadosMensagemStatus";
            this.colGrdDadosMensagemStatus.ReadOnly = true;
            this.colGrdDadosMensagemStatus.Visible = false;
            this.colGrdDadosMensagemStatus.Width = 196;
            // 
            // lblMensagemAlertaPt2
            // 
            this.lblMensagemAlertaPt2.AutoSize = true;
            this.lblMensagemAlertaPt2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMensagemAlertaPt2.Location = new System.Drawing.Point(101, 75);
            this.lblMensagemAlertaPt2.Name = "lblMensagemAlertaPt2";
            this.lblMensagemAlertaPt2.Size = new System.Drawing.Size(713, 20);
            this.lblMensagemAlertaPt2.TabIndex = 13;
            this.lblMensagemAlertaPt2.Text = "Selecione os pedidos que deseja dar prosseguimento na baixa do sistema somente.";
            // 
            // btnMarcarTodos
            // 
            this.btnMarcarTodos.Location = new System.Drawing.Point(11, 378);
            this.btnMarcarTodos.Name = "btnMarcarTodos";
            this.btnMarcarTodos.Size = new System.Drawing.Size(114, 23);
            this.btnMarcarTodos.TabIndex = 15;
            this.btnMarcarTodos.Text = "Marcar todos";
            this.btnMarcarTodos.UseVisualStyleBackColor = true;
            this.btnMarcarTodos.Click += new System.EventHandler(this.btnMarcarTodos_Click);
            // 
            // btnDesmarcarTodos
            // 
            this.btnDesmarcarTodos.Location = new System.Drawing.Point(131, 378);
            this.btnDesmarcarTodos.Name = "btnDesmarcarTodos";
            this.btnDesmarcarTodos.Size = new System.Drawing.Size(114, 23);
            this.btnDesmarcarTodos.TabIndex = 14;
            this.btnDesmarcarTodos.Text = "Desmarcar todos";
            this.btnDesmarcarTodos.UseVisualStyleBackColor = true;
            this.btnDesmarcarTodos.Click += new System.EventHandler(this.btnDesmarcarTodos_Click);
            // 
            // FConfirmaPedidoStatus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1017, 421);
            this.Controls.Add(this.btnMarcarTodos);
            this.Controls.Add(this.btnDesmarcarTodos);
            this.Controls.Add(this.lblMensagemAlertaPt2);
            this.Controls.Add(this.grdPedidosConfirma);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lblMensagemAlertaPt1);
            this.Controls.Add(this.pictureBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FConfirmaPedidoStatus";
            this.Text = "Confirmar Pedidos";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FConfirmaPedidoStatus_Closing);
            this.Shown += new System.EventHandler(this.FIntegracaoMarketplace_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdPedidosConfirma)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lblMensagemAlertaPt1;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.DataGridView grdPedidosConfirma;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colGrdDadosCheckBox;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colGrdDadosCheckBoxConfirma;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosTransportadora;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosNumPedido;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosNumMagento;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosNumMktplace;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosMarketplaceDescricao;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosMarketplace;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosCodigoOrigemPai;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosCliente;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosValor;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosRecebido;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosStatusDescricao;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrdDadosMensagemStatus;
        private System.Windows.Forms.Label lblMensagemAlertaPt2;
        private System.Windows.Forms.Button btnMarcarTodos;
        private System.Windows.Forms.Button btnDesmarcarTodos;
    }
}