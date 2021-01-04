namespace ADM2
{
    partial class FImportXML
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FImportXML));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle16 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblTitulo = new System.Windows.Forms.Label();
            this.gboxMsgErro = new System.Windows.Forms.GroupBox();
            this.lbErro = new System.Windows.Forms.ListBox();
            this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
            this.lbMensagem = new System.Windows.Forms.ListBox();
            this.btnAtualizaDatas = new System.Windows.Forms.Button();
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
            this.LblPeriodoEntradaNFde = new System.Windows.Forms.Label();
            this.dtpDataEntradaIni = new System.Windows.Forms.DateTimePicker();
            this.dtpDataEntradaFim = new System.Windows.Forms.DateTimePicker();
            this.LblPeriodoEntradaNFate = new System.Windows.Forms.Label();
            this.pnBotoes.SuspendLayout();
            this.pnCampos.SuspendLayout();
            this.gboxMsgErro.SuspendLayout();
            this.gboxMensagensInformativas.SuspendLayout();
            this.gboxDados.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdDados)).BeginInit();
            this.SuspendLayout();
            // 
            // pnBotoes
            // 
            this.pnBotoes.Controls.Add(this.btnAtualizaDatas);
            this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnAtualizaDatas, 0);
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
            // 
            // pnCampos
            // 
            this.pnCampos.Controls.Add(this.dtpDataEntradaFim);
            this.pnCampos.Controls.Add(this.LblPeriodoEntradaNFate);
            this.pnCampos.Controls.Add(this.dtpDataEntradaIni);
            this.pnCampos.Controls.Add(this.LblPeriodoEntradaNFde);
            this.pnCampos.Controls.Add(this.gboxDados);
            this.pnCampos.Controls.Add(this.gboxMsgErro);
            this.pnCampos.Controls.Add(this.gboxMensagensInformativas);
            this.pnCampos.Controls.Add(this.lblTitulo);
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
            this.lblTitulo.TabIndex = 1;
            this.lblTitulo.Text = "Atualizar Datas de Importação de Arquivos XML";
            this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // gboxMsgErro
            // 
            this.gboxMsgErro.Controls.Add(this.lbErro);
            this.gboxMsgErro.Location = new System.Drawing.Point(3, 362);
            this.gboxMsgErro.Name = "gboxMsgErro";
            this.gboxMsgErro.Size = new System.Drawing.Size(994, 95);
            this.gboxMsgErro.TabIndex = 9;
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
            // gboxMensagensInformativas
            // 
            this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
            this.gboxMensagensInformativas.Location = new System.Drawing.Point(3, 261);
            this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
            this.gboxMensagensInformativas.Size = new System.Drawing.Size(994, 95);
            this.gboxMensagensInformativas.TabIndex = 8;
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
            // btnAtualizaDatas
            // 
            this.btnAtualizaDatas.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnAtualizaDatas.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizaDatas.Image")));
            this.btnAtualizaDatas.Location = new System.Drawing.Point(869, 4);
            this.btnAtualizaDatas.Name = "btnAtualizaDatas";
            this.btnAtualizaDatas.Size = new System.Drawing.Size(40, 44);
            this.btnAtualizaDatas.TabIndex = 8;
            this.btnAtualizaDatas.TabStop = false;
            this.btnAtualizaDatas.UseVisualStyleBackColor = true;
            // 
            // gboxDados
            // 
            this.gboxDados.Controls.Add(this.lblTotalRegistros);
            this.gboxDados.Controls.Add(this.lblTitTotalRegistros);
            this.gboxDados.Controls.Add(this.grdDados);
            this.gboxDados.Location = new System.Drawing.Point(7, 82);
            this.gboxDados.Name = "gboxDados";
            this.gboxDados.Size = new System.Drawing.Size(994, 173);
            this.gboxDados.TabIndex = 10;
            this.gboxDados.TabStop = false;
            this.gboxDados.Text = "Dados do Arquivo";
            // 
            // lblTotalRegistros
            // 
            this.lblTotalRegistros.AutoSize = true;
            this.lblTotalRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotalRegistros.Location = new System.Drawing.Point(108, 154);
            this.lblTotalRegistros.Name = "lblTotalRegistros";
            this.lblTotalRegistros.Size = new System.Drawing.Size(28, 13);
            this.lblTotalRegistros.TabIndex = 2;
            this.lblTotalRegistros.Text = "999";
            // 
            // lblTitTotalRegistros
            // 
            this.lblTitTotalRegistros.AutoSize = true;
            this.lblTitTotalRegistros.Location = new System.Drawing.Point(20, 154);
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
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.grdDados.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle9;
            this.grdDados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdDados.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colCodigo,
            this.colEX,
            this.colTabela,
            this.colAliqNac,
            this.colAliqImp,
            this.colDescricao});
            dataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.grdDados.DefaultCellStyle = dataGridViewCellStyle16;
            this.grdDados.Location = new System.Drawing.Point(15, 19);
            this.grdDados.MultiSelect = false;
            this.grdDados.Name = "grdDados";
            this.grdDados.ReadOnly = true;
            this.grdDados.RowHeadersVisible = false;
            this.grdDados.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.grdDados.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.grdDados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grdDados.ShowEditingIcon = false;
            this.grdDados.Size = new System.Drawing.Size(965, 129);
            this.grdDados.StandardTab = true;
            this.grdDados.TabIndex = 0;
            // 
            // colCodigo
            // 
            this.colCodigo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            this.colCodigo.DefaultCellStyle = dataGridViewCellStyle10;
            this.colCodigo.HeaderText = "Código";
            this.colCodigo.MinimumWidth = 100;
            this.colCodigo.Name = "colCodigo";
            this.colCodigo.ReadOnly = true;
            // 
            // colEX
            // 
            this.colEX.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.colEX.DefaultCellStyle = dataGridViewCellStyle11;
            this.colEX.HeaderText = "EX";
            this.colEX.MinimumWidth = 60;
            this.colEX.Name = "colEX";
            this.colEX.ReadOnly = true;
            this.colEX.Width = 60;
            // 
            // colTabela
            // 
            this.colTabela.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.colTabela.DefaultCellStyle = dataGridViewCellStyle12;
            this.colTabela.HeaderText = "Tabela";
            this.colTabela.MinimumWidth = 80;
            this.colTabela.Name = "colTabela";
            this.colTabela.ReadOnly = true;
            this.colTabela.Width = 80;
            // 
            // colAliqNac
            // 
            this.colAliqNac.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.colAliqNac.DefaultCellStyle = dataGridViewCellStyle13;
            this.colAliqNac.HeaderText = "Alíquota Nac";
            this.colAliqNac.MinimumWidth = 110;
            this.colAliqNac.Name = "colAliqNac";
            this.colAliqNac.ReadOnly = true;
            this.colAliqNac.Width = 110;
            // 
            // colAliqImp
            // 
            this.colAliqImp.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.colAliqImp.DefaultCellStyle = dataGridViewCellStyle14;
            this.colAliqImp.HeaderText = "Alíquota Imp";
            this.colAliqImp.MinimumWidth = 110;
            this.colAliqImp.Name = "colAliqImp";
            this.colAliqImp.ReadOnly = true;
            this.colAliqImp.Width = 110;
            // 
            // colDescricao
            // 
            this.colDescricao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            dataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.colDescricao.DefaultCellStyle = dataGridViewCellStyle15;
            this.colDescricao.HeaderText = "Descrição";
            this.colDescricao.MinimumWidth = 150;
            this.colDescricao.Name = "colDescricao";
            this.colDescricao.ReadOnly = true;
            // 
            // LblPeriodoEntradaNFde
            // 
            this.LblPeriodoEntradaNFde.AutoSize = true;
            this.LblPeriodoEntradaNFde.Location = new System.Drawing.Point(15, 51);
            this.LblPeriodoEntradaNFde.Name = "LblPeriodoEntradaNFde";
            this.LblPeriodoEntradaNFde.Size = new System.Drawing.Size(166, 13);
            this.LblPeriodoEntradaNFde.TabIndex = 11;
            this.LblPeriodoEntradaNFde.Text = "Período de Entrada das NFs:  de ";
            // 
            // dtpDataEntradaIni
            // 
            this.dtpDataEntradaIni.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataEntradaIni.Location = new System.Drawing.Point(187, 51);
            this.dtpDataEntradaIni.Name = "dtpDataEntradaIni";
            this.dtpDataEntradaIni.Size = new System.Drawing.Size(101, 20);
            this.dtpDataEntradaIni.TabIndex = 12;
            // 
            // dtpDataEntradaFim
            // 
            this.dtpDataEntradaFim.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataEntradaFim.Location = new System.Drawing.Point(328, 51);
            this.dtpDataEntradaFim.Name = "dtpDataEntradaFim";
            this.dtpDataEntradaFim.Size = new System.Drawing.Size(101, 20);
            this.dtpDataEntradaFim.TabIndex = 14;
            // 
            // LblPeriodoEntradaNFate
            // 
            this.LblPeriodoEntradaNFate.AutoSize = true;
            this.LblPeriodoEntradaNFate.Location = new System.Drawing.Point(294, 51);
            this.LblPeriodoEntradaNFate.Name = "LblPeriodoEntradaNFate";
            this.LblPeriodoEntradaNFate.Size = new System.Drawing.Size(28, 13);
            this.LblPeriodoEntradaNFate.TabIndex = 13;
            this.LblPeriodoEntradaNFate.Text = " até ";
            // 
            // FImportXML
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.ClientSize = new System.Drawing.Size(1008, 562);
            this.Name = "FImportXML";
            this.Text = "ADM2  -  1.11 - 05.JAN.2021";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FImportXML_FormClosing);
            this.Load += new System.EventHandler(this.FImportXML_Load);
            this.Shown += new System.EventHandler(this.FImportXML_Shown);
            this.pnBotoes.ResumeLayout(false);
            this.pnCampos.ResumeLayout(false);
            this.pnCampos.PerformLayout();
            this.gboxMsgErro.ResumeLayout(false);
            this.gboxMensagensInformativas.ResumeLayout(false);
            this.gboxDados.ResumeLayout(false);
            this.gboxDados.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdDados)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.GroupBox gboxMsgErro;
        private System.Windows.Forms.ListBox lbErro;
        private System.Windows.Forms.GroupBox gboxMensagensInformativas;
        private System.Windows.Forms.ListBox lbMensagem;
        private System.Windows.Forms.Button btnAtualizaDatas;
        private System.Windows.Forms.GroupBox gboxDados;
        private System.Windows.Forms.Label lblTotalRegistros;
        private System.Windows.Forms.Label lblTitTotalRegistros;
        private System.Windows.Forms.DataGridView grdDados;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCodigo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colEX;
        private System.Windows.Forms.DataGridViewTextBoxColumn colTabela;
        private System.Windows.Forms.DataGridViewTextBoxColumn colAliqNac;
        private System.Windows.Forms.DataGridViewTextBoxColumn colAliqImp;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDescricao;
        private System.Windows.Forms.Label LblPeriodoEntradaNFde;
        private System.Windows.Forms.DateTimePicker dtpDataEntradaFim;
        private System.Windows.Forms.Label LblPeriodoEntradaNFate;
        private System.Windows.Forms.DateTimePicker dtpDataEntradaIni;
    }
}
