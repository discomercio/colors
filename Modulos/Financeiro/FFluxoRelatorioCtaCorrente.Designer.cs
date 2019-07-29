namespace Financeiro
{
	partial class FFluxoRelatorioCtaCorrente
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FFluxoRelatorioCtaCorrente));
            this.lblTitulo = new System.Windows.Forms.Label();
            this.pnParametros = new System.Windows.Forms.Panel();
            this.lbContaCorrente = new System.Windows.Forms.ListBox();
            this.lblTitContaCorrente = new System.Windows.Forms.Label();
            this.txtDataCompetenciaFinal = new System.Windows.Forms.TextBox();
            this.lblTitPeriodoCompetencia = new System.Windows.Forms.Label();
            this.txtDataCompetenciaInicial = new System.Windows.Forms.TextBox();
            this.lblDataCompetenciaAte = new System.Windows.Forms.Label();
            this.btnImprimir = new System.Windows.Forms.Button();
            this.btnLimpar = new System.Windows.Forms.Button();
            this.prnPreviewConsulta = new System.Windows.Forms.PrintPreviewDialog();
            this.prnDocConsulta = new System.Drawing.Printing.PrintDocument();
            this.prnDialogConsulta = new System.Windows.Forms.PrintDialog();
            this.btnPrintPreview = new System.Windows.Forms.Button();
            this.btnPrinterDialog = new System.Windows.Forms.Button();
            this.pnBotoes.SuspendLayout();
            this.pnCampos.SuspendLayout();
            this.pnParametros.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnBotoes
            // 
            this.pnBotoes.Controls.Add(this.btnPrinterDialog);
            this.pnBotoes.Controls.Add(this.btnPrintPreview);
            this.pnBotoes.Controls.Add(this.btnLimpar);
            this.pnBotoes.Controls.Add(this.btnImprimir);
            this.pnBotoes.Controls.SetChildIndex(this.btnImprimir, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnLimpar, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnPrintPreview, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnPrinterDialog, 0);
            // 
            // pnCampos
            // 
            this.pnCampos.Controls.Add(this.pnParametros);
            this.pnCampos.Controls.Add(this.lblTitulo);
            this.pnCampos.Size = new System.Drawing.Size(1018, 223);
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
            // 
            // btnFechar
            // 
            this.btnFechar.TabIndex = 7;
            // 
            // btnSobre
            // 
            this.btnSobre.TabIndex = 6;
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
            this.lblTitulo.TabIndex = 1;
            this.lblTitulo.Text = "Relatório Sintético de Fluxo de Caixa";
            this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pnParametros
            // 
            this.pnParametros.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnParametros.Controls.Add(this.lbContaCorrente);
            this.pnParametros.Controls.Add(this.lblTitContaCorrente);
            this.pnParametros.Controls.Add(this.txtDataCompetenciaFinal);
            this.pnParametros.Controls.Add(this.lblTitPeriodoCompetencia);
            this.pnParametros.Controls.Add(this.txtDataCompetenciaInicial);
            this.pnParametros.Controls.Add(this.lblDataCompetenciaAte);
            this.pnParametros.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnParametros.Location = new System.Drawing.Point(0, 40);
            this.pnParametros.Name = "pnParametros";
            this.pnParametros.Size = new System.Drawing.Size(1014, 182);
            this.pnParametros.TabIndex = 2;
            // 
            // lbContaCorrente
            // 
            this.lbContaCorrente.ColumnWidth = 400;
            this.lbContaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.lbContaCorrente.FormattingEnabled = true;
            this.lbContaCorrente.ItemHeight = 16;
            this.lbContaCorrente.Items.AddRange(new object[] {
            "item 1",
            "item 2",
            "item 3",
            "item 4",
            "item 5"});
            this.lbContaCorrente.Location = new System.Drawing.Point(184, 81);
            this.lbContaCorrente.Name = "lbContaCorrente";
            this.lbContaCorrente.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lbContaCorrente.Size = new System.Drawing.Size(400, 68);
            this.lbContaCorrente.TabIndex = 20;
            // 
            // lblTitContaCorrente
            // 
            this.lblTitContaCorrente.AutoSize = true;
            this.lblTitContaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitContaCorrente.Location = new System.Drawing.Point(86, 81);
            this.lblTitContaCorrente.Name = "lblTitContaCorrente";
            this.lblTitContaCorrente.Size = new System.Drawing.Size(92, 13);
            this.lblTitContaCorrente.TabIndex = 19;
            this.lblTitContaCorrente.Text = "Conta Corrente";
            // 
            // txtDataCompetenciaFinal
            // 
            this.txtDataCompetenciaFinal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDataCompetenciaFinal.Location = new System.Drawing.Point(301, 28);
            this.txtDataCompetenciaFinal.MaxLength = 10;
            this.txtDataCompetenciaFinal.Name = "txtDataCompetenciaFinal";
            this.txtDataCompetenciaFinal.Size = new System.Drawing.Size(91, 23);
            this.txtDataCompetenciaFinal.TabIndex = 1;
            this.txtDataCompetenciaFinal.Text = "01/01/2000";
            this.txtDataCompetenciaFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtDataCompetenciaFinal.Enter += new System.EventHandler(this.txtDataCompetenciaFinal_Enter);
            this.txtDataCompetenciaFinal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCompetenciaFinal_KeyDown);
            this.txtDataCompetenciaFinal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCompetenciaFinal_KeyPress);
            this.txtDataCompetenciaFinal.Leave += new System.EventHandler(this.txtDataCompetenciaFinal_Leave);
            // 
            // lblTitPeriodoCompetencia
            // 
            this.lblTitPeriodoCompetencia.AutoSize = true;
            this.lblTitPeriodoCompetencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitPeriodoCompetencia.Location = new System.Drawing.Point(98, 33);
            this.lblTitPeriodoCompetencia.Name = "lblTitPeriodoCompetencia";
            this.lblTitPeriodoCompetencia.Size = new System.Drawing.Size(80, 13);
            this.lblTitPeriodoCompetencia.TabIndex = 10;
            this.lblTitPeriodoCompetencia.Text = "Competência";
            // 
            // txtDataCompetenciaInicial
            // 
            this.txtDataCompetenciaInicial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDataCompetenciaInicial.Location = new System.Drawing.Point(184, 28);
            this.txtDataCompetenciaInicial.MaxLength = 10;
            this.txtDataCompetenciaInicial.Name = "txtDataCompetenciaInicial";
            this.txtDataCompetenciaInicial.Size = new System.Drawing.Size(91, 23);
            this.txtDataCompetenciaInicial.TabIndex = 0;
            this.txtDataCompetenciaInicial.Text = "01/01/2000";
            this.txtDataCompetenciaInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtDataCompetenciaInicial.Enter += new System.EventHandler(this.txtDataCompetenciaInicial_Enter);
            this.txtDataCompetenciaInicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCompetenciaInicial_KeyDown);
            this.txtDataCompetenciaInicial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCompetenciaInicial_KeyPress);
            this.txtDataCompetenciaInicial.Leave += new System.EventHandler(this.txtDataCompetenciaInicial_Leave);
            // 
            // lblDataCompetenciaAte
            // 
            this.lblDataCompetenciaAte.AutoSize = true;
            this.lblDataCompetenciaAte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDataCompetenciaAte.Location = new System.Drawing.Point(281, 33);
            this.lblDataCompetenciaAte.Name = "lblDataCompetenciaAte";
            this.lblDataCompetenciaAte.Size = new System.Drawing.Size(14, 13);
            this.lblDataCompetenciaAte.TabIndex = 9;
            this.lblDataCompetenciaAte.Text = "a";
            // 
            // btnImprimir
            // 
            this.btnImprimir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
            this.btnImprimir.Location = new System.Drawing.Point(789, 4);
            this.btnImprimir.Name = "btnImprimir";
            this.btnImprimir.Size = new System.Drawing.Size(40, 44);
            this.btnImprimir.TabIndex = 3;
            this.btnImprimir.TabStop = false;
            this.btnImprimir.UseVisualStyleBackColor = true;
            this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
            // 
            // btnLimpar
            // 
            this.btnLimpar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnLimpar.Image = ((System.Drawing.Image)(resources.GetObject("btnLimpar.Image")));
            this.btnLimpar.Location = new System.Drawing.Point(744, 4);
            this.btnLimpar.Name = "btnLimpar";
            this.btnLimpar.Size = new System.Drawing.Size(40, 44);
            this.btnLimpar.TabIndex = 1;
            this.btnLimpar.TabStop = false;
            this.btnLimpar.UseVisualStyleBackColor = true;
            this.btnLimpar.Click += new System.EventHandler(this.btnLimpar_Click);
            // 
            // prnPreviewConsulta
            // 
            this.prnPreviewConsulta.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.prnPreviewConsulta.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.prnPreviewConsulta.ClientSize = new System.Drawing.Size(400, 300);
            this.prnPreviewConsulta.Document = this.prnDocConsulta;
            this.prnPreviewConsulta.Enabled = true;
            this.prnPreviewConsulta.Icon = ((System.Drawing.Icon)(resources.GetObject("prnPreviewConsulta.Icon")));
            this.prnPreviewConsulta.Name = "prnPreview";
            this.prnPreviewConsulta.UseAntiAlias = true;
            this.prnPreviewConsulta.Visible = false;
            // 
            // prnDocConsulta
            // 
            this.prnDocConsulta.DocumentName = "Relatório Sintético de Fluxo de Caixa";
            this.prnDocConsulta.BeginPrint += new System.Drawing.Printing.PrintEventHandler(this.prnDocConsulta_BeginPrint);
            this.prnDocConsulta.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.prnDocConsulta_PrintPage);
            // 
            // prnDialogConsulta
            // 
            this.prnDialogConsulta.Document = this.prnDocConsulta;
            this.prnDialogConsulta.UseEXDialog = true;
            // 
            // btnPrintPreview
            // 
            this.btnPrintPreview.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnPrintPreview.Image = ((System.Drawing.Image)(resources.GetObject("btnPrintPreview.Image")));
            this.btnPrintPreview.Location = new System.Drawing.Point(834, 4);
            this.btnPrintPreview.Name = "btnPrintPreview";
            this.btnPrintPreview.Size = new System.Drawing.Size(40, 44);
            this.btnPrintPreview.TabIndex = 4;
            this.btnPrintPreview.TabStop = false;
            this.btnPrintPreview.UseVisualStyleBackColor = true;
            this.btnPrintPreview.Click += new System.EventHandler(this.btnPrintPreview_Click);
            // 
            // btnPrinterDialog
            // 
            this.btnPrinterDialog.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnPrinterDialog.Image = ((System.Drawing.Image)(resources.GetObject("btnPrinterDialog.Image")));
            this.btnPrinterDialog.Location = new System.Drawing.Point(879, 4);
            this.btnPrinterDialog.Name = "btnPrinterDialog";
            this.btnPrinterDialog.Size = new System.Drawing.Size(40, 44);
            this.btnPrinterDialog.TabIndex = 5;
            this.btnPrinterDialog.TabStop = false;
            this.btnPrinterDialog.UseVisualStyleBackColor = true;
            this.btnPrinterDialog.Click += new System.EventHandler(this.btnPrinterDialog_Click);
            // 
            // FFluxoRelatorioCtaCorrente
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.ClientSize = new System.Drawing.Size(1018, 320);
            this.KeyPreview = true;
            this.Name = "FFluxoRelatorioCtaCorrente";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FFluxoRelatorioCtaCorrente_FormClosing);
            this.Load += new System.EventHandler(this.FFluxoRelatorioCtaCorrente_Load);
            this.Shown += new System.EventHandler(this.FFluxoRelatorioCtaCorrente_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FFluxoRelatorioCtaCorrente_KeyDown);
            this.pnBotoes.ResumeLayout(false);
            this.pnCampos.ResumeLayout(false);
            this.pnParametros.ResumeLayout(false);
            this.pnParametros.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Panel pnParametros;
		private System.Windows.Forms.TextBox txtDataCompetenciaInicial;
		private System.Windows.Forms.Label lblDataCompetenciaAte;
		private System.Windows.Forms.Label lblTitPeriodoCompetencia;
		private System.Windows.Forms.TextBox txtDataCompetenciaFinal;
		private System.Windows.Forms.Label lblTitContaCorrente;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnLimpar;
		private System.Windows.Forms.PrintPreviewDialog prnPreviewConsulta;
		private System.Drawing.Printing.PrintDocument prnDocConsulta;
		private System.Windows.Forms.PrintDialog prnDialogConsulta;
		private System.Windows.Forms.Button btnPrinterDialog;
		private System.Windows.Forms.Button btnPrintPreview;
        private System.Windows.Forms.ListBox lbContaCorrente;
    }
}
