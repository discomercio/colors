namespace Financeiro
{
	partial class FBoletoHtml
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FBoletoHtml));
			this.webBrowser = new System.Windows.Forms.WebBrowser();
			this.btnPrinterDialog = new System.Windows.Forms.Button();
			this.btnEmail = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.prnDialogConsulta = new System.Windows.Forms.PrintDialog();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Controls.Add(this.btnPrinterDialog);
			this.pnBotoes.Controls.Add(this.btnEmail);
			this.pnBotoes.Controls.Add(this.btnImprimir);
			this.pnBotoes.Size = new System.Drawing.Size(818, 55);
			this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnImprimir, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnEmail, 0);
			this.pnBotoes.Controls.SetChildIndex(this.btnPrinterDialog, 0);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.webBrowser);
			this.pnCampos.Size = new System.Drawing.Size(818, 609);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.Location = new System.Drawing.Point(769, 4);
			this.btnFechar.TabIndex = 4;
			// 
			// btnSobre
			// 
			this.btnSobre.Location = new System.Drawing.Point(724, 4);
			this.btnSobre.TabIndex = 3;
			// 
			// webBrowser
			// 
			this.webBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
			this.webBrowser.Location = new System.Drawing.Point(0, 0);
			this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
			this.webBrowser.Name = "webBrowser";
			this.webBrowser.Size = new System.Drawing.Size(814, 605);
			this.webBrowser.TabIndex = 0;
			this.webBrowser.WebBrowserShortcutsEnabled = false;
			this.webBrowser.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser_DocumentCompleted);
			// 
			// btnPrinterDialog
			// 
			this.btnPrinterDialog.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnPrinterDialog.Image = ((System.Drawing.Image)(resources.GetObject("btnPrinterDialog.Image")));
			this.btnPrinterDialog.Location = new System.Drawing.Point(679, 4);
			this.btnPrinterDialog.Name = "btnPrinterDialog";
			this.btnPrinterDialog.Size = new System.Drawing.Size(40, 44);
			this.btnPrinterDialog.TabIndex = 2;
			this.btnPrinterDialog.TabStop = false;
			this.btnPrinterDialog.UseVisualStyleBackColor = true;
			this.btnPrinterDialog.Click += new System.EventHandler(this.btnPrinterDialog_Click);
			// 
			// btnEmail
			// 
			this.btnEmail.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnEmail.Image = ((System.Drawing.Image)(resources.GetObject("btnEmail.Image")));
			this.btnEmail.Location = new System.Drawing.Point(589, 4);
			this.btnEmail.Name = "btnEmail";
			this.btnEmail.Size = new System.Drawing.Size(40, 44);
			this.btnEmail.TabIndex = 0;
			this.btnEmail.TabStop = false;
			this.btnEmail.UseVisualStyleBackColor = true;
			this.btnEmail.Click += new System.EventHandler(this.btnEmail_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.Location = new System.Drawing.Point(634, 4);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(40, 44);
			this.btnImprimir.TabIndex = 1;
			this.btnImprimir.TabStop = false;
			this.btnImprimir.UseVisualStyleBackColor = true;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// prnDialogConsulta
			// 
			this.prnDialogConsulta.UseEXDialog = true;
			// 
			// FBoletoHtml
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(818, 706);
			this.Name = "FBoletoHtml";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.Load += new System.EventHandler(this.FBoletoHtml_Load);
			this.Shown += new System.EventHandler(this.FBoletoHtml_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.WebBrowser webBrowser;
		private System.Windows.Forms.Button btnPrinterDialog;
		private System.Windows.Forms.Button btnEmail;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.PrintDialog prnDialogConsulta;
	}
}
