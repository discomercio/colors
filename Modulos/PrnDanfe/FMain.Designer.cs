namespace PrnDANFE
{
	partial class FMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FMain));
            this.gboxMain = new System.Windows.Forms.GroupBox();
            this.btnImprimirDANFEPDF = new System.Windows.Forms.Button();
            this.pnBotoes.SuspendLayout();
            this.gboxMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
            // 
            // pnCampos
            // 
            this.pnCampos.Location = new System.Drawing.Point(0, 24);
            this.pnCampos.Size = new System.Drawing.Size(1008, 280);
            // 
            // gboxMain
            // 
            this.gboxMain.Controls.Add(this.btnImprimirDANFEPDF);
            this.gboxMain.Location = new System.Drawing.Point(264, 144);
            this.gboxMain.Name = "gboxMain";
            this.gboxMain.Size = new System.Drawing.Size(478, 67);
            this.gboxMain.TabIndex = 8;
            this.gboxMain.TabStop = false;
            // 
            // btnImprimirDANFEPDF
            // 
            this.btnImprimirDANFEPDF.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnImprimirDANFEPDF.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImprimirDANFEPDF.ForeColor = System.Drawing.Color.Black;
            this.btnImprimirDANFEPDF.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimirDANFEPDF.Image")));
            this.btnImprimirDANFEPDF.Location = new System.Drawing.Point(14, 16);
            this.btnImprimirDANFEPDF.Name = "btnImprimirDANFEPDF";
            this.btnImprimirDANFEPDF.Size = new System.Drawing.Size(450, 38);
            this.btnImprimirDANFEPDF.TabIndex = 1;
            this.btnImprimirDANFEPDF.Text = "Imprimir DANFE (PDF)";
            this.btnImprimirDANFEPDF.UseVisualStyleBackColor = true;
            this.btnImprimirDANFEPDF.Click += new System.EventHandler(this.btnImprimirDANFEPDF_Click);
            // 
            // FMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.ClientSize = new System.Drawing.Size(1008, 322);
            this.Controls.Add(this.gboxMain);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FMain";
            this.Text = "PrnDANFE  -  1.03 - 15.ABR.2014";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FMain_FormClosing);
            this.Shown += new System.EventHandler(this.FMain_Shown);
            this.Controls.SetChildIndex(this.pnCampos, 0);
            this.Controls.SetChildIndex(this.pnBotoes, 0);
            this.Controls.SetChildIndex(this.gboxMain, 0);
            this.pnBotoes.ResumeLayout(false);
            this.gboxMain.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxMain;
        private System.Windows.Forms.Button btnImprimirDANFEPDF;
	}
}
