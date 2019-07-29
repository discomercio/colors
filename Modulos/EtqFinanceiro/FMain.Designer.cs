namespace EtqFinanceiro
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
			this.btnImprimirEtiquetasFin = new System.Windows.Forms.Button();
			this.btnImprimirEtiquetasFinDesc = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxMain.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxMain);
			this.pnCampos.Size = new System.Drawing.Size(1008, 280);
			this.pnCampos.Controls.SetChildIndex(this.gboxMain, 0);
			this.pnCampos.Controls.SetChildIndex(this.pnBotoes, 0);
			// 
			// gboxMain
			// 
			this.gboxMain.Controls.Add(this.btnImprimirEtiquetasFinDesc);
			this.gboxMain.Controls.Add(this.btnImprimirEtiquetasFin);
			this.gboxMain.Location = new System.Drawing.Point(264, 110);
			this.gboxMain.Name = "gboxMain";
			this.gboxMain.Size = new System.Drawing.Size(478, 129);
			this.gboxMain.TabIndex = 8;
			this.gboxMain.TabStop = false;
			// 
			// btnImprimirEtiquetasFin
			// 
			this.btnImprimirEtiquetasFin.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnImprimirEtiquetasFin.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnImprimirEtiquetasFin.ForeColor = System.Drawing.Color.Black;
			this.btnImprimirEtiquetasFin.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimirEtiquetasFin.Image")));
			this.btnImprimirEtiquetasFin.Location = new System.Drawing.Point(14, 16);
			this.btnImprimirEtiquetasFin.Name = "btnImprimirEtiquetasFin";
			this.btnImprimirEtiquetasFin.Size = new System.Drawing.Size(450, 38);
			this.btnImprimirEtiquetasFin.TabIndex = 2;
			this.btnImprimirEtiquetasFin.Text = "Imprimir Etiquetas (Relação de Depósitos)";
			this.btnImprimirEtiquetasFin.UseVisualStyleBackColor = true;
			this.btnImprimirEtiquetasFin.Click += new System.EventHandler(this.btnImprimirEtiquetasFin_Click);
			// 
			// btnImprimirEtiquetasFinDesc
			// 
			this.btnImprimirEtiquetasFinDesc.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnImprimirEtiquetasFinDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnImprimirEtiquetasFinDesc.ForeColor = System.Drawing.Color.Black;
			this.btnImprimirEtiquetasFinDesc.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimirEtiquetasFinDesc.Image")));
			this.btnImprimirEtiquetasFinDesc.Location = new System.Drawing.Point(14, 77);
			this.btnImprimirEtiquetasFinDesc.Name = "btnImprimirEtiquetasFinDesc";
			this.btnImprimirEtiquetasFinDesc.Size = new System.Drawing.Size(450, 38);
			this.btnImprimirEtiquetasFinDesc.TabIndex = 3;
			this.btnImprimirEtiquetasFinDesc.Text = "Imprimir Etiquetas (Relação de Depósitos c/ Desconto)";
			this.btnImprimirEtiquetasFinDesc.UseVisualStyleBackColor = true;
			this.btnImprimirEtiquetasFinDesc.Click += new System.EventHandler(this.btnImprimirEtiquetasFinDesc_Click);
			// 
			// FMain
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1008, 322);
			this.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.Name = "FMain";
			this.Text = "EtqFin  -  1.00 - XX.XXX.20XX";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FMain_FormClosing);
			this.Shown += new System.EventHandler(this.FMain_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.gboxMain.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox gboxMain;
        private System.Windows.Forms.Button btnImprimirEtiquetasFin;
		private System.Windows.Forms.Button btnImprimirEtiquetasFinDesc;
	}
}