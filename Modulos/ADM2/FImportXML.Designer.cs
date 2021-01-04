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
            this.pnBotoes.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
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
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}
