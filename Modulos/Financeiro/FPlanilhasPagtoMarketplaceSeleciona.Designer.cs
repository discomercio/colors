namespace Financeiro
{
    partial class FPlanilhasPagtoMarketplaceSeleciona
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FPlanilhasPagtoMarketplaceSeleciona));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cbPlanilhaPagamentoEmpresas = new System.Windows.Forms.ComboBox();
            this.lblSelecioneEmpresa = new System.Windows.Forms.Label();
            this.btnCancela = new System.Windows.Forms.Button();
            this.btnAvancar = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cbPlanilhaPagamentoEmpresas);
            this.groupBox1.Controls.Add(this.lblSelecioneEmpresa);
            this.groupBox1.Location = new System.Drawing.Point(13, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(518, 112);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            // 
            // cbPlanilhaPagamentoEmpresas
            // 
            this.cbPlanilhaPagamentoEmpresas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbPlanilhaPagamentoEmpresas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbPlanilhaPagamentoEmpresas.FormattingEnabled = true;
            this.cbPlanilhaPagamentoEmpresas.Location = new System.Drawing.Point(63, 63);
            this.cbPlanilhaPagamentoEmpresas.Name = "cbPlanilhaPagamentoEmpresas";
            this.cbPlanilhaPagamentoEmpresas.Size = new System.Drawing.Size(392, 21);
            this.cbPlanilhaPagamentoEmpresas.TabIndex = 17;
            // 
            // lblSelecioneEmpresa
            // 
            this.lblSelecioneEmpresa.AutoSize = true;
            this.lblSelecioneEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSelecioneEmpresa.Location = new System.Drawing.Point(136, 32);
            this.lblSelecioneEmpresa.Name = "lblSelecioneEmpresa";
            this.lblSelecioneEmpresa.Size = new System.Drawing.Size(247, 16);
            this.lblSelecioneEmpresa.TabIndex = 16;
            this.lblSelecioneEmpresa.Text = "Selecione a Empresa Marketplace";
            // 
            // btnCancela
            // 
            this.btnCancela.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancela.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCancela.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancela.Image = ((System.Drawing.Image)(resources.GetObject("btnCancela.Image")));
            this.btnCancela.Location = new System.Drawing.Point(293, 137);
            this.btnCancela.Name = "btnCancela";
            this.btnCancela.Size = new System.Drawing.Size(100, 40);
            this.btnCancela.TabIndex = 14;
            this.btnCancela.Text = "&Cancelar";
            this.btnCancela.UseVisualStyleBackColor = true;
            this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
            // 
            // btnAvancar
            // 
            this.btnAvancar.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAvancar.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAvancar.Image = ((System.Drawing.Image)(resources.GetObject("btnAvancar.Image")));
            this.btnAvancar.Location = new System.Drawing.Point(149, 137);
            this.btnAvancar.Name = "btnAvancar";
            this.btnAvancar.Size = new System.Drawing.Size(100, 40);
            this.btnAvancar.TabIndex = 13;
            this.btnAvancar.Text = "&Avançar";
            this.btnAvancar.UseVisualStyleBackColor = true;
            this.btnAvancar.Click += new System.EventHandler(this.btnAvancar_Click);
            // 
            // FPlanilhasPagtoMarketplaceSeleciona
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(543, 196);
            this.Controls.Add(this.btnCancela);
            this.Controls.Add(this.btnAvancar);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FPlanilhasPagtoMarketplaceSeleciona";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Planilhas de Pagamentos Marketplace";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FPlanilhasPagtoMarketplaceSeleciona_FormClosing);
            this.Load += new System.EventHandler(this.FPlanilhasPagtoMarketplaceSeleciona_Load);
            this.Shown += new System.EventHandler(this.FPlanilhasPagtoMarketplaceSeleciona_Shown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cbPlanilhaPagamentoEmpresas;
        private System.Windows.Forms.Label lblSelecioneEmpresa;
        private System.Windows.Forms.Button btnCancela;
        private System.Windows.Forms.Button btnAvancar;
    }
}