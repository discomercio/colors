namespace ADM2
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
			this.gboxIBPT = new System.Windows.Forms.GroupBox();
			this.btnIbptCarregaArqCsv = new System.Windows.Forms.Button();
			this.btnAtualizarPlanilhaEstoque = new System.Windows.Forms.Button();
			this.pnBotoes.SuspendLayout();
			this.gboxIBPT.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// pnCampos
			// 
			this.pnCampos.Location = new System.Drawing.Point(0, 24);
			this.pnCampos.Size = new System.Drawing.Size(1008, 339);
			// 
			// gboxIBPT
			// 
			this.gboxIBPT.Controls.Add(this.btnAtualizarPlanilhaEstoque);
			this.gboxIBPT.Controls.Add(this.btnIbptCarregaArqCsv);
			this.gboxIBPT.Location = new System.Drawing.Point(264, 144);
			this.gboxIBPT.Name = "gboxIBPT";
			this.gboxIBPT.Size = new System.Drawing.Size(478, 126);
			this.gboxIBPT.TabIndex = 8;
			this.gboxIBPT.TabStop = false;
			// 
			// btnIbptCarregaArqCsv
			// 
			this.btnIbptCarregaArqCsv.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnIbptCarregaArqCsv.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnIbptCarregaArqCsv.ForeColor = System.Drawing.Color.Black;
			this.btnIbptCarregaArqCsv.Image = ((System.Drawing.Image)(resources.GetObject("btnIbptCarregaArqCsv.Image")));
			this.btnIbptCarregaArqCsv.Location = new System.Drawing.Point(14, 16);
			this.btnIbptCarregaArqCsv.Name = "btnIbptCarregaArqCsv";
			this.btnIbptCarregaArqCsv.Size = new System.Drawing.Size(450, 38);
			this.btnIbptCarregaArqCsv.TabIndex = 1;
			this.btnIbptCarregaArqCsv.Text = "Carrega Arquivo IBPT";
			this.btnIbptCarregaArqCsv.UseVisualStyleBackColor = true;
			this.btnIbptCarregaArqCsv.Click += new System.EventHandler(this.btnIbptCarregaArqCsv_Click);
			// 
			// btnAtualizarPlanilhaEstoque
			// 
			this.btnAtualizarPlanilhaEstoque.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnAtualizarPlanilhaEstoque.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnAtualizarPlanilhaEstoque.ForeColor = System.Drawing.Color.Black;
			this.btnAtualizarPlanilhaEstoque.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizarPlanilhaEstoque.Image")));
			this.btnAtualizarPlanilhaEstoque.Location = new System.Drawing.Point(14, 75);
			this.btnAtualizarPlanilhaEstoque.Name = "btnAtualizarPlanilhaEstoque";
			this.btnAtualizarPlanilhaEstoque.Size = new System.Drawing.Size(450, 38);
			this.btnAtualizarPlanilhaEstoque.TabIndex = 2;
			this.btnAtualizarPlanilhaEstoque.Text = "Atualizar Planilha do Estoque";
			this.btnAtualizarPlanilhaEstoque.UseVisualStyleBackColor = true;
			this.btnAtualizarPlanilhaEstoque.Click += new System.EventHandler(this.btnAtualizarPlanilhaEstoque_Click);
			// 
			// FMain
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1008, 381);
			this.Controls.Add(this.gboxIBPT);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "FMain";
			this.Text = "ADM2  -  1.00 - 01.JUN.2013";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FMain_FormClosing);
			this.Shown += new System.EventHandler(this.FMain_Shown);
			this.Controls.SetChildIndex(this.pnCampos, 0);
			this.Controls.SetChildIndex(this.pnBotoes, 0);
			this.Controls.SetChildIndex(this.gboxIBPT, 0);
			this.pnBotoes.ResumeLayout(false);
			this.gboxIBPT.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxIBPT;
		private System.Windows.Forms.Button btnIbptCarregaArqCsv;
		private System.Windows.Forms.Button btnAtualizarPlanilhaEstoque;
	}
}
