namespace Reciprocidade
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
			this.btnGeraArqRemessaRetificacao = new System.Windows.Forms.Button();
			this.btnTrataOcorrencias = new System.Windows.Forms.Button();
			this.btnCarregaArqRetorno = new System.Windows.Forms.Button();
			this.btnGeraArqRemessa = new System.Windows.Forms.Button();
			this.btnGeraArqRemessaConciliacao = new System.Windows.Forms.Button();
			this.btnTrataOcorrenciasConciliacao = new System.Windows.Forms.Button();
			this.btnCarregaArqRetornoConciliacao = new System.Windows.Forms.Button();
			this.gboxConciliacao = new System.Windows.Forms.GroupBox();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.gboxMain.SuspendLayout();
			this.gboxConciliacao.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.gboxConciliacao);
			this.pnCampos.Location = new System.Drawing.Point(0, 24);
			this.pnCampos.Size = new System.Drawing.Size(1008, 439);
			// 
			// gboxMain
			// 
			this.gboxMain.Controls.Add(this.btnGeraArqRemessaRetificacao);
			this.gboxMain.Controls.Add(this.btnTrataOcorrencias);
			this.gboxMain.Controls.Add(this.btnCarregaArqRetorno);
			this.gboxMain.Controls.Add(this.btnGeraArqRemessa);
			this.gboxMain.Location = new System.Drawing.Point(264, 85);
			this.gboxMain.Name = "gboxMain";
			this.gboxMain.Size = new System.Drawing.Size(478, 197);
			this.gboxMain.TabIndex = 8;
			this.gboxMain.TabStop = false;
			// 
			// btnGeraArqRemessaRetificacao
			// 
			this.btnGeraArqRemessaRetificacao.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnGeraArqRemessaRetificacao.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnGeraArqRemessaRetificacao.ForeColor = System.Drawing.Color.Black;
			this.btnGeraArqRemessaRetificacao.Image = ((System.Drawing.Image)(resources.GetObject("btnGeraArqRemessaRetificacao.Image")));
			this.btnGeraArqRemessaRetificacao.Location = new System.Drawing.Point(14, 148);
			this.btnGeraArqRemessaRetificacao.Name = "btnGeraArqRemessaRetificacao";
			this.btnGeraArqRemessaRetificacao.Size = new System.Drawing.Size(450, 38);
			this.btnGeraArqRemessaRetificacao.TabIndex = 3;
			this.btnGeraArqRemessaRetificacao.Text = "Retorno: Gera Arquivo de Remessa";
			this.btnGeraArqRemessaRetificacao.UseVisualStyleBackColor = true;
			this.btnGeraArqRemessaRetificacao.Click += new System.EventHandler(this.btnGeraArqRemessaRetificacao_Click);
			// 
			// btnTrataOcorrencias
			// 
			this.btnTrataOcorrencias.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnTrataOcorrencias.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnTrataOcorrencias.ForeColor = System.Drawing.Color.Black;
			this.btnTrataOcorrencias.Image = ((System.Drawing.Image)(resources.GetObject("btnTrataOcorrencias.Image")));
			this.btnTrataOcorrencias.Location = new System.Drawing.Point(14, 104);
			this.btnTrataOcorrencias.Name = "btnTrataOcorrencias";
			this.btnTrataOcorrencias.Size = new System.Drawing.Size(450, 38);
			this.btnTrataOcorrencias.TabIndex = 2;
			this.btnTrataOcorrencias.Text = "Retorno: Trata Ocorrências";
			this.btnTrataOcorrencias.UseVisualStyleBackColor = true;
			this.btnTrataOcorrencias.Click += new System.EventHandler(this.btnTrataOcorrencias_Click);
			// 
			// btnCarregaArqRetorno
			// 
			this.btnCarregaArqRetorno.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnCarregaArqRetorno.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCarregaArqRetorno.ForeColor = System.Drawing.Color.Black;
			this.btnCarregaArqRetorno.Image = ((System.Drawing.Image)(resources.GetObject("btnCarregaArqRetorno.Image")));
			this.btnCarregaArqRetorno.Location = new System.Drawing.Point(14, 60);
			this.btnCarregaArqRetorno.Name = "btnCarregaArqRetorno";
			this.btnCarregaArqRetorno.Size = new System.Drawing.Size(450, 38);
			this.btnCarregaArqRetorno.TabIndex = 1;
			this.btnCarregaArqRetorno.Text = "Retorno: Carrega Arquivo de Retorno";
			this.btnCarregaArqRetorno.UseVisualStyleBackColor = true;
			this.btnCarregaArqRetorno.Click += new System.EventHandler(this.btnCarregaArqRetorno_Click);
			// 
			// btnGeraArqRemessa
			// 
			this.btnGeraArqRemessa.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnGeraArqRemessa.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnGeraArqRemessa.ForeColor = System.Drawing.Color.Black;
			this.btnGeraArqRemessa.Image = ((System.Drawing.Image)(resources.GetObject("btnGeraArqRemessa.Image")));
			this.btnGeraArqRemessa.Location = new System.Drawing.Point(14, 16);
			this.btnGeraArqRemessa.Name = "btnGeraArqRemessa";
			this.btnGeraArqRemessa.Size = new System.Drawing.Size(450, 38);
			this.btnGeraArqRemessa.TabIndex = 0;
			this.btnGeraArqRemessa.Text = "Normal: Gera Arquivo de Remessa";
			this.btnGeraArqRemessa.UseVisualStyleBackColor = true;
			this.btnGeraArqRemessa.Click += new System.EventHandler(this.btnGeraArqRemessa_Click);
			// 
			// btnGeraArqRemessaConciliacao
			// 
			this.btnGeraArqRemessaConciliacao.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnGeraArqRemessaConciliacao.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnGeraArqRemessaConciliacao.ForeColor = System.Drawing.Color.Black;
			this.btnGeraArqRemessaConciliacao.Image = ((System.Drawing.Image)(resources.GetObject("btnGeraArqRemessaConciliacao.Image")));
			this.btnGeraArqRemessaConciliacao.Location = new System.Drawing.Point(14, 104);
			this.btnGeraArqRemessaConciliacao.Name = "btnGeraArqRemessaConciliacao";
			this.btnGeraArqRemessaConciliacao.Size = new System.Drawing.Size(450, 38);
			this.btnGeraArqRemessaConciliacao.TabIndex = 2;
			this.btnGeraArqRemessaConciliacao.Text = "Conciliação: Gera Arquivo de Remessa";
			this.btnGeraArqRemessaConciliacao.UseVisualStyleBackColor = true;
			this.btnGeraArqRemessaConciliacao.Click += new System.EventHandler(this.btnGeraArqRemessaConciliacao_Click);
			// 
			// btnTrataOcorrenciasConciliacao
			// 
			this.btnTrataOcorrenciasConciliacao.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnTrataOcorrenciasConciliacao.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnTrataOcorrenciasConciliacao.ForeColor = System.Drawing.Color.Black;
			this.btnTrataOcorrenciasConciliacao.Image = ((System.Drawing.Image)(resources.GetObject("btnTrataOcorrenciasConciliacao.Image")));
			this.btnTrataOcorrenciasConciliacao.Location = new System.Drawing.Point(14, 60);
			this.btnTrataOcorrenciasConciliacao.Name = "btnTrataOcorrenciasConciliacao";
			this.btnTrataOcorrenciasConciliacao.Size = new System.Drawing.Size(450, 38);
			this.btnTrataOcorrenciasConciliacao.TabIndex = 1;
			this.btnTrataOcorrenciasConciliacao.Text = "Conciliação: Trata Ocorrências";
			this.btnTrataOcorrenciasConciliacao.UseVisualStyleBackColor = true;
			this.btnTrataOcorrenciasConciliacao.Click += new System.EventHandler(this.btnTrataOcorrenciasConciliacao_Click);
			// 
			// btnCarregaArqRetornoConciliacao
			// 
			this.btnCarregaArqRetornoConciliacao.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnCarregaArqRetornoConciliacao.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnCarregaArqRetornoConciliacao.ForeColor = System.Drawing.Color.Black;
			this.btnCarregaArqRetornoConciliacao.Image = ((System.Drawing.Image)(resources.GetObject("btnCarregaArqRetornoConciliacao.Image")));
			this.btnCarregaArqRetornoConciliacao.Location = new System.Drawing.Point(14, 16);
			this.btnCarregaArqRetornoConciliacao.Name = "btnCarregaArqRetornoConciliacao";
			this.btnCarregaArqRetornoConciliacao.Size = new System.Drawing.Size(450, 38);
			this.btnCarregaArqRetornoConciliacao.TabIndex = 0;
			this.btnCarregaArqRetornoConciliacao.Text = "Conciliação: Carrega Arquivo da Serasa";
			this.btnCarregaArqRetornoConciliacao.UseVisualStyleBackColor = true;
			this.btnCarregaArqRetornoConciliacao.Click += new System.EventHandler(this.btnCarregaArqRetornoConciliacao_Click);
			// 
			// gboxConciliacao
			// 
			this.gboxConciliacao.Controls.Add(this.btnGeraArqRemessaConciliacao);
			this.gboxConciliacao.Controls.Add(this.btnCarregaArqRetornoConciliacao);
			this.gboxConciliacao.Controls.Add(this.btnTrataOcorrenciasConciliacao);
			this.gboxConciliacao.Location = new System.Drawing.Point(264, 268);
			this.gboxConciliacao.Name = "gboxConciliacao";
			this.gboxConciliacao.Size = new System.Drawing.Size(478, 151);
			this.gboxConciliacao.TabIndex = 0;
			this.gboxConciliacao.TabStop = false;
			// 
			// FMain
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(1008, 481);
			this.Controls.Add(this.gboxMain);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "FMain";
			this.Text = "Serasa Reciprocidade  -  1.00 - 01.JUN.2014";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FMain_FormClosing);
			this.Shown += new System.EventHandler(this.FMain_Shown);
			this.Controls.SetChildIndex(this.pnCampos, 0);
			this.Controls.SetChildIndex(this.pnBotoes, 0);
			this.Controls.SetChildIndex(this.gboxMain, 0);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.gboxMain.ResumeLayout(false);
			this.gboxConciliacao.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox gboxMain;
		private System.Windows.Forms.Button btnGeraArqRemessa;
        private System.Windows.Forms.Button btnCarregaArqRetorno;
        private System.Windows.Forms.Button btnCarregaArqRetornoConciliacao;
        private System.Windows.Forms.Button btnGeraArqRemessaConciliacao;
        private System.Windows.Forms.Button btnTrataOcorrencias;
        private System.Windows.Forms.Button btnGeraArqRemessaRetificacao;
        private System.Windows.Forms.Button btnTrataOcorrenciasConciliacao;
		private System.Windows.Forms.GroupBox gboxConciliacao;
	}
}
