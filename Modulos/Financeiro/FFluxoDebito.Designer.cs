namespace Financeiro
{
	partial class FFluxoDebito
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FFluxoDebito));
			this.lblTitulo = new System.Windows.Forms.Label();
			this.lblContaCorrente = new System.Windows.Forms.Label();
			this.cbContaCorrente = new System.Windows.Forms.ComboBox();
			this.lblPlanoContasEmpresa = new System.Windows.Forms.Label();
			this.cbPlanoContasEmpresa = new System.Windows.Forms.ComboBox();
			this.lblPlanoContasConta = new System.Windows.Forms.Label();
			this.cbPlanoContasConta = new System.Windows.Forms.ComboBox();
			this.lblDataCompetencia = new System.Windows.Forms.Label();
			this.lblValor = new System.Windows.Forms.Label();
			this.lblDescricao = new System.Windows.Forms.Label();
			this.txtDescricao = new System.Windows.Forms.TextBox();
			this.txtDataCompetencia = new System.Windows.Forms.TextBox();
			this.txtValor = new System.Windows.Forms.TextBox();
			this.btnGravar = new System.Windows.Forms.Button();
			this.btnLimpar = new System.Windows.Forms.Button();
			this.lblCnpjCpf = new System.Windows.Forms.Label();
			this.txtCnpjCpf = new System.Windows.Forms.TextBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.lblComp2 = new System.Windows.Forms.Label();
			this.txtComp2 = new System.Windows.Forms.TextBox();
			this.txtNF = new System.Windows.Forms.TextBox();
			this.lblNF = new System.Windows.Forms.Label();
			this.pnBotoes.SuspendLayout();
			this.pnCampos.SuspendLayout();
			this.SuspendLayout();
			// 
			// pnBotoes
			// 
			this.pnBotoes.Size = new System.Drawing.Size(675, 55);
			// 
			// pnCampos
			// 
			this.pnCampos.Controls.Add(this.txtNF);
			this.pnCampos.Controls.Add(this.lblNF);
			this.pnCampos.Controls.Add(this.txtComp2);
			this.pnCampos.Controls.Add(this.lblComp2);
			this.pnCampos.Controls.Add(this.lblContador);
			this.pnCampos.Controls.Add(this.txtCnpjCpf);
			this.pnCampos.Controls.Add(this.lblCnpjCpf);
			this.pnCampos.Controls.Add(this.btnLimpar);
			this.pnCampos.Controls.Add(this.btnGravar);
			this.pnCampos.Controls.Add(this.txtValor);
			this.pnCampos.Controls.Add(this.txtDataCompetencia);
			this.pnCampos.Controls.Add(this.txtDescricao);
			this.pnCampos.Controls.Add(this.lblDescricao);
			this.pnCampos.Controls.Add(this.lblValor);
			this.pnCampos.Controls.Add(this.lblDataCompetencia);
			this.pnCampos.Controls.Add(this.cbPlanoContasConta);
			this.pnCampos.Controls.Add(this.lblPlanoContasConta);
			this.pnCampos.Controls.Add(this.cbPlanoContasEmpresa);
			this.pnCampos.Controls.Add(this.lblPlanoContasEmpresa);
			this.pnCampos.Controls.Add(this.cbContaCorrente);
			this.pnCampos.Controls.Add(this.lblContaCorrente);
			this.pnCampos.Controls.Add(this.lblTitulo);
			this.pnCampos.Size = new System.Drawing.Size(675, 339);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(375, -200);
			// 
			// btnFechar
			// 
			this.btnFechar.Location = new System.Drawing.Point(626, 4);
			this.btnFechar.TabIndex = 1;
			// 
			// btnSobre
			// 
			this.btnSobre.Location = new System.Drawing.Point(581, 4);
			this.btnSobre.TabIndex = 0;
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
			this.lblTitulo.Size = new System.Drawing.Size(671, 40);
			this.lblTitulo.TabIndex = 0;
			this.lblTitulo.Text = "Lançamento de Débito";
			this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblContaCorrente
			// 
			this.lblContaCorrente.AutoSize = true;
			this.lblContaCorrente.Location = new System.Drawing.Point(48, 55);
			this.lblContaCorrente.Name = "lblContaCorrente";
			this.lblContaCorrente.Size = new System.Drawing.Size(78, 13);
			this.lblContaCorrente.TabIndex = 1;
			this.lblContaCorrente.Text = "Conta Corrente";
			// 
			// cbContaCorrente
			// 
			this.cbContaCorrente.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbContaCorrente.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbContaCorrente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbContaCorrente.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbContaCorrente.FormattingEnabled = true;
			this.cbContaCorrente.Location = new System.Drawing.Point(132, 50);
			this.cbContaCorrente.MaxDropDownItems = 12;
			this.cbContaCorrente.Name = "cbContaCorrente";
			this.cbContaCorrente.Size = new System.Drawing.Size(518, 24);
			this.cbContaCorrente.TabIndex = 0;
			this.cbContaCorrente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbContaCorrente_KeyDown);
			// 
			// lblPlanoContasEmpresa
			// 
			this.lblPlanoContasEmpresa.AutoSize = true;
			this.lblPlanoContasEmpresa.Location = new System.Drawing.Point(78, 91);
			this.lblPlanoContasEmpresa.Name = "lblPlanoContasEmpresa";
			this.lblPlanoContasEmpresa.Size = new System.Drawing.Size(48, 13);
			this.lblPlanoContasEmpresa.TabIndex = 3;
			this.lblPlanoContasEmpresa.Text = "Empresa";
			// 
			// cbPlanoContasEmpresa
			// 
			this.cbPlanoContasEmpresa.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasEmpresa.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasEmpresa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasEmpresa.FormattingEnabled = true;
			this.cbPlanoContasEmpresa.Location = new System.Drawing.Point(132, 86);
			this.cbPlanoContasEmpresa.MaxDropDownItems = 12;
			this.cbPlanoContasEmpresa.Name = "cbPlanoContasEmpresa";
			this.cbPlanoContasEmpresa.Size = new System.Drawing.Size(518, 24);
			this.cbPlanoContasEmpresa.TabIndex = 1;
			this.cbPlanoContasEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasEmpresa_KeyDown);
			// 
			// lblPlanoContasConta
			// 
			this.lblPlanoContasConta.AutoSize = true;
			this.lblPlanoContasConta.Location = new System.Drawing.Point(46, 127);
			this.lblPlanoContasConta.Name = "lblPlanoContasConta";
			this.lblPlanoContasConta.Size = new System.Drawing.Size(80, 13);
			this.lblPlanoContasConta.TabIndex = 5;
			this.lblPlanoContasConta.Text = "Plano de Conta";
			// 
			// cbPlanoContasConta
			// 
			this.cbPlanoContasConta.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
			this.cbPlanoContasConta.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.cbPlanoContasConta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbPlanoContasConta.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.cbPlanoContasConta.FormattingEnabled = true;
			this.cbPlanoContasConta.Location = new System.Drawing.Point(132, 122);
			this.cbPlanoContasConta.MaxDropDownItems = 12;
			this.cbPlanoContasConta.Name = "cbPlanoContasConta";
			this.cbPlanoContasConta.Size = new System.Drawing.Size(518, 24);
			this.cbPlanoContasConta.TabIndex = 2;
			this.cbPlanoContasConta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbPlanoContasConta_KeyDown);
			// 
			// lblDataCompetencia
			// 
			this.lblDataCompetencia.AutoSize = true;
			this.lblDataCompetencia.Location = new System.Drawing.Point(16, 163);
			this.lblDataCompetencia.Name = "lblDataCompetencia";
			this.lblDataCompetencia.Size = new System.Drawing.Size(110, 13);
			this.lblDataCompetencia.TabIndex = 7;
			this.lblDataCompetencia.Text = "Data de Competência";
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.Location = new System.Drawing.Point(72, 196);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(54, 13);
			this.lblValor.TabIndex = 9;
			this.lblValor.Text = "Valor (R$)";
			// 
			// lblDescricao
			// 
			this.lblDescricao.AutoSize = true;
			this.lblDescricao.Location = new System.Drawing.Point(71, 236);
			this.lblDescricao.Name = "lblDescricao";
			this.lblDescricao.Size = new System.Drawing.Size(55, 13);
			this.lblDescricao.TabIndex = 11;
			this.lblDescricao.Text = "Descrição";
			// 
			// txtDescricao
			// 
			this.txtDescricao.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDescricao.Location = new System.Drawing.Point(132, 231);
			this.txtDescricao.MaxLength = 40;
			this.txtDescricao.Name = "txtDescricao";
			this.txtDescricao.Size = new System.Drawing.Size(518, 23);
			this.txtDescricao.TabIndex = 8;
			this.txtDescricao.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescricao_KeyDown);
			this.txtDescricao.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescricao_KeyPress);
			// 
			// txtDataCompetencia
			// 
			this.txtDataCompetencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtDataCompetencia.Location = new System.Drawing.Point(132, 158);
			this.txtDataCompetencia.MaxLength = 10;
			this.txtDataCompetencia.Name = "txtDataCompetencia";
			this.txtDataCompetencia.Size = new System.Drawing.Size(91, 23);
			this.txtDataCompetencia.TabIndex = 3;
			this.txtDataCompetencia.Text = "01/01/2000";
			this.txtDataCompetencia.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtDataCompetencia.Enter += new System.EventHandler(this.txtDataCompetencia_Enter);
			this.txtDataCompetencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDataCompetencia_KeyDown);
			this.txtDataCompetencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDataCompetencia_KeyPress);
			this.txtDataCompetencia.Leave += new System.EventHandler(this.txtDataCompetencia_Leave);
			// 
			// txtValor
			// 
			this.txtValor.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtValor.Location = new System.Drawing.Point(132, 191);
			this.txtValor.MaxLength = 18;
			this.txtValor.Name = "txtValor";
			this.txtValor.Size = new System.Drawing.Size(111, 23);
			this.txtValor.TabIndex = 5;
			this.txtValor.Text = "999.999.999,99";
			this.txtValor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
			this.txtValor.Leave += new System.EventHandler(this.txtValor_Leave);
			// 
			// btnGravar
			// 
			this.btnGravar.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnGravar.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnGravar.Image = ((System.Drawing.Image)(resources.GetObject("btnGravar.Image")));
			this.btnGravar.Location = new System.Drawing.Point(132, 278);
			this.btnGravar.Name = "btnGravar";
			this.btnGravar.Size = new System.Drawing.Size(129, 40);
			this.btnGravar.TabIndex = 9;
			this.btnGravar.Text = "&Gravar";
			this.btnGravar.UseVisualStyleBackColor = true;
			this.btnGravar.Click += new System.EventHandler(this.btnGravar_Click);
			// 
			// btnLimpar
			// 
			this.btnLimpar.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnLimpar.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnLimpar.Image = ((System.Drawing.Image)(resources.GetObject("btnLimpar.Image")));
			this.btnLimpar.Location = new System.Drawing.Point(521, 278);
			this.btnLimpar.Name = "btnLimpar";
			this.btnLimpar.Size = new System.Drawing.Size(129, 40);
			this.btnLimpar.TabIndex = 10;
			this.btnLimpar.Text = "&Limpar";
			this.btnLimpar.UseVisualStyleBackColor = true;
			this.btnLimpar.Click += new System.EventHandler(this.btnLimpar_Click);
			// 
			// lblCnpjCpf
			// 
			this.lblCnpjCpf.AutoSize = true;
			this.lblCnpjCpf.Location = new System.Drawing.Point(267, 196);
			this.lblCnpjCpf.Name = "lblCnpjCpf";
			this.lblCnpjCpf.Size = new System.Drawing.Size(59, 13);
			this.lblCnpjCpf.TabIndex = 12;
			this.lblCnpjCpf.Text = "CNPJ/CPF";
			// 
			// txtCnpjCpf
			// 
			this.txtCnpjCpf.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtCnpjCpf.Location = new System.Drawing.Point(332, 191);
			this.txtCnpjCpf.MaxLength = 18;
			this.txtCnpjCpf.Name = "txtCnpjCpf";
			this.txtCnpjCpf.Size = new System.Drawing.Size(145, 23);
			this.txtCnpjCpf.TabIndex = 6;
			this.txtCnpjCpf.Text = "00.000.000/0000-00";
			this.txtCnpjCpf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtCnpjCpf.Enter += new System.EventHandler(this.txtCnpjCpf_Enter);
			this.txtCnpjCpf.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCnpjCpf_KeyDown);
			this.txtCnpjCpf.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCnpjCpf_KeyPress);
			this.txtCnpjCpf.Leave += new System.EventHandler(this.txtCnpjCpf_Leave);
			// 
			// lblContador
			// 
			this.lblContador.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 22F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblContador.Image = ((System.Drawing.Image)(resources.GetObject("lblContador.Image")));
			this.lblContador.Location = new System.Drawing.Point(10, 278);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(56, 40);
			this.lblContador.TabIndex = 13;
			this.lblContador.Text = "99";
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblComp2
			// 
			this.lblComp2.AutoSize = true;
			this.lblComp2.Location = new System.Drawing.Point(286, 163);
			this.lblComp2.Name = "lblComp2";
			this.lblComp2.Size = new System.Drawing.Size(40, 13);
			this.lblComp2.TabIndex = 14;
			this.lblComp2.Text = "Comp2";
			// 
			// txtComp2
			// 
			this.txtComp2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtComp2.Location = new System.Drawing.Point(332, 158);
			this.txtComp2.MaxLength = 7;
			this.txtComp2.Name = "txtComp2";
			this.txtComp2.Size = new System.Drawing.Size(91, 23);
			this.txtComp2.TabIndex = 4;
			this.txtComp2.Text = "01/2000";
			this.txtComp2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtComp2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtComp2_KeyDown);
			this.txtComp2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtComp2_KeyPress);
			this.txtComp2.Leave += new System.EventHandler(this.txtComp2_Leave);
			// 
			// txtNF
			// 
			this.txtNF.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNF.Location = new System.Drawing.Point(557, 191);
			this.txtNF.MaxLength = 18;
			this.txtNF.Name = "txtNF";
			this.txtNF.Size = new System.Drawing.Size(93, 23);
			this.txtNF.TabIndex = 7;
			this.txtNF.Text = "999.999.999";
			this.txtNF.Enter += new System.EventHandler(this.txtNF_Enter);
			this.txtNF.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNF_KeyDown);
			this.txtNF.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNF_KeyPress);
			this.txtNF.Leave += new System.EventHandler(this.txtNF_Leave);
			// 
			// lblNF
			// 
			this.lblNF.AutoSize = true;
			this.lblNF.Location = new System.Drawing.Point(530, 196);
			this.lblNF.Name = "lblNF";
			this.lblNF.Size = new System.Drawing.Size(21, 13);
			this.lblNF.TabIndex = 16;
			this.lblNF.Text = "NF";
			// 
			// FFluxoDebito
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.ClientSize = new System.Drawing.Size(675, 436);
			this.Name = "FFluxoDebito";
			this.Text = "Artven - Financeiro  -  1.00 - xx.JUL.2009";
			this.Load += new System.EventHandler(this.FFluxoDebito_Load);
			this.Shown += new System.EventHandler(this.FFluxoDebito_Shown);
			this.pnBotoes.ResumeLayout(false);
			this.pnCampos.ResumeLayout(false);
			this.pnCampos.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.ComboBox cbContaCorrente;
		private System.Windows.Forms.Label lblContaCorrente;
		private System.Windows.Forms.ComboBox cbPlanoContasConta;
		private System.Windows.Forms.Label lblPlanoContasConta;
		private System.Windows.Forms.ComboBox cbPlanoContasEmpresa;
		private System.Windows.Forms.Label lblPlanoContasEmpresa;
		private System.Windows.Forms.Label lblDataCompetencia;
		private System.Windows.Forms.Label lblValor;
		private System.Windows.Forms.TextBox txtDescricao;
		private System.Windows.Forms.Label lblDescricao;
		private System.Windows.Forms.TextBox txtValor;
		private System.Windows.Forms.TextBox txtDataCompetencia;
		private System.Windows.Forms.Button btnLimpar;
		private System.Windows.Forms.Button btnGravar;
		private System.Windows.Forms.TextBox txtCnpjCpf;
		private System.Windows.Forms.Label lblCnpjCpf;
		private System.Windows.Forms.Label lblContador;
        private System.Windows.Forms.Label lblComp2;
        private System.Windows.Forms.TextBox txtComp2;
		private System.Windows.Forms.TextBox txtNF;
		private System.Windows.Forms.Label lblNF;
	}
}
