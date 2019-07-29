namespace EtqWms
{
	partial class FVolumeSeleciona
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FVolumeSeleciona));
			this.btnOk = new System.Windows.Forms.Button();
			this.btnCancela = new System.Windows.Forms.Button();
			this.btnDummy = new System.Windows.Forms.Button();
			this.gboxNf = new System.Windows.Forms.GroupBox();
			this.txtNF = new System.Windows.Forms.TextBox();
			this.lblTitNf = new System.Windows.Forms.Label();
			this.gboxOpcao = new System.Windows.Forms.GroupBox();
			this.lblTitIntervaloAte = new System.Windows.Forms.Label();
			this.txtIntervaloFim = new System.Windows.Forms.TextBox();
			this.txtIntervaloInicio = new System.Windows.Forms.TextBox();
			this.txtVolumeUnico = new System.Windows.Forms.TextBox();
			this.rbVolumeRange = new System.Windows.Forms.RadioButton();
			this.rbVolumeUnico = new System.Windows.Forms.RadioButton();
			this.gboxNf.SuspendLayout();
			this.gboxOpcao.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnOk
			// 
			this.btnOk.Image = ((System.Drawing.Image)(resources.GetObject("btnOk.Image")));
			this.btnOk.Location = new System.Drawing.Point(50, 219);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(85, 31);
			this.btnOk.TabIndex = 3;
			this.btnOk.UseVisualStyleBackColor = true;
			this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
			// 
			// btnCancela
			// 
			this.btnCancela.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancela.Image = ((System.Drawing.Image)(resources.GetObject("btnCancela.Image")));
			this.btnCancela.Location = new System.Drawing.Point(298, 219);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(75, 31);
			this.btnCancela.TabIndex = 4;
			this.btnCancela.UseVisualStyleBackColor = true;
			this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
			// 
			// btnDummy
			// 
			this.btnDummy.Location = new System.Drawing.Point(184, 253);
			this.btnDummy.Name = "btnDummy";
			this.btnDummy.Size = new System.Drawing.Size(75, 23);
			this.btnDummy.TabIndex = 0;
			this.btnDummy.Text = "btnDummy";
			this.btnDummy.UseVisualStyleBackColor = true;
			// 
			// gboxNf
			// 
			this.gboxNf.Controls.Add(this.txtNF);
			this.gboxNf.Controls.Add(this.lblTitNf);
			this.gboxNf.Location = new System.Drawing.Point(50, 12);
			this.gboxNf.Name = "gboxNf";
			this.gboxNf.Size = new System.Drawing.Size(323, 50);
			this.gboxNf.TabIndex = 1;
			this.gboxNf.TabStop = false;
			// 
			// txtNF
			// 
			this.txtNF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtNF.Location = new System.Drawing.Point(61, 18);
			this.txtNF.Name = "txtNF";
			this.txtNF.Size = new System.Drawing.Size(72, 20);
			this.txtNF.TabIndex = 1;
			this.txtNF.Text = "999";
			this.txtNF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNF.Enter += new System.EventHandler(this.txtNF_Enter);
			this.txtNF.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNF_KeyDown);
			this.txtNF.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNF_KeyPress);
			this.txtNF.Leave += new System.EventHandler(this.txtNF_Leave);
			// 
			// lblTitNf
			// 
			this.lblTitNf.AutoSize = true;
			this.lblTitNf.Location = new System.Drawing.Point(34, 21);
			this.lblTitNf.Name = "lblTitNf";
			this.lblTitNf.Size = new System.Drawing.Size(21, 13);
			this.lblTitNf.TabIndex = 0;
			this.lblTitNf.Text = "NF";
			// 
			// gboxOpcao
			// 
			this.gboxOpcao.Controls.Add(this.lblTitIntervaloAte);
			this.gboxOpcao.Controls.Add(this.txtIntervaloFim);
			this.gboxOpcao.Controls.Add(this.txtIntervaloInicio);
			this.gboxOpcao.Controls.Add(this.txtVolumeUnico);
			this.gboxOpcao.Controls.Add(this.rbVolumeRange);
			this.gboxOpcao.Controls.Add(this.rbVolumeUnico);
			this.gboxOpcao.Location = new System.Drawing.Point(50, 78);
			this.gboxOpcao.Name = "gboxOpcao";
			this.gboxOpcao.Size = new System.Drawing.Size(323, 113);
			this.gboxOpcao.TabIndex = 2;
			this.gboxOpcao.TabStop = false;
			// 
			// lblTitIntervaloAte
			// 
			this.lblTitIntervaloAte.AutoSize = true;
			this.lblTitIntervaloAte.Location = new System.Drawing.Point(187, 75);
			this.lblTitIntervaloAte.Name = "lblTitIntervaloAte";
			this.lblTitIntervaloAte.Size = new System.Drawing.Size(22, 13);
			this.lblTitIntervaloAte.TabIndex = 4;
			this.lblTitIntervaloAte.Text = "até";
			// 
			// txtIntervaloFim
			// 
			this.txtIntervaloFim.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtIntervaloFim.Location = new System.Drawing.Point(216, 72);
			this.txtIntervaloFim.Name = "txtIntervaloFim";
			this.txtIntervaloFim.Size = new System.Drawing.Size(72, 20);
			this.txtIntervaloFim.TabIndex = 5;
			this.txtIntervaloFim.Text = "999";
			this.txtIntervaloFim.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtIntervaloFim.Enter += new System.EventHandler(this.txtIntervaloFim_Enter);
			this.txtIntervaloFim.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIntervaloFim_KeyDown);
			this.txtIntervaloFim.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIntervaloFim_KeyPress);
			this.txtIntervaloFim.Leave += new System.EventHandler(this.txtIntervaloFim_Leave);
			// 
			// txtIntervaloInicio
			// 
			this.txtIntervaloInicio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtIntervaloInicio.Location = new System.Drawing.Point(106, 72);
			this.txtIntervaloInicio.Name = "txtIntervaloInicio";
			this.txtIntervaloInicio.Size = new System.Drawing.Size(72, 20);
			this.txtIntervaloInicio.TabIndex = 3;
			this.txtIntervaloInicio.Text = "999";
			this.txtIntervaloInicio.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtIntervaloInicio.Enter += new System.EventHandler(this.txtIntervaloInicio_Enter);
			this.txtIntervaloInicio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIntervaloInicio_KeyDown);
			this.txtIntervaloInicio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIntervaloInicio_KeyPress);
			this.txtIntervaloInicio.Leave += new System.EventHandler(this.txtIntervaloInicio_Leave);
			// 
			// txtVolumeUnico
			// 
			this.txtVolumeUnico.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.txtVolumeUnico.Location = new System.Drawing.Point(106, 28);
			this.txtVolumeUnico.Name = "txtVolumeUnico";
			this.txtVolumeUnico.Size = new System.Drawing.Size(72, 20);
			this.txtVolumeUnico.TabIndex = 1;
			this.txtVolumeUnico.Text = "999";
			this.txtVolumeUnico.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtVolumeUnico.Enter += new System.EventHandler(this.txtVolumeUnico_Enter);
			this.txtVolumeUnico.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtVolumeUnico_KeyDown);
			this.txtVolumeUnico.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtVolumeUnico_KeyPress);
			this.txtVolumeUnico.Leave += new System.EventHandler(this.txtVolumeUnico_Leave);
			// 
			// rbVolumeRange
			// 
			this.rbVolumeRange.AutoSize = true;
			this.rbVolumeRange.Location = new System.Drawing.Point(19, 73);
			this.rbVolumeRange.Name = "rbVolumeRange";
			this.rbVolumeRange.Size = new System.Drawing.Size(81, 17);
			this.rbVolumeRange.TabIndex = 2;
			this.rbVolumeRange.TabStop = true;
			this.rbVolumeRange.Text = "Intervalo de";
			this.rbVolumeRange.UseVisualStyleBackColor = true;
			// 
			// rbVolumeUnico
			// 
			this.rbVolumeUnico.AutoSize = true;
			this.rbVolumeUnico.Location = new System.Drawing.Point(19, 29);
			this.rbVolumeUnico.Name = "rbVolumeUnico";
			this.rbVolumeUnico.Size = new System.Drawing.Size(75, 17);
			this.rbVolumeUnico.TabIndex = 0;
			this.rbVolumeUnico.TabStop = true;
			this.rbVolumeUnico.Text = "Nº Volume";
			this.rbVolumeUnico.UseVisualStyleBackColor = true;
			// 
			// FVolumeSeleciona
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btnCancela;
			this.ClientSize = new System.Drawing.Size(423, 273);
			this.ControlBox = false;
			this.Controls.Add(this.btnDummy);
			this.Controls.Add(this.gboxNf);
			this.Controls.Add(this.gboxOpcao);
			this.Controls.Add(this.btnCancela);
			this.Controls.Add(this.btnOk);
			this.Name = "FVolumeSeleciona";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Informe o Nº do Volume para Impressão da Etiqueta";
			this.Load += new System.EventHandler(this.FVolumeSeleciona_Load);
			this.Shown += new System.EventHandler(this.FVolumeSeleciona_Shown);
			this.gboxNf.ResumeLayout(false);
			this.gboxNf.PerformLayout();
			this.gboxOpcao.ResumeLayout(false);
			this.gboxOpcao.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnOk;
		private System.Windows.Forms.Button btnCancela;
		private System.Windows.Forms.Button btnDummy;
		private System.Windows.Forms.GroupBox gboxNf;
		private System.Windows.Forms.Label lblTitNf;
		private System.Windows.Forms.TextBox txtNF;
		private System.Windows.Forms.GroupBox gboxOpcao;
		private System.Windows.Forms.Label lblTitIntervaloAte;
		private System.Windows.Forms.TextBox txtIntervaloFim;
		private System.Windows.Forms.TextBox txtIntervaloInicio;
		private System.Windows.Forms.TextBox txtVolumeUnico;
		private System.Windows.Forms.RadioButton rbVolumeRange;
		private System.Windows.Forms.RadioButton rbVolumeUnico;
	}
}