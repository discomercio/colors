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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FImportXML));
            this.lblTitulo = new System.Windows.Forms.Label();
            this.gboxMsgErro = new System.Windows.Forms.GroupBox();
            this.lbErro = new System.Windows.Forms.ListBox();
            this.gboxMensagensInformativas = new System.Windows.Forms.GroupBox();
            this.lbMensagem = new System.Windows.Forms.ListBox();
            this.btnAtualizaDatas = new System.Windows.Forms.Button();
            this.pnBotoes.SuspendLayout();
            this.pnCampos.SuspendLayout();
            this.gboxMsgErro.SuspendLayout();
            this.gboxMensagensInformativas.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnBotoes
            // 
            this.pnBotoes.Controls.Add(this.btnAtualizaDatas);
            this.pnBotoes.Controls.SetChildIndex(this.btnFechar, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnDummy, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnSobre, 0);
            this.pnBotoes.Controls.SetChildIndex(this.btnAtualizaDatas, 0);
            // 
            // btnDummy
            // 
            this.btnDummy.Location = new System.Drawing.Point(375, -200);
            // 
            // pnCampos
            // 
            this.pnCampos.Controls.Add(this.gboxMsgErro);
            this.pnCampos.Controls.Add(this.gboxMensagensInformativas);
            this.pnCampos.Controls.Add(this.lblTitulo);
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
            this.lblTitulo.Size = new System.Drawing.Size(1004, 40);
            this.lblTitulo.TabIndex = 1;
            this.lblTitulo.Text = "Atualizar Datas de Importação de Arquivos XML";
            this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // gboxMsgErro
            // 
            this.gboxMsgErro.Controls.Add(this.lbErro);
            this.gboxMsgErro.Location = new System.Drawing.Point(3, 362);
            this.gboxMsgErro.Name = "gboxMsgErro";
            this.gboxMsgErro.Size = new System.Drawing.Size(994, 95);
            this.gboxMsgErro.TabIndex = 9;
            this.gboxMsgErro.TabStop = false;
            this.gboxMsgErro.Text = "Mensagens de Erro";
            // 
            // lbErro
            // 
            this.lbErro.ForeColor = System.Drawing.Color.Red;
            this.lbErro.FormattingEnabled = true;
            this.lbErro.Location = new System.Drawing.Point(15, 19);
            this.lbErro.Name = "lbErro";
            this.lbErro.ScrollAlwaysVisible = true;
            this.lbErro.Size = new System.Drawing.Size(965, 69);
            this.lbErro.TabIndex = 0;
            // 
            // gboxMensagensInformativas
            // 
            this.gboxMensagensInformativas.Controls.Add(this.lbMensagem);
            this.gboxMensagensInformativas.Location = new System.Drawing.Point(3, 261);
            this.gboxMensagensInformativas.Name = "gboxMensagensInformativas";
            this.gboxMensagensInformativas.Size = new System.Drawing.Size(994, 95);
            this.gboxMensagensInformativas.TabIndex = 8;
            this.gboxMensagensInformativas.TabStop = false;
            this.gboxMensagensInformativas.Text = "Mensagens Informativas";
            // 
            // lbMensagem
            // 
            this.lbMensagem.FormattingEnabled = true;
            this.lbMensagem.Location = new System.Drawing.Point(15, 19);
            this.lbMensagem.Name = "lbMensagem";
            this.lbMensagem.ScrollAlwaysVisible = true;
            this.lbMensagem.Size = new System.Drawing.Size(965, 69);
            this.lbMensagem.TabIndex = 0;
            // 
            // btnAtualizaDatas
            // 
            this.btnAtualizaDatas.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnAtualizaDatas.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizaDatas.Image")));
            this.btnAtualizaDatas.Location = new System.Drawing.Point(869, 4);
            this.btnAtualizaDatas.Name = "btnAtualizaDatas";
            this.btnAtualizaDatas.Size = new System.Drawing.Size(40, 44);
            this.btnAtualizaDatas.TabIndex = 8;
            this.btnAtualizaDatas.TabStop = false;
            this.btnAtualizaDatas.UseVisualStyleBackColor = true;
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
            this.pnCampos.ResumeLayout(false);
            this.gboxMsgErro.ResumeLayout(false);
            this.gboxMensagensInformativas.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.GroupBox gboxMsgErro;
        private System.Windows.Forms.ListBox lbErro;
        private System.Windows.Forms.GroupBox gboxMensagensInformativas;
        private System.Windows.Forms.ListBox lbMensagem;
        private System.Windows.Forms.Button btnAtualizaDatas;
    }
}
