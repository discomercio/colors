namespace EmailSenderService
{
    partial class EmailSenderInstaller
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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.EmailSenderServiceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.EmailSenderServiceInstaller = new System.ServiceProcess.ServiceInstaller();
            // 
            // EmailSenderServiceProcessInstaller
            // 
            this.EmailSenderServiceProcessInstaller.Password = null;
            this.EmailSenderServiceProcessInstaller.Username = null;
            // 
            // EmailSenderServiceInstaller
            // 
            this.EmailSenderServiceInstaller.Description = "Serviço para envio automático de e-mails";
            this.EmailSenderServiceInstaller.DisplayName = "EMailSenderService";
            this.EmailSenderServiceInstaller.ServiceName = "EmailSenderService";
            // 
            // EmailSenderInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.EmailSenderServiceProcessInstaller,
            this.EmailSenderServiceInstaller});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller EmailSenderServiceProcessInstaller;
        private System.ServiceProcess.ServiceInstaller EmailSenderServiceInstaller;
    }
}