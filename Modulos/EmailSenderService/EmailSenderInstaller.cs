using System;
using System.Collections;
using System.ComponentModel;
using System.Reflection;
using System.Configuration;

namespace EmailSenderService
{
	[RunInstaller(true)]
	public partial class EmailSenderInstaller : System.Configuration.Install.Installer
	{
		public EmailSenderInstaller()
		{
			InitializeComponent();
		}

		private string GetConfigurationValue(string key)
		{
			Assembly service = Assembly.GetAssembly(typeof(EmailSenderInstaller));
			Configuration config = ConfigurationManager.OpenExeConfiguration(service.Location);
			if (config.AppSettings.Settings[key] != null)
			{
				return config.AppSettings.Settings[key].Value;
			}
			else
			{
				throw new IndexOutOfRangeException("Settings collection does not contain the requested key:" + key);
			}
		}

		private void SetServiceName()
		{
			this.EmailSenderServiceInstaller.ServiceName = GetConfigurationValue("ServiceName");
			this.EmailSenderServiceInstaller.DisplayName = GetConfigurationValue("DisplayName");
		}

		protected override void OnBeforeInstall(IDictionary savedState)
		{
			SetServiceName();
			base.OnBeforeInstall(savedState);
		}

		protected override void OnBeforeUninstall(IDictionary savedState)
		{
			SetServiceName();
			base.OnBeforeUninstall(savedState);
		}
	}
}
