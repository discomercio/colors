using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration;
using System.Configuration.Install;
using System.Reflection;


namespace FinanceiroService
{
	[RunInstaller(true)]
	public partial class FinanceiroProjectInstaller : Installer
	{
		public FinanceiroProjectInstaller()
		{
			InitializeComponent();
		}

		private string GetConfigurationValue(string key)
		{
			Assembly service = Assembly.GetAssembly(typeof(FinanceiroProjectInstaller));
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
			this.FinanceiroServiceInstaller.ServiceName = GetConfigurationValue("ServiceName");
			this.FinanceiroServiceInstaller.DisplayName = GetConfigurationValue("DisplayName");
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
