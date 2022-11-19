using System.ServiceProcess;

namespace yomigaeri_server
{
	public partial class YGSERVICE : ServiceBase
	{
		public YGSERVICE()
		{
			InitializeComponent();
		}

		protected override void OnStart(string[] args)
		{
		}

		protected override void OnStop()
		{
		}
	}
}
