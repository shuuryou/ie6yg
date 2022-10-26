using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace yomigaeri_backend
{
	public partial class BrowserForm : Form
	{
		public BrowserForm()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			RDPVirtualChannel.WriteChannel("INVISIB");
			Application.Exit();
		}

		private void BrowserForm_Load(object sender, EventArgs e)
		{
			RDPVirtualChannel.WriteChannel("VISIBLE");
		}

		private void button2_Click(object sender, EventArgs e)
		{
			Process.Start("C:\\windows\\system32\\cmd.exe");
		}
	}
}
