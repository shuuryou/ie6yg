using Microsoft.Win32;
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
			Program.WebBrowser.Dock = DockStyle.Fill;
			Controls.Add(Program.WebBrowser);

			Program.WebBrowser.LoadingStateChanged += WebBrowser_LoadingStateChanged; 
			Program.WebBrowser.StatusMessage += WebBrowser_StatusMessage;
			Program.WebBrowser.TitleChanged += WebBrowser_TitleChanged;
			Program.WebBrowser.AddressChanged += WebBrowser_AddressChanged;
		}

		private void WebBrowser_LoadingStateChanged(object sender, CefSharp.LoadingStateChangedEventArgs e)
		{
			throw new NotImplementedException();
		}

		private void WebBrowser_StatusMessage(object sender, CefSharp.StatusMessageEventArgs e)
		{
			throw new NotImplementedException();
		}

		private void WebBrowser_TitleChanged(object sender, CefSharp.TitleChangedEventArgs e)
		{
			throw new NotImplementedException();
		}

		private void WebBrowser_AddressChanged(object sender, CefSharp.AddressChangedEventArgs e)
		{
			throw new NotImplementedException();
		}

		private void BrowserForm_Load(object sender, EventArgs e)
		{
			RDPVirtualChannel.WriteChannel("VISIBLE");
		}

		private void VirtualChannelTimer_Tick(object sender, EventArgs e)
		{
			// This is not the nicest way but it works until something
			// better comes along.

			string message = RDPVirtualChannel.ReadChannel(true);
			if (string.IsNullOrEmpty(message))
				return;

			Logging.WriteLineToLog("VirtualChannelTimer: Get message: \"{0}\"", message);

			if (message.ToUpperInvariant().StartsWith("NAVIGATE:"))
			{
				Program.WebBrowser.LoadUrl(message.Substring(9));
				return;
			}
		}

		private void BrowserForm_ResizeBegin(object sender, EventArgs e)
		{
			SuspendLayout();
		}

		private void BrowserForm_ResizeEnd(object sender, EventArgs e)
		{
			ResumeLayout();
		}
	}
}
