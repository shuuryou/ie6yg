using CefSharp;
using CefSharp.WinForms;
using Microsoft.Win32;
using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace yomigaeri_backend
{
	public partial class BrowserForm : Form
	{
		public BrowserForm()
		{
			InitializeComponent();

			//this.BackColor = Color.Red;

			Program.WebBrowser.Dock = DockStyle.Fill;
			Controls.Add(Program.WebBrowser);

			Program.WebBrowser.IsBrowserInitializedChanged += WebBrowser_IsBrowserInitializedChanged;
			Program.WebBrowser.LoadingStateChanged += WebBrowser_LoadingStateChanged;
			Program.WebBrowser.StatusMessage += WebBrowser_StatusMessage;
			Program.WebBrowser.TitleChanged += WebBrowser_TitleChanged;
			Program.WebBrowser.AddressChanged += WebBrowser_AddressChanged;
			Program.WebBrowser.DisplayHandler = new MyTestDisplayHandler();

			SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
		}

		private void WebBrowser_IsBrowserInitializedChanged(object sender, EventArgs e)
		{
			var b = ((ChromiumWebBrowser)sender);

			this.InvokeOnUiThreadIfRequired(() => b.Focus());
		}

		private class MyTestDisplayHandler : CefSharp.Handler.DisplayHandler
		{
			protected override void OnLoadingProgressChange(IWebBrowser chromiumWebBrowser, IBrowser browser, double progress)
			{
				RDPVirtualChannel.WriteChannel("PROGRES" + (progress * 100).ToString(CultureInfo.InvariantCulture));
				base.OnLoadingProgressChange(chromiumWebBrowser, browser, progress);
			}
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
				SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
			}
			base.Dispose(disposing);
		}


		private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
		{
			Logging.WriteLineToLog("SystemEvents_SessionSwitch in BrowserForm. Reason: {0}", e.Reason);

			if (e.Reason == SessionSwitchReason.RemoteDisconnect)
			{
				StopRDPProcessing();
				return;
			}

			if (e.Reason == SessionSwitchReason.RemoteConnect)
			{
				this.WindowState = FormWindowState.Normal;
				this.WindowState = FormWindowState.Maximized;
				StartRDPProcessing();
				return;
			}
		}

		private void WebBrowser_LoadingStateChanged(object sender, CefSharp.LoadingStateChangedEventArgs e)
		{
			if (e.CanGoBack)
				RDPVirtualChannel.WriteChannel("BBACKON");
			else
				RDPVirtualChannel.WriteChannel("BBACKOF");

			if (e.CanGoForward)
				RDPVirtualChannel.WriteChannel("BFORWON");
			else
				RDPVirtualChannel.WriteChannel("BFORWOF");

			if (e.CanReload)
				RDPVirtualChannel.WriteChannel("BREFRON");
			else
				RDPVirtualChannel.WriteChannel("BREFROF");

			if (e.IsLoading)
				RDPVirtualChannel.WriteChannel("BSTOPON");
			else
				RDPVirtualChannel.WriteChannel("BSTOPOF");
		}

		private void WebBrowser_StatusMessage(object sender, CefSharp.StatusMessageEventArgs e)
		{
			RDPVirtualChannel.WriteChannel("STATUST" + e.Value);
		}

		private void WebBrowser_TitleChanged(object sender, CefSharp.TitleChangedEventArgs e)
		{
			RDPVirtualChannel.WriteChannel("PGTITLE" + e.Title);
		}

		private void WebBrowser_AddressChanged(object sender, CefSharp.AddressChangedEventArgs e)
		{
			RDPVirtualChannel.WriteChannel("ADDRESS" + e.Address);
		}


		private void VirtualChannelTimer_Tick(object sender, EventArgs e)
		{
			// This is not the nicest way but it works until something
			// better comes along.
			string message = null;

			try
			{
				 message = RDPVirtualChannel.ReadChannel(true);
			} catch (Exception ex)
			{
				Logging.WriteLineToLog("VirtualChannelTimer: Failed in ReadChannel: {0}", ex.ToString());
				StopRDPProcessing();
			}

			if (string.IsNullOrEmpty(message))
				return;

			Logging.WriteLineToLog("VirtualChannelTimer: Get message: \"{0}\"", message);


			if (message.StartsWith("NAVIGATE:", StringComparison.Ordinal))
			{
				Program.WebBrowser.LoadUrl(message.Substring(9));
				return;
			}
		}

		private void StopRDPProcessing()
		{
			VirtualChannelTimer.Enabled = false;
			RDPVirtualChannel.Reset();
		}

		private void StartRDPProcessing()
		{
			Logging.WriteLineToLog("Attempting to make window visible on frontend.");

			RDPVirtualChannel.OpenChannel(); // In case frontend is reconnecting, otherwise a no-op.

			RDPVirtualChannel.WriteChannel("VISIBLE");
			RDPVirtualChannel.WriteChannel("STATUSTReady.");

			VirtualChannelTimer.Enabled = true;
		}

		private void BrowserForm_ResizeBegin(object sender, EventArgs e)
		{
			SuspendLayout();
		}

		private void BrowserForm_ResizeEnd(object sender, EventArgs e)
		{
			ResumeLayout();
		}

		private void BrowserForm_Shown(object sender, EventArgs e)
		{
			StartRDPProcessing();
		}
	}
}
