using CefSharp;
using CefSharp.WinForms;
using Microsoft.Win32;
using System;
using System.Globalization;
using System.Windows.Forms;

namespace yomigaeri_backend.UI
{
	public partial class BrowserForm : Form
	{
		private SynchronizerState m_SyncState;
		private SynchronizerState? m_PrevSyncState;

		public BrowserForm()
		{
			InitializeComponent();

			m_SyncState = new SynchronizerState();
			m_PrevSyncState = null;

			Program.WebBrowser.Dock = DockStyle.Fill;
			Controls.Add(Program.WebBrowser);

			Program.WebBrowser.IsBrowserInitializedChanged += WebBrowser_IsBrowserInitializedChanged;
			Program.WebBrowser.LoadingStateChanged += WebBrowser_LoadingStateChanged;
			Program.WebBrowser.StatusMessage += WebBrowser_StatusMessage;
			Program.WebBrowser.TitleChanged += WebBrowser_TitleChanged;
			Program.WebBrowser.AddressChanged += WebBrowser_AddressChanged;
			Program.WebBrowser.DisplayHandler = new MyDisplayHandler(m_SyncState, SynchronizeWithFrontend);

			SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
		}

		private void WebBrowser_IsBrowserInitializedChanged(object sender, EventArgs e)
		{
			var b = ((ChromiumWebBrowser)sender);

			this.InvokeOnUiThreadIfRequired(() => b.Focus());
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
				m_PrevSyncState = null;
				StartRDPProcessing();
				return;
			}
		}

		private void WebBrowser_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
		{
			m_SyncState.CanGoBack = e.CanGoBack;
			m_SyncState.CanGoForward = e.CanGoForward;
			m_SyncState.CanReload = e.CanReload;
			m_SyncState.IsLoading = e.IsLoading;

			SynchronizeWithFrontend();
		}

		private void WebBrowser_StatusMessage(object sender, StatusMessageEventArgs e)
		{
			m_SyncState.StatusText = e.Value;
			SynchronizeWithFrontend();
		}

		private void WebBrowser_TitleChanged(object sender, TitleChangedEventArgs e)
		{
			m_SyncState.PageTitle = e.Title;
			SynchronizeWithFrontend();
		}

		private void WebBrowser_AddressChanged(object sender, AddressChangedEventArgs e)
		{
			m_SyncState.Address = e.Address;
			SynchronizeWithFrontend();
		}


		private void VirtualChannelTimer_Tick(object sender, EventArgs e)
		{
			// This is not the nicest way but it works until something
			// better comes along.
			string message = null;

			try
			{
				message = RDPVirtualChannel.ReadChannel(true);
			}
			catch (Exception ex)
			{
				Logging.WriteLineToLog("VirtualChannelTimer: Failed in ReadChannel: {0}", ex.ToString());
				StopRDPProcessing();
				// XXX should probably crash here or at least show a message?
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
			m_SyncState.Visible = false;
			m_PrevSyncState = null;
			RDPVirtualChannel.Reset();
		}

		private void StartRDPProcessing()
		{
			RDPVirtualChannel.OpenChannel(); // In case frontend is reconnecting, otherwise a no-op.
			m_SyncState.Visible = true;
			VirtualChannelTimer.Enabled = true;
			SynchronizeWithFrontend();
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