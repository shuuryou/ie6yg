using CefSharp;
using CefSharp.WinForms;
using Microsoft.Win32;
using System;
using System.Windows.Forms;

namespace yomigaeri_backend.UI
{
	public partial class BrowserForm : Form
	{
		private readonly SynchronizerState m_SyncState;
		private readonly TravelLog m_TravelLog;

		public BrowserForm()
		{
			InitializeComponent();

			m_SyncState = new SynchronizerState();
			m_TravelLog = new TravelLog();

			m_TravelLog.NewHistoryReady += HistoryProcessor_NewHistoryReady;
			Program.WebBrowser.Dock = DockStyle.Fill;
			Controls.Add(Program.WebBrowser);
			
			Program.WebBrowser.IsBrowserInitializedChanged += WebBrowser_IsBrowserInitializedChanged;
			Program.WebBrowser.LoadingStateChanged += WebBrowser_LoadingStateChanged;
			Program.WebBrowser.FrameLoadStart += WebBrowser_FrameLoadStart;
			Program.WebBrowser.FrameLoadEnd += WebBrowser_FrameLoadEnd;

			Program.WebBrowser.DisplayHandler = new MyDisplayHandler(m_SyncState, SynchronizeWithFrontend);
			Program.WebBrowser.RequestHandler = new MyRequestHandler(m_SyncState, SynchronizeWithFrontend);

			// This is important! See BrowserForm.DisplayHandler.cs, OnCursorChange method
			Cursor.Hide();

			SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
		}


		private void HistoryProcessor_NewHistoryReady(object sender, EventArgs e)
		{
			m_SyncState.TravelLog = true;
			SynchronizeWithFrontend();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
				SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
				// todo unsubscribe all
			}
			base.Dispose(disposing);
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

			ProcessMessage(message);
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
		private void StopRDPProcessing()
		{
			VirtualChannelTimer.Enabled = false;
			m_SyncState.Visible = false;
			RDPVirtualChannel.Reset();
		}

		private void StartRDPProcessing()
		{
			RDPVirtualChannel.OpenChannel(); // In case frontend is reconnecting, otherwise a no-op.
			m_SyncState.Visible = true;
			VirtualChannelTimer.Enabled = true;
			SynchronizeWithFrontend();
		}
	}
}