using CefSharp;
using CefSharp.WinForms;
using Microsoft.Win32;
using System;
using System.Globalization;
using System.Web;
using System.Windows.Forms;
using yomigaeri_backend.Browser.Handlers;

namespace yomigaeri_backend.Browser
{
	public partial class BrowserForm : Form
	{
		private readonly SynchronizerState m_SyncState;
		private readonly TravelLog m_TravelLog;

		public BrowserForm()
		{
			InitializeComponent();

			SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;

			m_SyncState = new SynchronizerState();
			m_TravelLog = new TravelLog();

			m_TravelLog.NewLogPrepared += TravelLog_NewLogPrepared;

			Program.WebBrowser.Dock = DockStyle.Fill;
			Controls.Add(Program.WebBrowser);

			Program.WebBrowser.StatusMessage += WebBrowser_StatusMessage;
			
			Program.WebBrowser.IsBrowserInitializedChanged += WebBrowser_IsBrowserInitializedChanged;

			Program.WebBrowser.LoadHandler = new MyLoadHandler(m_TravelLog, m_SyncState, SynchronizeWithFrontend);
			Program.WebBrowser.DisplayHandler = new MyDisplayHandler(m_SyncState, SynchronizeWithFrontend);
			Program.WebBrowser.RequestHandler = new MyRequestHandler(m_SyncState, SynchronizeWithFrontend);

			// This is important! See Handlers/DisplayHandler.cs, OnCursorChange method
			Cursor.Hide();
		}

		private void WebBrowser_StatusMessage(object sender, StatusMessageEventArgs e)
		{
			// DisplayHandler.OnStatusMessage never gets called for some reason.
			
			if (e.Browser.IsLoading)
				return; // It keeps trying to clear it while its displaying "Loading..."

			m_SyncState.StatusText = e.Value;
			SynchronizeWithFrontend();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
				Program.WebBrowser.StatusMessage -= WebBrowser_StatusMessage;
				Program.WebBrowser.IsBrowserInitializedChanged -= WebBrowser_IsBrowserInitializedChanged;

				components.Dispose();
				Program.WebBrowser.Dispose();
			}
			base.Dispose(disposing);
		}


		private void VirtualChannelTimer_Tick(object sender, EventArgs e)
		{
			// TODO This is not the nicest way but it works until something better comes along.

			string message = null;

			try
			{
				message = RDPVirtualChannel.Read(true);
			}
			catch (Exception ex)
			{
				Logging.WriteLineToLog("BrowserForm: VirtualChannelTimer: Failed in ReadChannel: {0}", ex.ToString());

				VirtualChannelTimer.Enabled = false;

				string error_html = Resources.BACKENDERROR_HTML;

				error_html = error_html.Replace("%TITLE_TEXT%", HttpUtility.HtmlEncode(Resources.E_BackendVirtualChannel_Title));
				error_html = error_html.Replace("%INTRODUCTION_TEXT%", HttpUtility.HtmlEncode(Resources.E_BackendVirtualChannel_Text));
				error_html = error_html.Replace("%ERROR_TEXT%", HttpUtility.HtmlEncode(ex.Message));

				Program.WebBrowser.SetMainFrameDocumentContentAsync(error_html);
				return;
			}
			
			if (string.IsNullOrEmpty(message))
				return;

			Logging.WriteLineToLog("BrowserForm: VirtualChannelTimer: Get message: \"{0}\"", message);

			ProcessMessage(message);
		}

		private void WebBrowser_IsBrowserInitializedChanged(object sender, EventArgs e)
		{
			var b = ((ChromiumWebBrowser)sender);

			this.InvokeOnUiThreadIfRequired(() => b.Focus());
		}

		private void BrowserForm_Shown(object sender, EventArgs e)
		{
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
		private void TravelLog_NewLogPrepared(object sender, EventArgs e)
		{
			m_SyncState.TravelLog = true;
			SynchronizeWithFrontend();
		}

		private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
		{
			Logging.WriteLineToLog("BrowserForm: SystemEvents_SessionSwitch: Reason: {0}", e.Reason);

			switch (e.Reason)
			{
				case SessionSwitchReason.SessionLogoff:
				case SessionSwitchReason.ConsoleDisconnect:
				case SessionSwitchReason.RemoteDisconnect:
					VirtualChannelTimer.Enabled = false;
					Application.Exit();
					return;
			}
		}

		private void SynchronizeWithFrontend()
		{
			#region Visible
			if (m_SyncState.IsChanged(SynchronizerState.Change.Visible))
			{
				if (m_SyncState.Visible)
					RDPVirtualChannel.Write("VISIBLE");
				else
					RDPVirtualChannel.Write("INVISIB");
			}
			#endregion

			#region CanGoBack
			if (m_SyncState.IsChanged(SynchronizerState.Change.CanGoBack))
			{
				if (m_SyncState.CanGoBack)
					RDPVirtualChannel.Write("BBACKON");
				else
					RDPVirtualChannel.Write("BBACKOF");
			}
			#endregion

			#region CanGoForward
			if (m_SyncState.IsChanged(SynchronizerState.Change.CanGoForward))
			{
				if (m_SyncState.CanGoForward)
					RDPVirtualChannel.Write("BFORWON");
				else
					RDPVirtualChannel.Write("BFORWOF");
			}
			#endregion

			#region CanReload
			if (m_SyncState.IsChanged(SynchronizerState.Change.CanReload))
			{
				if (m_SyncState.CanReload)
				{
					RDPVirtualChannel.Write("BREFRON");
				}
				else
				{
					RDPVirtualChannel.Write("BREFROF");
				}
			}
			#endregion

			#region IsLoading
			if (m_SyncState.IsChanged(SynchronizerState.Change.IsLoading))
			{
				if (m_SyncState.IsLoading)
				{
					RDPVirtualChannel.Write("BSTOPON");
				}
				else
				{
					RDPVirtualChannel.Write("BSTOPOF");
					RDPVirtualChannel.Write("PROGRES0");
				}
			}
			#endregion

			#region AddHistoryItem
			if (m_SyncState.IsChanged(SynchronizerState.Change.AddHistoryItem))
			{
				if (!string.IsNullOrEmpty(m_SyncState.Address))
					RDPVirtualChannel.Write("ADDHIST" +
						m_SyncState.PageTitle + '\x1' + m_SyncState.Address);

				m_SyncState.AddHistoryItem = false;
			}
			#endregion

			#region StatusProgress
			if (m_SyncState.IsChanged(SynchronizerState.Change.StatusProgress))
			{
				int progress = m_SyncState.StatusProgress;

				if (progress < 0)
					progress = 0;

				if (progress > 100)
					progress = 100;

				RDPVirtualChannel.Write("PROGRES" + progress.ToString(CultureInfo.InvariantCulture));
			}
			#endregion

			#region Address
			if (m_SyncState.IsChanged(SynchronizerState.Change.Address))
			{
				RDPVirtualChannel.Write("ADDRESS" + m_SyncState.Address);
			}
			#endregion

			#region StatusText
			if (m_SyncState.IsChanged(SynchronizerState.Change.StatusText))
			{
				RDPVirtualChannel.Write("STATUST" + m_SyncState.StatusText);
			}
			#endregion

			#region PageTitle
			if (m_SyncState.IsChanged(SynchronizerState.Change.PageTitle))
			{
				RDPVirtualChannel.Write("PGTITLE" + m_SyncState.PageTitle);
			}
			#endregion

			#region TravelLog
			if (m_SyncState.IsChanged(SynchronizerState.Change.TravelLog))
			{
				RDPVirtualChannel.Write("TRAVLBK" +
					m_TravelLog.MakeMenuStringForFrontend(TravelLog.TravelDirection.Back));

				RDPVirtualChannel.Write("TRAVLFW" +
					m_TravelLog.MakeMenuStringForFrontend(TravelLog.TravelDirection.Forward));

				m_SyncState.TravelLog = false;
			}
			#endregion

			#region Cursor
			if (m_SyncState.IsChanged(SynchronizerState.Change.Cursor))
			{
				RDPVirtualChannel.Write("SETCURS" + m_SyncState.Cursor);
			}
			#endregion

			#region Tooltip
			if (m_SyncState.IsChanged(SynchronizerState.Change.Tooltip))
			{
				RDPVirtualChannel.Write("TOOLTIP" + m_SyncState.Tooltip);
			}
			#endregion

			#region SSL Icon
			if (m_SyncState.IsChanged(SynchronizerState.Change.SSLIcon))
			{
				switch (m_SyncState.SSLIcon)
				{
					case SynchronizerState.SSLIconState.None:
						RDPVirtualChannel.Write("SSLICOF");
						break;
					case SynchronizerState.SSLIconState.Secure:
						RDPVirtualChannel.Write("SSLICON");
						break;
					case SynchronizerState.SSLIconState.SecureBadCert:
						RDPVirtualChannel.Write("SSLICBD");
						break;
				}
			}
			#endregion

			m_SyncState.SyncNone();
		}

		private void ProcessMessage(string message)
		{
			if (message == null)
				throw new ArgumentNullException("message");

			// There is no technical reason that the messages start with
			// seven letter long commands; they could be longer. It just
			// emerged as a pattern and now it's here to stay like that.

			#region NAVIGAT -- Navigate to URL
			if (message.StartsWith("NAVIGAT", StringComparison.Ordinal))
			{
				string url = message.Substring(7);

				if (string.IsNullOrEmpty(url))
					return;

				Program.WebBrowser.LoadUrl(url);
				return;
			}
			#endregion

			#region WINSIZE -- Adjust window size
			if (message.StartsWith("WINSIZE", StringComparison.Ordinal))
			{
				// e.g. WINSIZE800,600
				//      01234567890123...

				string size = message.Substring(7);

				if (String.IsNullOrEmpty(size))
					goto winsize_err;

				int idx = size.IndexOf(',');

				if (idx == -1 || idx + 1 > size.Length)
					goto winsize_err;

				if (int.TryParse(size.Substring(0, idx), NumberStyles.None, CultureInfo.InvariantCulture, out int width) &&
					int.TryParse(size.Substring(idx + 1), NumberStyles.None, CultureInfo.InvariantCulture, out int height))
				{
					this.Width = width;
					this.Height = height;
					return;
				}

			winsize_err:
				Logging.WriteLineToLog("BrowserForm: ProcessMessage: Illegal WINSIZE message receievd: \"{0}\"", message);

				return;
			}
			#endregion

			#region BTNBACK, BTNFORW -- Forward and Back buttons
			if (message == "BTNBACK")
			{
				if (!Program.WebBrowser.CanGoBack)
					return;

				Program.WebBrowser.Back();

				return;
			}

			if (message == "BTNFORW")
			{
				if (!Program.WebBrowser.CanGoForward)
					return;

				Program.WebBrowser.Forward();
				return;
			}
			#endregion

			#region MNUBACK, MNUFORW -- Forward and Back travel log
			if (message.StartsWith("MNUBACK") || message.StartsWith("MNUFORW"))
			{
				if (message.Length != 8)
					return;

				// e.g. MNUBACK1
				//      01234567

				if (!int.TryParse(message[7].ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out int offset))
					return;

				if (message.StartsWith("MNUBACK"))
					offset *= -1;

				// TODO - This needs to be done another way at some point. Right now
				// CEF doesn't offer another way, but websites that fuck around with
				// the HTML5 history API will probably break this easily.
				Program.WebBrowser.ExecuteScriptAsync("window.history.go", offset);
			}
			#endregion
		}
	}
}