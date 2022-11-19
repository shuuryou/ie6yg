using CefSharp;
using CefSharp.WinForms;
using Microsoft.Win32;
using System;
using System.Globalization;
using System.Web;
using System.Windows.Forms;
using yomigaeri_backend.Browser.Handlers;
using yomigaeri_shared;

namespace yomigaeri_backend.Browser
{
	public partial class BrowserForm : Form
	{
		#region Win32 FindReplaceFlags
		[Flags]
		private enum Win32FindReplaceFlags
		{
			FR_DOWN = 0x1,
			FR_WHOLEWORD = 0x2,
			FR_MATCHCASE = 0x4,
			FR_ENABLEHOOK = 0x100,
			FR_ENABLETEMPLATE = 0x200,
			FR_NOUPDOWN = 0x400,
			FR_NOMATCHCASE = 0x800,
			FR_NOWHOLEWORD = 0x1000,
			FR_ENABLETEMPLATEHANDLE = 0x2000,
			FR_HIDEUPDOWN = 0x4000,
			FR_HIDEMATCHCASE = 0x8000,
			FR_HIDEWHOLEWORD = 0x10000
		}
		#endregion

		private readonly SynchronizerState m_SyncState;
		private readonly TravelLog m_TravelLog;

		private bool m_ExitAfterDownload;

		public BrowserForm()
		{
			InitializeComponent();

			m_ExitAfterDownload = false;

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
			Program.WebBrowser.JsDialogHandler = new MyJsDialogHandler(m_SyncState, SynchronizeWithFrontend);
			Program.WebBrowser.DownloadHandler = new MyDownloadHandler(m_SyncState, SynchronizeWithFrontend);

			((MyDownloadHandler)Program.WebBrowser.DownloadHandler).AllDownloadsCompleted += DownloadHandler_AllDownloadsCompleted;

			// This is important! See Handlers/DisplayHandler.cs, OnCursorChange method
			Cursor.Hide();
		}

		private void WebBrowser_StatusMessage(object sender, StatusMessageEventArgs e)
		{
			// DisplayHandler.OnStatusMessage never gets called for some reason.

			if (e.Browser.IsLoading)
				return; // It keeps trying to clear it while page load is still in progress *sigh*

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

					// Only exit here if there are no downloads running. If there
					// are active downloads, the event handler will take care of
					// exiting once there are none left.

					if (((MyDownloadHandler)Program.WebBrowser.DownloadHandler).DownloadsInProgress)
						m_ExitAfterDownload = true;
					else
						Application.Exit();

					return;
			}
		}

		private void DownloadHandler_AllDownloadsCompleted(object sender, EventArgs e)
		{
			if (m_ExitAfterDownload)
				Application.Exit();
		}

		private void VirtualChannelTimer_Tick(object sender, EventArgs e)
		{
			// TODO This is not the nicest way but it works until something better comes along.

			string message;

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

		private void ExitTimer_Tick(object sender, EventArgs e)
		{
			// TODO This is not the nicest way but it works until something better comes along.

			if (!((MyDownloadHandler)Program.WebBrowser.DownloadHandler).DownloadsInProgress)
			{
				ExitTimer.Enabled = false;
				Application.Exit();
			}
		}

		private void SynchronizeWithFrontend()
		{
			// There is no technical reason that the messages start with
			// seven letter long commands; they could be longer. It just
			// emerged as a pattern and now it's here to stay like that.

			#region Visible
			if (m_SyncState.IsChanged(SynchronizerState.Change.Visible))
			{
				if (m_SyncState.Visible)
					RDPVirtualChannel.Write("VISIBLE TRUE");
				else
					RDPVirtualChannel.Write("VISIBLE FALSE");
			}
			#endregion

			#region Toolbar
			if (m_SyncState.IsChanged(SynchronizerState.Change.Toolbar))
			{
				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "TOOLBAR {0} {1}", "BACK",
					m_SyncState.ToolbarButtons.HasFlag(SynchronizerState.FrontendToolbarButtons.Back) ? "TRUE" : "FALSE"));

				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "TOOLBAR {0} {1}", "FORWARD",
					m_SyncState.ToolbarButtons.HasFlag(SynchronizerState.FrontendToolbarButtons.Forward) ? "TRUE" : "FALSE"));

				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "TOOLBAR {0} {1}", "HOME",
					m_SyncState.ToolbarButtons.HasFlag(SynchronizerState.FrontendToolbarButtons.Home) ? "TRUE" : "FALSE"));

				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "TOOLBAR {0} {1}", "MEDIA",
					m_SyncState.ToolbarButtons.HasFlag(SynchronizerState.FrontendToolbarButtons.Media) ? "TRUE" : "FALSE"));

				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "TOOLBAR {0} {1}", "REFRESH",
					m_SyncState.ToolbarButtons.HasFlag(SynchronizerState.FrontendToolbarButtons.Refresh) ? "TRUE" : "FALSE"));

				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "TOOLBAR {0} {1}", "STOP",
					m_SyncState.ToolbarButtons.HasFlag(SynchronizerState.FrontendToolbarButtons.Stop) ? "TRUE" : "FALSE"));
			}
			#endregion

			#region MenuBar
			if (m_SyncState.IsChanged(SynchronizerState.Change.MenuBar))
			{
				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "MENUSET {0} {1}", "STOP",
					m_SyncState.MenuBarItems.HasFlag(SynchronizerState.FrontendMenuBarItems.Stop) ? "TRUE" : "FALSE"));

				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "MENUSET {0} {1}", "REFRESH",
					m_SyncState.MenuBarItems.HasFlag(SynchronizerState.FrontendMenuBarItems.Refresh) ? "TRUE" : "FALSE"));

				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "MENUSET {0} {1}", "CUT",
					m_SyncState.MenuBarItems.HasFlag(SynchronizerState.FrontendMenuBarItems.Cut) ? "TRUE" : "FALSE"));

				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "MENUSET {0} {1}", "COPY",
					m_SyncState.MenuBarItems.HasFlag(SynchronizerState.FrontendMenuBarItems.Copy) ? "TRUE" : "FALSE"));

				RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture, "MENUSET {0} {1}", "PASTE",
					m_SyncState.MenuBarItems.HasFlag(SynchronizerState.FrontendMenuBarItems.Paste) ? "TRUE" : "FALSE"));
			}
			#endregion

			#region AddHistoryItem
			if (m_SyncState.IsChanged(SynchronizerState.Change.AddHistoryItem))
			{
				if (!string.IsNullOrEmpty(m_SyncState.Address))
					RDPVirtualChannel.Write("ADDHIST " +
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

				RDPVirtualChannel.Write("PROGRES " + progress.ToString(CultureInfo.InvariantCulture));
			}
			#endregion

			#region Address
			if (m_SyncState.IsChanged(SynchronizerState.Change.Address))
			{
				RDPVirtualChannel.Write("ADDRESS " + m_SyncState.Address);
			}
			#endregion

			#region StatusText
			if (m_SyncState.IsChanged(SynchronizerState.Change.StatusText))
			{
				RDPVirtualChannel.Write("STATUST " + m_SyncState.StatusText);
			}
			#endregion

			#region PageTitle
			if (m_SyncState.IsChanged(SynchronizerState.Change.PageTitle))
			{
				RDPVirtualChannel.Write("PGTITLE " + m_SyncState.PageTitle);
			}
			#endregion

			#region TravelLog
			if (m_SyncState.IsChanged(SynchronizerState.Change.TravelLog))
			{
				RDPVirtualChannel.Write("TRAVLBK " +
					m_TravelLog.MakeMenuStringForFrontend(TravelLog.TravelDirection.Back));

				RDPVirtualChannel.Write("TRAVLFW " +
					m_TravelLog.MakeMenuStringForFrontend(TravelLog.TravelDirection.Forward));

				m_SyncState.TravelLog = false;
			}
			#endregion

			#region Cursor
			if (m_SyncState.IsChanged(SynchronizerState.Change.Cursor))
			{
				RDPVirtualChannel.Write("SETCURS " + m_SyncState.Cursor);
			}
			#endregion

			#region Tooltip
			if (m_SyncState.IsChanged(SynchronizerState.Change.Tooltip))
			{
				RDPVirtualChannel.Write("TOOLTIP " + m_SyncState.Tooltip);
			}
			#endregion

			#region SSLIcon
			if (m_SyncState.IsChanged(SynchronizerState.Change.SSLIcon))
			{
				switch (m_SyncState.SSLIcon)
				{
					case SynchronizerState.SSLIconState.None:
						RDPVirtualChannel.Write("SSLICON OFF");
						break;
					case SynchronizerState.SSLIconState.Secure:
						RDPVirtualChannel.Write("SSLICON OK");
						break;
					case SynchronizerState.SSLIconState.SecureBadCert:
						RDPVirtualChannel.Write("SSLICON BAD");
						break;
				}
			}
			#endregion

			#region CertificateState
			if (m_SyncState.IsChanged(SynchronizerState.Change.CertificateState))
			{
				RDPVirtualChannel.Write(string.Format("CERSTAT {0}", (int)m_SyncState.CertificateState));
			}
			#endregion

			#region CertificateData
			if (m_SyncState.IsChanged(SynchronizerState.Change.CertificateData))
			{
				if (m_SyncState.CertificateData == null)
					RDPVirtualChannel.Write("CERDATA");
				else
					RDPVirtualChannel.Write("CERDATA " + Convert.ToBase64String(m_SyncState.CertificateData));
			}
			#endregion

			#region CertificatePrompt
			if (m_SyncState.IsChanged(SynchronizerState.Change.CertificatePrompt))
			{
				RDPVirtualChannel.Write("CERSHOW");
			}
			#endregion

			#region JSDialogPrompt
			if (m_SyncState.IsChanged(SynchronizerState.Change.JSDialogPrompt) && m_SyncState.JSDialogPrompt != null)
			{
				switch (m_SyncState.JSDialogPrompt.Type)
				{
					case SynchronizerState.FrontendJSDialogs.Alert:
						RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture,
							"JSDIALG ALERT {0}", m_SyncState.JSDialogPrompt.Prompt));
						break;
					case SynchronizerState.FrontendJSDialogs.Confirm:
						RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture,
							"JSDIALG CONFIRM {0}", m_SyncState.JSDialogPrompt.Prompt));
						break;
					case SynchronizerState.FrontendJSDialogs.Prompt:
						// Strings limited to 32767 because it ought to be enough.
						// Anything longer will probably cause issues somewhere in
						// the frontend due to the use of VB6.

						string prompt = m_SyncState.JSDialogPrompt.Prompt;
						if (prompt.Length >= 32767)
							prompt = prompt.Substring(0, 32767);

						string default_text = m_SyncState.JSDialogPrompt.DefaultText;
						if (default_text.Length >= 32767)
							default_text = default_text.Substring(0, 32767);

						RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture,
							"JSDIALG PROMPT {0:D8}{1}{2:D8}{3}", prompt.Length, prompt,
							default_text.Length, default_text));

						break;
					case SynchronizerState.FrontendJSDialogs.OnBeforeUnload:
						RDPVirtualChannel.Write(string.Format(CultureInfo.InvariantCulture,
							"JSDIALG ONBEFOREUNLOAD {0}", m_SyncState.JSDialogPrompt.Prompt));
						break;
					default:
						throw new InvalidOperationException("Unsupported JS dialog type.");
				}
			}
			#endregion

			#region DownloadStart
			if (m_SyncState.IsChanged(SynchronizerState.Change.CertificateData))
			{
				RDPVirtualChannel.Write("DWNLOAD " + m_SyncState.DownloadStart);
			}
			#endregion

			m_SyncState.SyncNone();
		}

		private void ProcessMessage(string message)
		{
			if (message == null)
				throw new ArgumentNullException("message");

			#region NAVIGATE -- Navigate to URL
			if (message.StartsWith("NAVIGATE ", StringComparison.Ordinal))
			{
				string url = message.Substring(8);

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

			#region CERTCALLBACK -- Response to OnCertificateError
			if (message.StartsWith("CERTCALLBACK ", StringComparison.Ordinal))
			{
				// CERTCALLBACK CONTINUE
				// CERTCALLBACK CANCEL

				string response = message.Substring(13).ToUpperInvariant();

				IRequestCallback cb = ((MyRequestHandler)Program.WebBrowser.RequestHandler).SSLCertificate_CurrentErrorCallback;

				if (cb == null || cb.IsDisposed)
					return;

				if (response == "CONTINUE")
				{
					cb.Continue(true);
					return;
				}

				cb.Cancel();
			}
			#endregion

			#region JSCALLBACK - Response to a JavaScript dialog
			if (message.StartsWith("JSCALLBACK ", StringComparison.Ordinal))
			{
				// All except prompt(): JSCALLBACK OK
				// All:                 JSCALLBACK CANCEL
				// prompt():            JSCALLBACK OK abcdef...
				string response = message.Substring(11);
				string userInput = string.Empty;

				{
					// Deal with the input supplied by prompt(), if necessary

					int idx = response.IndexOf(' ');

					if (idx != -1 && idx + 1 <= response.Length)
					{
						userInput = response.Substring(idx + 1);
						response = response.Substring(0, idx);
					}
				}

				IJsDialogCallback cb = ((MyJsDialogHandler)Program.WebBrowser.JsDialogHandler).JSDialog_Callback;

				if (cb == null || cb.IsDisposed)
					return;

				bool success = (response.ToUpperInvariant() == "OK");

				cb.Continue(success, userInput);
			}
			#endregion

			#region FIND -- Find in page
			if (message.StartsWith("FIND ", StringComparison.Ordinal))
			{
				// FIND [NewFind=1 or 0] [Flags=00000000] [Text]
				// NewFind: Whether to start a new search or continue an old one
				// Flags:   See Win32FindReplaceFlags
				// Text:    Text to find
				// Sample:
				// FIND 1 00000000 abcdef...

				if (message == "FIND END")
				{
					Program.WebBrowser.StopFinding(true);
					return;
				}

				if (message.Length < 17)
				{
					Logging.WriteLineToLog("BrowserForm: ProcessMessage: FIND command has bad format.");
					return;
				}

				bool new_find = (message[5] == '1');

				if (!int.TryParse(message.Substring(7, 8), NumberStyles.None, CultureInfo.InvariantCulture, out int flags_int))
				{
					Logging.WriteLineToLog("BrowserForm: ProcessMessage: FIND command has bad flags format.");
					return;
				}

				Win32FindReplaceFlags flags = (Win32FindReplaceFlags)flags_int;

				Program.WebBrowser.Find(message.Substring(16),
					flags.HasFlag(Win32FindReplaceFlags.FR_DOWN),
					flags.HasFlag(Win32FindReplaceFlags.FR_MATCHCASE),
					!new_find);
			}
			#endregion
		}

	}
}