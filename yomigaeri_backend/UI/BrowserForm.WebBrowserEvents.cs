using CefSharp;
using CefSharp.WinForms;
using Microsoft.Win32;
using System;
using System.Globalization;

namespace yomigaeri_backend.UI
{
	partial class BrowserForm
	{
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
				m_SyncState.SyncAll();
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
			if (string.IsNullOrEmpty(e.Title))
				m_SyncState.PageTitle = Program.WebBrowser.Address;
			else
				m_SyncState.PageTitle = e.Title;

			SynchronizeWithFrontend();
		}

		private void WebBrowser_AddressChanged(object sender, AddressChangedEventArgs e)
		{
			m_SyncState.Address = e.Address;
			SynchronizeWithFrontend();
		}

		private void WebBrowser_FrameLoadStart(object sender, FrameLoadStartEventArgs e)
		{
			Logging.WriteLineToLog("WebBrowserEvents: FrameLoadStart: Main? {0}", e.Frame.IsMain);

			if (!e.Frame.IsMain)
				return;

			this.InvokeOnUiThreadIfRequired(() =>
			{
				try
				{
					Uri uri = new Uri(e.Url);
					m_SyncState.StatusText = string.Format(CultureInfo.CurrentUICulture,
						Resources.Strings.StatusConnecting, uri.Authority);
				} catch (UriFormatException)
				{
					m_SyncState.StatusText = string.Format(CultureInfo.CurrentUICulture,
						Resources.Strings.StatusConnecting, e.Url);
				}

				m_SyncState.PageTitle = Program.WebBrowser.Address;
				SynchronizeWithFrontend();
			});
		}

		private void WebBrowser_FrameLoadEnd(object sender, FrameLoadEndEventArgs e)
		{
			Logging.WriteLineToLog("WebBrowserEvents: FrameLoadEnd: Main? {0}", e.Frame.IsMain);

			if (!e.Frame.IsMain)
				return;

			this.InvokeOnUiThreadIfRequired(() =>
			{
				m_HistoryProcessor.PrepareForVisit();
				Program.WebBrowser.GetBrowserHost().GetNavigationEntries(m_HistoryProcessor, false);

				m_SyncState.AddHistoryItem = true;
				SynchronizeWithFrontend();
			});
		}

		private void WebBrowser_IsBrowserInitializedChanged(object sender, EventArgs e)
		{
			var b = ((ChromiumWebBrowser)sender);

			this.InvokeOnUiThreadIfRequired(() => b.Focus());
		}

	}
}
