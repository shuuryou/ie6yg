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
			Logging.WriteLineToLog("WebBrowserEvents: SessionSwitch with eason: {0}", e.Reason);

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

		private void WebBrowser_FrameLoadStart(object sender, FrameLoadStartEventArgs e)
		{
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
			if (!e.Frame.IsMain)
				return;

			m_TravelLog.PrepareForVisit();
			Program.WebBrowser.GetBrowserHost().GetNavigationEntries(m_TravelLog, false);

			this.InvokeOnUiThreadIfRequired(() =>
			{
				m_SyncState.StatusText = string.Empty;
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
