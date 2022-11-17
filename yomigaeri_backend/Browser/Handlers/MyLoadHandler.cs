using CefSharp;
using CefSharp.DevTools;
using CefSharp.DevTools.Page;
using System;
using System.Globalization;
using System.Web;
using static yomigaeri_backend.Browser.SynchronizerState;

namespace yomigaeri_backend.Browser.Handlers
{
	internal class MyLoadHandler : CefSharp.Handler.LoadHandler
	{
		private const string SEARCH_URL_DEFAULT = "https://www.startpage.com/sp/search?query=%s";

		private readonly SynchronizerState m_SyncState;
		private readonly TravelLog m_TravelLog;
		private readonly Action m_SyncProc;

		public MyLoadHandler(TravelLog travelLog, SynchronizerState syncState, Action syncProc)
		{
			m_TravelLog = travelLog ?? throw new ArgumentNullException("travelLog");
			m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
			m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");
		}

		protected override void OnLoadingStateChange(IWebBrowser chromiumWebBrowser, LoadingStateChangedEventArgs loadingStateChangedArgs)
		{
			if (loadingStateChangedArgs.CanGoBack)
				m_SyncState.ToolbarButtons |= FrontendToolbarButtons.Back;
			else
				m_SyncState.ToolbarButtons &= ~FrontendToolbarButtons.Back;

			if (loadingStateChangedArgs.CanGoForward)
				m_SyncState.ToolbarButtons |= FrontendToolbarButtons.Forward;
			else
				m_SyncState.ToolbarButtons &= ~FrontendToolbarButtons.Forward;

			if (loadingStateChangedArgs.CanReload)
			{
				m_SyncState.ToolbarButtons |= FrontendToolbarButtons.Refresh;
				m_SyncState.MenuBarItems |= FrontendMenuBarItems.Refresh;
			}
			else
			{
				m_SyncState.ToolbarButtons &= ~FrontendToolbarButtons.Refresh;
				m_SyncState.MenuBarItems &= ~FrontendMenuBarItems.Refresh;
			}

			if (loadingStateChangedArgs.IsLoading)
			{
				m_SyncState.ToolbarButtons |= FrontendToolbarButtons.Stop;
				m_SyncState.MenuBarItems |= FrontendMenuBarItems.Stop;
			}
			else
			{
				m_SyncState.ToolbarButtons &= ~FrontendToolbarButtons.Stop;
				m_SyncState.MenuBarItems &= ~FrontendMenuBarItems.Stop;
				m_SyncState.StatusProgress = 0;
			}

			m_SyncProc.Invoke();

			base.OnLoadingStateChange(chromiumWebBrowser, loadingStateChangedArgs);
		}

		protected override void OnLoadError(IWebBrowser chromiumWebBrowser, LoadErrorEventArgs loadErrorArgs)
		{
			if (loadErrorArgs.ErrorCode == CefErrorCode.Aborted)
				return;

			Logging.WriteLineToLog("LoadHandler: LoadError: URL: \"{0}\" ErrorCode: \"{1}\". ErrorText: \"{2}\". Main Frame? {3}",
				loadErrorArgs.FailedUrl, loadErrorArgs.ErrorCode, loadErrorArgs.ErrorText, loadErrorArgs.Frame.IsMain);

			string error_html = Resources.BROWSERERROR_HTML;

			Uri url = new Uri(loadErrorArgs.FailedUrl);

			{
				string homepage_url = string.Format(CultureInfo.InvariantCulture, "{0}://{1}", url.Scheme, url.Authority);
				string homepage_name = url.Host;

				error_html = error_html.Replace("%HOMEPAGE_URL%", HttpUtility.HtmlEncode(homepage_url));
				error_html = error_html.Replace("%HOMEPAGE_NAME%", HttpUtility.HtmlEncode(homepage_name));
			}

			{
				string search_url = Program.Settings.Get("IE6YG", "ErrorPageSearchURL");
				if (string.IsNullOrEmpty(search_url))
					search_url = SEARCH_URL_DEFAULT;

				search_url = search_url.Replace("%s", HttpUtility.UrlEncode(loadErrorArgs.FailedUrl));

				error_html = error_html.Replace("%SEARCH_URL%", HttpUtility.HtmlEncode(search_url));
			}

			{
				string error_code = string.Format(CultureInfo.InvariantCulture, "0x{0:X8}", (int)loadErrorArgs.ErrorCode);
				string error_text = loadErrorArgs.ErrorText;

				error_html = error_html.Replace("%ERROR_CODE%", HttpUtility.HtmlEncode(error_code));
				error_html = error_html.Replace("%ERROR_TEXT%", HttpUtility.HtmlEncode(error_text));
			}

			if (loadErrorArgs.Frame.IsMain)
			{
				loadErrorArgs.Browser.SetMainFrameDocumentContentAsync(error_html);
				m_SyncState.Address = loadErrorArgs.FailedUrl;
				m_SyncProc.Invoke();
			}
			else
			{
				loadErrorArgs.Frame.LoadHtml(error_html);
			}
		}

		protected override void OnFrameLoadStart(IWebBrowser chromiumWebBrowser, FrameLoadStartEventArgs frameLoadStartArgs)
		{
			if (!frameLoadStartArgs.Frame.IsMain)
			{
				base.OnFrameLoadStart(chromiumWebBrowser, frameLoadStartArgs);
				return;
			}

			try
			{
				Uri uri = new Uri(frameLoadStartArgs.Url);
				m_SyncState.StatusText = string.Format(CultureInfo.CurrentUICulture,
					Resources.StatusConnecting, uri.Authority);
			}
			catch (UriFormatException)
			{
				m_SyncState.StatusText = string.Format(CultureInfo.CurrentUICulture,
					Resources.StatusConnecting, frameLoadStartArgs.Url);
			}

			m_SyncState.SSLIcon = SynchronizerState.SSLIconState.None;
			m_SyncState.PageTitle = Program.WebBrowser.Address;
			m_SyncProc.Invoke();

			base.OnFrameLoadStart(chromiumWebBrowser, frameLoadStartArgs);

		}

		protected override void OnFrameLoadEnd(IWebBrowser chromiumWebBrowser, FrameLoadEndEventArgs frameLoadEndArgs)
		{
			if (!frameLoadEndArgs.Frame.IsMain)
			{
				base.OnFrameLoadEnd(chromiumWebBrowser, frameLoadEndArgs);
				return;
			}

			m_TravelLog.PrepareForVisit();
			Program.WebBrowser.GetBrowserHost().GetNavigationEntries(m_TravelLog, false);

			m_SyncState.StatusText = string.Empty;
			m_SyncState.AddHistoryItem = true;
			m_SyncProc.Invoke();

			base.OnFrameLoadEnd(chromiumWebBrowser, frameLoadEndArgs);
		}
	}
}