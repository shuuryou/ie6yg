using CefSharp;
using System;

namespace yomigaeri_backend.Browser.Handlers
{
	internal class MyRequestHandler : CefSharp.Handler.RequestHandler
	{
		private readonly SynchronizerState m_SyncState;
		private readonly Action m_SyncProc;

		public MyRequestHandler(SynchronizerState syncState, Action syncProc)
		{
			m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
			m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");
		}

		protected async override void OnDocumentAvailableInMainFrame(IWebBrowser chromiumWebBrowser, IBrowser browser)
		{
			NavigationEntry current = await Program.WebBrowser.GetVisibleNavigationEntryAsync();

			if (!current.SslStatus.IsSecureConnection)
				m_SyncState.SSLIcon = SynchronizerState.SSLIconState.None;
			else if (current.SslStatus.CertStatus == CertStatus.None)
				m_SyncState.SSLIcon = SynchronizerState.SSLIconState.Secure;
			else
				m_SyncState.SSLIcon = SynchronizerState.SSLIconState.SecureBadCert;

			base.OnDocumentAvailableInMainFrame(chromiumWebBrowser, browser);
		}


	}
}