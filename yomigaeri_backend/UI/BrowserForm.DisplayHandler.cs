using CefSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace yomigaeri_backend.UI
{
	partial class BrowserForm
	{
		private class MyDisplayHandler : CefSharp.Handler.DisplayHandler
		{
			private BrowserState m_BrowserState;
			private readonly Action m_SyncProc;

			public MyDisplayHandler(BrowserState? browserState, Action syncProc)
			{
				m_BrowserState = browserState ?? throw new ArgumentNullException("browserState");
				m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");
			}

			protected override void OnLoadingProgressChange(IWebBrowser chromiumWebBrowser, IBrowser browser, double progress)
			{
				base.OnLoadingProgressChange(chromiumWebBrowser, browser, progress);

				m_BrowserState.StatusProgress = (int)(progress * 100D);
				m_SyncProc.Invoke();
			}
		}
	}
}
