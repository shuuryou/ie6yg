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
			private SynchronizerState m_SyncState;
			private readonly Action m_SyncProc;

			public MyDisplayHandler(SynchronizerState? syncState, Action syncProc)
			{
				m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
				m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");
			}

			protected override void OnLoadingProgressChange(IWebBrowser chromiumWebBrowser, IBrowser browser, double progress)
			{
				base.OnLoadingProgressChange(chromiumWebBrowser, browser, progress);

				m_SyncState.StatusProgress = (int)(progress * 100D);
				m_SyncProc.Invoke();
			}
		}
	}
}
