using CefSharp;
using CefSharp.Enums;
using CefSharp.Structs;
using System;
using yomigaeri_shared;

namespace yomigaeri_backend.Browser.Handlers
{
	internal sealed class MyDisplayHandler : CefSharp.Handler.DisplayHandler
	{
		private readonly SynchronizerState m_SyncState;
		private readonly Action m_SyncProc;

		public MyDisplayHandler(SynchronizerState syncState, Action syncProc)
		{
			m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
			m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");
		}

		protected override bool OnTooltipChanged(IWebBrowser chromiumWebBrowser, ref string text)
		{
			// Maybe they will implement it one day. Frontend is ready for it.

			m_SyncState.Tooltip = text;
			m_SyncProc.Invoke();

			return true;
		}

		protected override void OnAddressChanged(IWebBrowser chromiumWebBrowser, AddressChangedEventArgs addressChangedArgs)
		{
			m_SyncState.Address = addressChangedArgs.Address;
			m_SyncProc.Invoke();
		}

		protected override void OnLoadingProgressChange(IWebBrowser chromiumWebBrowser, IBrowser browser, double progress)
		{
			m_SyncState.StatusProgress = (int)(progress * 100D);
			m_SyncProc.Invoke();
		}

		protected override void OnTitleChanged(IWebBrowser chromiumWebBrowser, TitleChangedEventArgs titleChangedArgs)
		{
			if (string.IsNullOrEmpty(titleChangedArgs.Title))
				m_SyncState.PageTitle = chromiumWebBrowser.Address;
			else
				m_SyncState.PageTitle = titleChangedArgs.Title;

			m_SyncProc.Invoke();
		}
	}
}