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
		private class MyRequestHandler : CefSharp.Handler.RequestHandler 
		{
			private readonly SynchronizerState m_SyncState;
			private readonly Action m_SyncProc;

			public MyRequestHandler(SynchronizerState syncState, Action syncProc)
			{
				m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
				m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");
			}
		}
	}
}
