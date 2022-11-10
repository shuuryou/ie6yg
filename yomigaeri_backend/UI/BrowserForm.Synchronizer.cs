using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace yomigaeri_backend.UI
{
	partial class BrowserForm
	{
		private class SynchronizerState
		{
			private readonly BitArray m_Changes;
			public enum Change : int
			{
				Visible = 0,
				CanGoBack = 1,
				CanGoForward = 2,
				CanReload = 3,
				IsLoading = 4,
				StatusText = 5,
				StatusProgress = 6,
				Address = 7,
				PageTitle = 8,
				AddHistoryItem = 9,
				EnableTextCut = 10,
				EnableTextCopy = 11,
				EnableTextPaste = 12,
				MiniHistory = 13
			}

			private const int CHANGE_ITEMS = 14;

			private bool m_Visible;
			public bool Visible
			{
				get { return m_Visible; }
				set
				{
					m_Visible = value;
					m_Changes[(int)Change.Visible] = true;
				}
			}

			private bool m_CanGoBack;
			public bool CanGoBack
			{
				get { return m_CanGoBack; }
				set
				{
					m_CanGoBack = value;
					m_Changes[(int)Change.CanGoBack] = true;
				}
			}

			private bool m_CanGoForward;
			public bool CanGoForward
			{
				get { return m_CanGoForward; }
				set
				{
					m_CanGoForward = value;
					m_Changes[(int)Change.CanGoForward] = true;
				}
			}

			private bool m_CanReload;
			public bool CanReload
			{
				get { return m_CanReload; }
				set
				{
					m_CanReload = value;
					m_Changes[(int)Change.CanReload] = true;
				}
			}

			private bool m_IsLoading;
			public bool IsLoading
			{
				get { return m_IsLoading; }
				set
				{
					m_IsLoading = value;
					m_Changes[(int)Change.IsLoading] = true;
				}
			}

			private string m_StatusText;
			public string StatusText
			{
				get { return m_StatusText; }
				set
				{
					m_StatusText = value;
					m_Changes[(int)Change.StatusText] = true;
				}
			}

			private int m_StatusProgress;
			public int StatusProgress
			{
				get { return m_StatusProgress; }
				set
				{
					m_StatusProgress = value;
					m_Changes[(int)Change.StatusProgress] = true;
				}
			}

			private string m_Address;
			public string Address
			{
				get { return m_Address; }
				set
				{
					m_Address = value;
					m_Changes[(int)Change.Address] = true;
				}
			}

			private string m_PageTitle;
			public string PageTitle
			{
				get { return m_PageTitle; }
				set
				{
					m_PageTitle = value;
					m_Changes[(int)Change.PageTitle] = true;
				}
			}

			private bool m_AddHistoryItem;
			public bool AddHistoryItem
			{
				get { return m_AddHistoryItem; }
				set
				{
					m_AddHistoryItem = value;
					m_Changes[(int)Change.AddHistoryItem] = true;
				}
			}

			private bool m_EnableTextCut;
			public bool EnableTextCut
			{
				get { return m_EnableTextCut; }
				set
				{
					m_EnableTextCut = value;
					m_Changes[(int)Change.EnableTextCut] = true;
				}
			}

			private bool m_EnableTextCopy;
			public bool EnableTextCopy
			{
				get { return m_EnableTextCopy; }
				set
				{
					m_EnableTextCopy = value;
					m_Changes[(int)Change.EnableTextCopy] = true;
				}
			}

			private bool m_EnableTextPaste;
			public bool EnableTextPaste
			{
				get { return m_EnableTextPaste; }
				set
				{
					m_EnableTextPaste = value;
					m_Changes[(int)Change.EnableTextPaste] = true;
				}
			}

			private bool m_MiniHistory;

			public bool MiniHistory
			{
				get { return m_MiniHistory; }
				set
				{
					m_MiniHistory = value;
					m_Changes[(int)Change.MiniHistory] = true;
				}
			}

			public SynchronizerState()
			{
				m_Changes = new BitArray(CHANGE_ITEMS);
			}

			public void SyncAll()
			{
				m_Changes.SetAll(true);
			}

			public void SyncNone()
			{
				m_Changes.SetAll(false);
			}

			public bool IsChanged(Change what)
			{
				return m_Changes[(int)what];
			}
		}

		private void SynchronizeWithFrontend()
		{ 
			if (m_SyncState.IsChanged(SynchronizerState.Change.Visible))
			{
				if (m_SyncState.Visible)
					RDPVirtualChannel.WriteChannel("VISIBLE");
				else
					RDPVirtualChannel.WriteChannel("INVISIB");
			}

			if (m_SyncState.IsChanged(SynchronizerState.Change.CanGoBack))
			{
				if (m_SyncState.CanGoBack)
					RDPVirtualChannel.WriteChannel("BBACKON");
				else
					RDPVirtualChannel.WriteChannel("BBACKOF");
			}

			if (m_SyncState.IsChanged(SynchronizerState.Change.CanGoForward))
			{
				if (m_SyncState.CanGoForward)
					RDPVirtualChannel.WriteChannel("BFORWON");
				else
					RDPVirtualChannel.WriteChannel("BFORWOF");
			}

			if (m_SyncState.IsChanged(SynchronizerState.Change.CanReload))
			{
				if (m_SyncState.CanReload)
				{
					RDPVirtualChannel.WriteChannel("BREFRON");
				}
				else
				{
					RDPVirtualChannel.WriteChannel("BREFROF");
				}
			}

			if (m_SyncState.IsChanged(SynchronizerState.Change.IsLoading))
			{
				if (m_SyncState.IsLoading)
				{
					RDPVirtualChannel.WriteChannel("BSTOPON");
				}
				else
				{
					RDPVirtualChannel.WriteChannel("BSTOPOF");
					RDPVirtualChannel.WriteChannel("PROGRES0");
				}
			}

			if (m_SyncState.IsChanged(SynchronizerState.Change.AddHistoryItem))
			{
				if (!string.IsNullOrEmpty(m_SyncState.Address))
					RDPVirtualChannel.WriteChannel("ADDHIST" +
						m_SyncState.PageTitle + '\x1' + m_SyncState.Address);

				m_SyncState.AddHistoryItem = false;
			}

			if (m_SyncState.IsChanged(SynchronizerState.Change.StatusProgress ))
			{
				int progress = m_SyncState.StatusProgress;

				if (progress < 0)
					progress = 0;

				if (progress > 100)
					progress = 100;

				RDPVirtualChannel.WriteChannel("PROGRES" + progress.ToString(CultureInfo.InvariantCulture));
			}

			if (m_SyncState.IsChanged(SynchronizerState.Change.Address))
			{
				RDPVirtualChannel.WriteChannel("ADDRESS" + m_SyncState.Address);
			}

			if (m_SyncState.IsChanged(SynchronizerState.Change.StatusText))
			{
				RDPVirtualChannel.WriteChannel("STATUST" + m_SyncState.StatusText);
			}

			if (m_SyncState.IsChanged(SynchronizerState.Change.PageTitle))
			{
				RDPVirtualChannel.WriteChannel("PGTITLE" + m_SyncState.PageTitle);
			}

			if (m_SyncState.IsChanged( SynchronizerState.Change.MiniHistory))
			{
				RDPVirtualChannel.WriteChannel("MINHIBK" +
					m_HistoryProcessor.MakeMenuStringForFrontend(TravelLog.NavigationKind.Back));

				RDPVirtualChannel.WriteChannel("MINHIFW" +
					m_HistoryProcessor.MakeMenuStringForFrontend(TravelLog.NavigationKind.Forward));
			}

			m_SyncState.SyncNone();
		}
	}
}
