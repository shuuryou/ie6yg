using System.Globalization;

namespace yomigaeri_backend.UI
{
	partial class BrowserForm
	{
		private struct SynchronizerState
		{
			public bool Visible { get; set; }
			public bool CanGoBack { get; set; }
			public bool CanGoForward { get; set; }
			public bool CanReload { get; set; }
			public bool IsLoading { get; set; }

			public string StatusText { get; set; }
			public int StatusProgress { get; set; }

			public string Address { get; set; }
			public string PageTitle { get; set; }

			public bool EnableTextCut { get; set; }
			public bool EnableTextCopy { get; set; }
			public bool EnableTextPaste { get; set; }
		}

		private void SynchronizeWithFrontend()
		{
			if (!m_PrevSyncState.HasValue || m_PrevSyncState.Value.Visible != m_SyncState.Visible)
			{
				if (m_SyncState.Visible)
					RDPVirtualChannel.WriteChannel("VISIBLE");
				else
					RDPVirtualChannel.WriteChannel("INVISIB");
			}

			if (!m_PrevSyncState.HasValue || m_PrevSyncState.Value.CanGoBack != m_SyncState.CanGoBack)
			{
				if (m_SyncState.CanGoBack)
					RDPVirtualChannel.WriteChannel("BBACKON");
				else
					RDPVirtualChannel.WriteChannel("BBACKOF");
			}

			if (!m_PrevSyncState.HasValue || m_PrevSyncState.Value.CanGoForward != m_SyncState.CanGoForward)
			{
				if (m_SyncState.CanGoForward)
					RDPVirtualChannel.WriteChannel("BFORWON");
				else
					RDPVirtualChannel.WriteChannel("BFORWOF");
			}

			if (!m_PrevSyncState.HasValue || m_PrevSyncState.Value.CanReload != m_SyncState.CanReload)
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

			if (!m_PrevSyncState.HasValue || m_PrevSyncState.Value.IsLoading != m_SyncState.IsLoading)
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

			if (!m_PrevSyncState.HasValue || m_PrevSyncState.Value.StatusProgress != m_SyncState.StatusProgress)
			{
				int progress = m_SyncState.StatusProgress;

				if (progress < 0)
					progress = 0;

				if (progress > 100)
					progress = 100;

				RDPVirtualChannel.WriteChannel("PROGRES" + progress.ToString(CultureInfo.InvariantCulture));
			}

			if (!m_PrevSyncState.HasValue || m_PrevSyncState.Value.Address != m_SyncState.Address)
			{
				RDPVirtualChannel.WriteChannel("ADDRESS" + m_SyncState.Address);
			}

			if (!m_PrevSyncState.HasValue || m_PrevSyncState.Value.StatusText != m_SyncState.StatusText)
			{
				RDPVirtualChannel.WriteChannel("STATUST" + m_SyncState.StatusText);
			}

			if (!m_PrevSyncState.HasValue || m_PrevSyncState.Value.PageTitle != m_SyncState.PageTitle)
			{
				RDPVirtualChannel.WriteChannel("PGTITLE" + m_SyncState.PageTitle);
			}

			m_PrevSyncState = m_SyncState;
		}
	}
}
