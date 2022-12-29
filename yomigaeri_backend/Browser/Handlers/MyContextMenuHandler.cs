using CefSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using yomigaeri_shared;

namespace yomigaeri_backend.Browser.Handlers
{
	internal sealed class MyContextMenuHandler : CefSharp.Handler.ContextMenuHandler
	{
		/* Menu handling is not easy. The way it was implemented here was an
		 * attempt to keep it as generic as possible, but it's never that
		 * simple.
		 * 
		 * Basically what I am attempting to do here is:
		 * 
		 * 1. Intercept the request to CEF to pop up its menu, then determine
		 *    state and prepare a string for the frontend so it can configure
		 *    one of four menu templates accordingly.
		 *    
		 * 2. Send the menu string to the frontend and the store the callback.
		 *    The frontend then pops up the menu, and the user can select an
		 *    item.
		 *    
		 * 3. When the item is selected, transmit the selected index back to
		 *    the backend, which invokes the stored callback, bringing us back
		 *    into this class.
		 *    
		 * 4. Translate the index into requested command, and perform that
		 *    command.
		 *    
		 * 5. If no further interaction is required, stop. Otherwise, send
		 *    another command to the frontend, so that it can perform whatever
		 *    is necessary. For example, this is required for anything that
		 *    needs a "Save As" dialog to be shown on the frontend.
		 */

		private readonly SynchronizerState m_SyncState;
		private readonly Action m_SyncProc;

		public IRunContextMenuCallback ContextMenu_Callback { get; private set; }

		/* The following enumerations must be kept strictly in sync with the
		 * menu in the frontend, or things will go horribly wrong. To add new
		 * commands:
		 * 
		 * 1. Define them with an index in the VB6 menu editor
		 * 2. Modify the relevant enumeration below
		 * 3. Add code to show/hide/enable/disable the new menu item
		 * 4. Add code to handle the command the menu item represents
		 */

		private enum FrontendContextMenuKind
		{
			None = 0,
			Selection = 1,
			Default = 2,
			Image = 3,
			Anchor = 4
		}

		private enum FrontendContextMenuSelectionItems
		{
			Cut = 0,
			Copy = 1,
			Paste = 2,
			SelectAll = 3,
			Print = 4
		}

		private enum FrontendContextMenuDefaultItems
		{
			Back = 0,
			Forward = 1,
			SelectAll = 3,
			CreateShortcut = 5,
			AddToFavorites = 6,
			ViewSource = 7,
			Print = 9,
			Refresh = 10,
			Properties = 12
		}

		private enum FrontendContextMenuImageItems
		{
			OpenLink = 0,
			OpenLinkInNewWindow = 1,
			ShowPicture = 3,
			ShowVideo = 4,
			ShowAudio = 5,
			SavePictureAs = 6,
			SaveVideoAs = 7,
			SaveAudioAs = 8,
			SetAsBackground = 9,
			Copy = 11,
			CopyShortcut = 12,
			AddToFavorites = 14,
			Properties = 16
		}

		private enum FrontendContextMenuAnchorItems
		{
			Open = 0,
			OpenInNewWindow = 1,
			SaveTargetAs = 2,
			Copy = 4,
			CopyShortcut = 5,
			AddToFavorites = 7,
			Properties = 9
		}

		public MyContextMenuHandler(SynchronizerState syncState, Action syncProc)
		{
			m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
			m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");

			ContextMenu_Callback = null;
		}

		protected override void OnBeforeContextMenu(IWebBrowser chromiumWebBrowser, IBrowser browser, IFrame frame, IContextMenuParams parameters, IMenuModel model)
		{
			// Do not allow another context menu while one is still being processed

			if (ContextMenu_Callback != null && !ContextMenu_Callback.IsDisposed)
				model.Clear();
		}

		protected override bool RunContextMenu(IWebBrowser chromiumWebBrowser, IBrowser browser, IFrame frame, IContextMenuParams parameters, IMenuModel model, IRunContextMenuCallback callback)
		{

			FrontendContextMenuKind kind = FrontendContextMenuKind.None;

			switch (parameters.TypeFlags)
			{
				case ContextMenuType.Editable:
				case ContextMenuType.Selection:
					kind = FrontendContextMenuKind.Selection;
					break;
				case ContextMenuType.Page:
				case ContextMenuType.Frame:
				case ContextMenuType.None:
					kind = FrontendContextMenuKind.Default;
					break;
				case ContextMenuType.Media:
					kind = FrontendContextMenuKind.Image;
					break;
				case ContextMenuType.Link:
					kind = FrontendContextMenuKind.Anchor;
					break;
			}

			if (kind == FrontendContextMenuKind.None)
			{
				Logging.WriteLineToLog("RunContextMenu: Requested context menu type not understood: {0}", parameters.TypeFlags);
				callback.Cancel();
				return true;
			}

			/*
			 * I don't like this very much, but it should do the trick for the
			 * time being. The syntax is as follows:
			 * 
			 * [X] [Y] [SELECTION|DEFAULT|IMAGE|ANCHOR] [I1][I|V][E|D] [I2][I|V][E|D] ... [In][I|V][E|D]
			 * 1   2   3                                4   5    6     4   5    6         4   5    6
			 *
			 * Where:
			 * 1 is the X coordinate to pop up the menu at
			 * 2 is the Y coordinate to pop up the menu at
			 * 3 is one of the mentioned menu identifiers
			 * 4 is theindex from one of the FrontendContextMenu*Items enumerations
			 * 5 is either "I" = Invisible, or "V" = Visible
			 * 6 is either "E" = Enabled, or "D" = Disabled
			 * 
			 * Sample:
			 * 321 123 SELECTION 0VD 1VE 2VD 3VE 4VE
			 * 
			 * The frontend parses the above and then makes the menu appear as
			 * desired.
			 */

			StringBuilder menu_params = new StringBuilder();

			menu_params.AppendFormat(CultureInfo.InvariantCulture, "{0} {1} ", parameters.XCoord, parameters.YCoord);

			#region Selection
			if (kind == FrontendContextMenuKind.Selection)
			{
				menu_params.Append("SELECTION");
				menu_params.Append(' ');

				#region Cut
				{
					menu_params.Append((int)FrontendContextMenuSelectionItems.Cut);

					menu_params.Append('V');

					bool flag = parameters.TypeFlags == ContextMenuType.Editable &&
						parameters.EditStateFlags.HasFlag(ContextMenuEditState.CanCut);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region Copy
				{
					menu_params.Append((int)FrontendContextMenuSelectionItems.Copy);

					menu_params.Append('V');

					bool flag = parameters.TypeFlags == ContextMenuType.Selection ||
						(parameters.TypeFlags == ContextMenuType.Editable &&
						parameters.EditStateFlags.HasFlag(ContextMenuEditState.CanCopy));

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region Paste
				{
					menu_params.Append((int)FrontendContextMenuSelectionItems.Paste);

					menu_params.Append('V');

					bool flag = parameters.TypeFlags == ContextMenuType.Editable &&
						parameters.EditStateFlags.HasFlag(ContextMenuEditState.CanPaste);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region SelectAll
				{
					menu_params.Append((int)FrontendContextMenuSelectionItems.SelectAll);

					menu_params.Append('V');

					bool flag = parameters.TypeFlags == ContextMenuType.Selection ||
						(parameters.TypeFlags == ContextMenuType.Editable &&
						parameters.EditStateFlags.HasFlag(ContextMenuEditState.CanSelectAll));

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region Print
				{
					menu_params.Append((int)FrontendContextMenuSelectionItems.Print);

					bool flag;

					flag = parameters.TypeFlags == ContextMenuType.Selection;

					if (flag)
						menu_params.Append('V');
					else
						menu_params.Append('I');

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				goto done;
			}
			#endregion

			#region Default
			if (kind == FrontendContextMenuKind.Default)
			{
				menu_params.Append("DEFAULT");
				menu_params.Append(' ');

				#region Back
				{
					menu_params.Append((int)FrontendContextMenuDefaultItems.Back);

					menu_params.Append('V');

					bool flag = browser.CanGoBack;

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region Forward
				{
					menu_params.Append((int)FrontendContextMenuDefaultItems.Forward);

					menu_params.Append('V');

					bool flag = browser.CanGoForward;

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region SelectAll
				{
					menu_params.Append((int)FrontendContextMenuDefaultItems.SelectAll);

					menu_params.Append('V');
					menu_params.Append('E');

					menu_params.Append(' ');
				}
				#endregion

				#region CreateShortcut
				{
					menu_params.Append((int)FrontendContextMenuDefaultItems.CreateShortcut);

					menu_params.Append('V');
					menu_params.Append('E');

					menu_params.Append(' ');
				}
				#endregion

				#region AddToFavorites
				{
					menu_params.Append((int)FrontendContextMenuDefaultItems.AddToFavorites);

					menu_params.Append('V');
					menu_params.Append('E');

					menu_params.Append(' ');
				}
				#endregion

				#region ViewSource
				{
					menu_params.Append((int)FrontendContextMenuDefaultItems.ViewSource);

					menu_params.Append('V');
					menu_params.Append('E');

					menu_params.Append(' ');
				}
				#endregion

				#region Print
				{
					menu_params.Append((int)FrontendContextMenuDefaultItems.Print);

					menu_params.Append('V');
					menu_params.Append('E');

					menu_params.Append(' ');
				}
				#endregion

				#region Refresh
				{
					menu_params.Append((int)FrontendContextMenuDefaultItems.Refresh);

					menu_params.Append('V');
					menu_params.Append('E');

					menu_params.Append(' ');
				}
				#endregion

				#region Properties
				{
					menu_params.Append((int)FrontendContextMenuDefaultItems.Properties);

					menu_params.Append('V');
					menu_params.Append('E');

					menu_params.Append(' ');
				}
				#endregion

				goto done;
			}
			#endregion

			#region Image (any kind of media)
			if (kind == FrontendContextMenuKind.Image)
			{
				menu_params.Append("IMAGE");
				menu_params.Append(' ');

				#region OpenLink, OpenLinkInNewWindow
				{
					menu_params.Append((int)FrontendContextMenuImageItems.OpenLink);

					menu_params.Append('V');

					bool flag = parameters.TypeFlags.HasFlag(ContextMenuType.Link);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');


					menu_params.Append((int)FrontendContextMenuImageItems.OpenLinkInNewWindow);

					menu_params.Append('V');

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');

				}
				#endregion

				#region ShowPicture
				{
					menu_params.Append((int)FrontendContextMenuImageItems.ShowPicture);

					bool flag = parameters.MediaType == ContextMenuMediaType.Image;

					if (flag)
						menu_params.Append('V');
					else
						menu_params.Append('I');

					flag = flag && !string.IsNullOrEmpty(parameters.SourceUrl);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region ShowVideo
				{
					menu_params.Append((int)FrontendContextMenuImageItems.ShowVideo);

					bool flag = parameters.MediaType == ContextMenuMediaType.Video;

					if (flag)
						menu_params.Append('V');
					else
						menu_params.Append('I');

					flag = flag && !string.IsNullOrEmpty(parameters.SourceUrl);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region ShowAudio
				{
					menu_params.Append((int)FrontendContextMenuImageItems.ShowAudio);

					bool flag = parameters.MediaType == ContextMenuMediaType.Audio;

					if (flag)
						menu_params.Append('V');
					else
						menu_params.Append('I');

					flag = flag && !string.IsNullOrEmpty(parameters.SourceUrl);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region SavePictureAs
				{
					menu_params.Append((int)FrontendContextMenuImageItems.SavePictureAs);

					bool flag = parameters.MediaType == ContextMenuMediaType.Image;

					if (flag)
						menu_params.Append('V');
					else
						menu_params.Append('I');

					flag = flag && parameters.MediaStateFlags.HasFlag(ContextMenuMediaState.CanSave);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region SaveVideoAs
				{
					menu_params.Append((int)FrontendContextMenuImageItems.SaveVideoAs);

					bool flag = parameters.MediaType == ContextMenuMediaType.Video;

					if (flag)
						menu_params.Append('V');
					else
						menu_params.Append('I');

					flag = flag && parameters.MediaStateFlags.HasFlag(ContextMenuMediaState.CanSave);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region SaveAudioAs
				{
					menu_params.Append((int)FrontendContextMenuImageItems.SaveAudioAs);

					bool flag = parameters.MediaType == ContextMenuMediaType.Image;

					if (flag)
						menu_params.Append('V');
					else
						menu_params.Append('I');

					flag = flag && parameters.MediaStateFlags.HasFlag(ContextMenuMediaState.CanSave);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region SetAsBackground
				{
					menu_params.Append((int)FrontendContextMenuImageItems.SetAsBackground);

					menu_params.Append('V');

					bool flag = parameters.MediaType == ContextMenuMediaType.Image &&
						!string.IsNullOrEmpty(parameters.SourceUrl);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region Copy
				{
					menu_params.Append((int)FrontendContextMenuImageItems.Copy);

					menu_params.Append('V');

					bool flag = parameters.MediaType == ContextMenuMediaType.Image &&
						!string.IsNullOrEmpty(parameters.SourceUrl);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region CopyShortcut
				{
					menu_params.Append((int)FrontendContextMenuImageItems.CopyShortcut);

					menu_params.Append('V');

					bool flag = !string.IsNullOrEmpty(parameters.SourceUrl);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region AddToFavorites
				{
					menu_params.Append((int)FrontendContextMenuImageItems.AddToFavorites);

					menu_params.Append('V');

					bool flag = !string.IsNullOrEmpty(parameters.SourceUrl);

					if (flag)
						menu_params.Append('E');
					else
						menu_params.Append('D');

					menu_params.Append(' ');
				}
				#endregion

				#region Properties
				{
					menu_params.Append((int)FrontendContextMenuImageItems.Properties);

					menu_params.Append('V');
					menu_params.Append('E');

					menu_params.Append(' ');
				}
				#endregion

				goto done;
			}
			#endregion

			#region Anchor
			if (kind == FrontendContextMenuKind.Anchor)
			{
				menu_params.Append("ANCHOR");
				menu_params.Append(' ');

				// Everything is always allowed

				#region Open, OpenInNewWindow, SaveTargetAs, Copy, CopyShortcut, AddToFavorites, Properties
				{
					menu_params.Append((int)FrontendContextMenuAnchorItems.Open);
					menu_params.Append('V');
					menu_params.Append('E');
					menu_params.Append(' ');

					menu_params.Append((int)FrontendContextMenuAnchorItems.OpenInNewWindow);
					menu_params.Append('V');
					menu_params.Append('E');
					menu_params.Append(' ');

					menu_params.Append((int)FrontendContextMenuAnchorItems.SaveTargetAs);
					menu_params.Append('V');
					menu_params.Append('E');
					menu_params.Append(' ');

					menu_params.Append((int)FrontendContextMenuAnchorItems.Copy);
					menu_params.Append('V');
					menu_params.Append('E');
					menu_params.Append(' ');

					menu_params.Append((int)FrontendContextMenuAnchorItems.CopyShortcut);
					menu_params.Append('V');
					menu_params.Append('E');
					menu_params.Append(' ');

					menu_params.Append((int)FrontendContextMenuAnchorItems.AddToFavorites);
					menu_params.Append('V');
					menu_params.Append('E');
					menu_params.Append(' ');

					menu_params.Append((int)FrontendContextMenuAnchorItems.Properties);
					menu_params.Append('V');
					menu_params.Append('E');
					menu_params.Append(' ');
				}
				#endregion

				goto done;
			}
		#endregion

		done:

			// xxx todo remove this
			Debug.Assert(menu_params[menu_params.Length - 1] == ' ', "Forgot to append a space to the end of the final command?");

			menu_params.Remove(menu_params.Length - 1, 1);

			Logging.WriteLineToLog("RunContextMenu: Parameters for frontend: {0}", menu_params.ToString());

			m_SyncState.ContextMenu = menu_params.ToString();
			m_SyncProc.Invoke();

			ContextMenu_Callback = callback;

			return true;

		}

		protected override bool OnContextMenuCommand(IWebBrowser chromiumWebBrowser, IBrowser browser, IFrame frame, IContextMenuParams parameters, CefMenuCommand commandId, CefEventFlags eventFlags)
		{
			Logging.WriteLineToLog("DEBUG remove: {0}", commandId);

			ContextMenu_Callback.Dispose();
			ContextMenu_Callback = null;

			return true;
		}
	}
}