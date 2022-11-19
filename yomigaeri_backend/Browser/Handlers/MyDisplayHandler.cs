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

		private enum IDC_STANDARD_CURSORS
		{
			IDC_ARROW = 32512,
			IDC_IBEAM = 32513,
			IDC_WAIT = 32514,
			IDC_CROSS = 32515,
			IDC_UPARROW = 32516,
			IDC_SIZE = 32640,
			IDC_ICON = 32641,
			IDC_SIZENWSE = 32642,
			IDC_SIZENESW = 32643,
			IDC_SIZEWE = 32644,
			IDC_SIZENS = 32645,
			IDC_SIZEALL = 32646,
			IDC_NO = 32648,
			IDC_HAND = 32649,
			IDC_APPSTARTING = 32650,
			IDC_HELP = 32651
		}

		private enum IDC_CUSTOM_CURSORS
		{
			// The IDs are the same as in the ACTIVEX.RES file of the frontend.

			IDC_ALIAS = 101,
			IDC_CELL = 102,
			IDC_COL_RESIZE = 103,
			IDC_COPY = 104,
			IDC_HAND_GRAB = 105,
			IDC_HAND_GRABBING = 106,
			IDC_NONE = 107,
			IDC_PAN_EAST = 108,
			IDC_PAN_MIDDLE = 109,
			IDC_PAN_NORTH = 110,
			IDC_PAN_NORTH_EAST = 111,
			IDC_PAN_NORTH_WEST = 112,
			IDC_PAN_SOUTH = 113,
			IDC_PAN_SOUTH_EAST = 114,
			IDC_PAN_SOUTH_WEST = 115,
			IDC_PAN_WEST = 116,
			IDC_ROW_RESIZE = 117,
			IDC_VERTICAL_TEXT = 118,
			IDC_ZOOM_IN = 119,
			IDC_ZOOM_OUT = 120
		}

		public MyDisplayHandler(SynchronizerState syncState, Action syncProc)
		{
			m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
			m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");
		}

		protected override bool OnCursorChange(IWebBrowser chromiumWebBrowser, IBrowser browser, IntPtr cursor, CursorType type, CursorInfo customCursorInfo)
		{
			/* Before you piss all over this code to mark your territory, heed this notice:
			 * 
			 * Ancient MSTSCAX.DLL 5.1.2600.2180, the last version that works on Win9x, has
			 * severe issues with cursor handling when not in fullscreen mode. The POC for
			 * cursors was done at the very beginning when frontend development used a fully
			 * patched WinXP SP3 system with MSTSCAX.DLL 6.1.7601.22476. That version does
			 * cursors correctly in windowed mode, so a lot of assumptions were made that
			 * ultimately turned out to be completely wrong.
			 *
			 * If you use MSTSCAX.DLL 5.1.2600.2180 and let CEFSharp deal with cursors like
			 * it wants to by default, the result is that the cursor either never changes or
			 * only changes for a fraction of a second before returning back to the default
			 * "Arrow" cursor on the client.
			 * 
			 * The workaround, that works well and saves reading the registry, copying .CUR
			 * or .ANI files from the frontend to the backend, and finally calling
			 * SystemParametersInfo(SPI_SETCURSORS,...) to apply them, is as follows:
			 *  
			 * 1. Local cursor is hidden in BrowserForm's constructor to prevent crazy
			 *    flickering that I cannot find the source of.
			 *   
			 * 2. CEFSharp tells us what it wants the current cursor to be (this method)
			 * 
			 * 3. We filter CEFSharp's wish to cursors Windows supports by default,
			 *    silently ignoring unsupported cursors by setting them to "Arrow".
			 *   
			 * 4. The Filtered cursor is sent to frontend, which sets it locally on the
			 *    ActiveX control.
			 *    
			 * 5. Since the cursor inside the RDP session got hidden, MSTSCAX thinks that
			 *    there is no cursor and doesn't interfere. It crucially also doesn't
			 *    care to hide the cursor locally, due to abovementioned bugginess in
			 *    windowed mode.
			 * 
			 * This approach seems to work very well. It only took about a day to figure
			 * out all of the details. *sigh*
			 */

			int cursorid;

			switch (type)
			{
				/* Standard cusrors that Windows supports. */

				case CursorType.Custom:
					Logging.WriteLineToLog("Lame! Replacing unsupported CSS custom cursor with IDC_ARROW.");
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_ARROW;
					break;
				default:
				case CursorType.ContextMenu:
				case CursorType.Pointer:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_ARROW;
					break;
				case CursorType.Cross:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_CROSS;
					break;
				case CursorType.Hand:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_HAND;
					break;
				case CursorType.IBeam:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_IBEAM;
					break;
				case CursorType.Wait:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_WAIT;
					break;
				case CursorType.Help:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_HELP;
					break;
				case CursorType.EastResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZEWE;
					break;
				case CursorType.NorthResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZENS;
					break;
				case CursorType.NortheastResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZENESW;
					break;
				case CursorType.NorthwestResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZENWSE;
					break;
				case CursorType.SouthResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZENS;
					break;
				case CursorType.SoutheastResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZENWSE;
					break;
				case CursorType.SouthwestResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZENESW;
					break;
				case CursorType.WestResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZEWE;
					break;
				case CursorType.NorthSouthResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZENS;
					break;
				case CursorType.EastWestResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZEWE;
					break;
				case CursorType.NortheastSouthwestResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZENESW;
					break;
				case CursorType.NorthwestSoutheastResize:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZENWSE;
					break;
				case CursorType.Move:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_SIZEALL;
					break;
				case CursorType.Progress:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_APPSTARTING;
					break;
				case CursorType.NoDrop:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_NO;
					break;
				case CursorType.NotAllowed:
					cursorid = (int)IDC_STANDARD_CURSORS.IDC_NO;
					break;

				/* Custom crap that Chromium supports as well. */

				case CursorType.ColumnResize:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_COL_RESIZE;
					break;
				case CursorType.RowResize:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_ROW_RESIZE;
					break;
				case CursorType.MiddlePanning:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_PAN_MIDDLE;
					break;
				case CursorType.EastPanning:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_PAN_EAST;
					break;
				case CursorType.NorthPanning:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_PAN_NORTH;
					break;
				case CursorType.NortheastPanning:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_PAN_NORTH_EAST;
					break;
				case CursorType.NorthwestPanning:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_PAN_NORTH_WEST;
					break;
				case CursorType.SouthPanning:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_PAN_SOUTH;
					break;
				case CursorType.SoutheastPanning:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_PAN_SOUTH_EAST;
					break;
				case CursorType.SouthwestPanning:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_PAN_SOUTH_WEST;
					break;
				case CursorType.WestPanning:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_PAN_WEST;
					break;
				case CursorType.VerticalText:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_VERTICAL_TEXT;
					break;
				case CursorType.Cell:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_CELL;
					break;
				case CursorType.ZoomIn:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_ZOOM_IN;
					break;
				case CursorType.ZoomOut:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_ZOOM_OUT;
					break;
				case CursorType.Grab:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_HAND_GRAB;
					break;
				case CursorType.Grabbing:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_HAND_GRABBING;
					break;
				case CursorType.Copy:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_COPY;
					break;
				case CursorType.Alias:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_ALIAS;
					break;
				case CursorType.None:
					cursorid = (int)IDC_CUSTOM_CURSORS.IDC_NONE;
					break;
			}

			m_SyncState.Cursor = cursorid;
			m_SyncProc.Invoke();

			return true;
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