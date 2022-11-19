using System;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using yomigaeri_shared;

namespace yomigaeri_backend
{
	public static class FrontendStyling
	{
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool SetSysColors(int cElements, int[] lpaElements, uint[] lpaRgbValues);

		// This is correct; C# int is as VB6 long. VB6 int is a C# short. 
		private const int COLOR_SCROLLBAR = 0;
		private const int COLOR_BACKGROUND = 1;
		private const int COLOR_ACTIVECAPTION = 2;
		private const int COLOR_INACTIVECAPTION = 3;
		private const int COLOR_MENU = 4;
		private const int COLOR_WINDOW = 5;
		private const int COLOR_WINDOWFRAME = 6;
		private const int COLOR_MENUTEXT = 7;
		private const int COLOR_WINDOWTEXT = 8;
		private const int COLOR_CAPTIONTEXT = 9;
		private const int COLOR_ACTIVEBORDER = 10;
		private const int COLOR_INACTIVEBORDER = 11;
		private const int COLOR_APPWORKSPACE = 12;
		private const int COLOR_HIGHLIGHT = 13;
		private const int COLOR_HIGHLIGHTTEXT = 14;
		private const int COLOR_BTNFACE = 15;
		private const int COLOR_BTNSHADOW = 16;
		private const int COLOR_GRAYTEXT = 17;
		private const int COLOR_BTNTEXT = 18;
		private const int COLOR_INACTIVECAPTIONTEXT = 19;
		private const int COLOR_BTNHIGHLIGHT = 20;
		private const int COLOR_2NDACTIVECAPTION = 27;
		private const int COLOR_2NDINACTIVECAPTION = 28;

		public static Font DesiredFont { get; private set; }

		static FrontendStyling()
		{
			DesiredFont = SystemFonts.MessageBoxFont;
		}

		public static void ApplyStyling(string fe_styling)
		{
			/* What the frontend sends:
			 *	strFont & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_SCROLLBAR)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_BACKGROUND)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_ACTIVECAPTION)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_INACTIVECAPTION)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_MENU)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_WINDOW)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_WINDOWFRAME)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_MENUTEXT)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_WINDOWTEXT)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_CAPTIONTEXT)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_ACTIVEBORDER)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_INACTIVEBORDER)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_APPWORKSPACE)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_HIGHLIGHT)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_HIGHLIGHTTEXT)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_BTNFACE)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_BTNSHADOW)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_GRAYTEXT)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_BTNTEXT)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_INACTIVECAPTIONTEXT)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_BTNHIGHLIGHT)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_2NDACTIVECAPTION)) & vbTab & _
             *	LongToRGBHex(GetSysColor(COLOR_2NDINACTIVECAPTION))
             *	
             *	fe_styling_str looks like this: 
             *	
             *	Tahoma	8.25	D4D0C8	3A6EA5	0A246A	808080	D4D0C8	FFFFFF	
             *	000000	000000	000000	FFFFFF	D4D0C8	D4D0C8	808080	0A246A	
             *	FFFFFF	D4D0C8	808080	808080	000000	D4D0C8	FFFFFF	A6CAF0	
             *	C0C0C0
             *	
             *	Hex colors are RGB.
			*/

			if (fe_styling == null)
				throw new ArgumentNullException("fe_styling");

			if (string.IsNullOrWhiteSpace(fe_styling))
				throw new ArgumentException("Styling string is empty.", "fe_styling");

			string[] items = fe_styling.Split('\t');

			Logging.WriteLineToLog("FrontendStyling: Number of elements in fe_styling: {0:n0}", items.Length);

			if (items.Length != 25)
				throw new ArgumentOutOfRangeException("fe_styling");

			DesiredFont = new Font(items[0], float.Parse(items[1],
				NumberStyles.Integer | NumberStyles.AllowDecimalPoint,
				CultureInfo.InvariantCulture));

			int[] elements = new int[]
			{
				COLOR_SCROLLBAR,			// items[2]
				COLOR_BACKGROUND,			// items[3]
				COLOR_ACTIVECAPTION,		// ...
				COLOR_INACTIVECAPTION,
				COLOR_MENU,
				COLOR_WINDOW,
				COLOR_WINDOWFRAME,
				COLOR_MENUTEXT,
				COLOR_WINDOWTEXT,
				COLOR_CAPTIONTEXT,
				COLOR_ACTIVEBORDER,
				COLOR_INACTIVEBORDER,
				COLOR_APPWORKSPACE,
				COLOR_HIGHLIGHT,
				COLOR_HIGHLIGHTTEXT,
				COLOR_BTNFACE,
				COLOR_BTNSHADOW,
				COLOR_GRAYTEXT,
				COLOR_BTNTEXT,
				COLOR_INACTIVECAPTIONTEXT,
				COLOR_BTNHIGHLIGHT,
				COLOR_2NDACTIVECAPTION,
				COLOR_2NDINACTIVECAPTION	// items[23]
			};

			uint[] values = new uint[]
			{
				// Not very pretty or efficient, but convenient.
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[2])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[3])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[4])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[5])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[6])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[7])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[8])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[9])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[10])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[11])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[12])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[13])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[14])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[15])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[16])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[17])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[18])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[19])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[20])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[21])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[22])),
				(uint)ColorTranslator.ToWin32(ColorTranslator.FromHtml('#' + items[23])),
			};

			bool ret = SetSysColors(elements.Length, elements, values);

			Logging.WriteLineToLog("FrontendStyling: SetSysColors result: {0}", ret);
		}
	}
}