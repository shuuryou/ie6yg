using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace yomigaeri_backend
{
	/// <summary>
	/// Provides RDP Virtual Channels communication between YOMIGAERIFE ActiveX
	/// control and YOMIGAERIBE application.
	/// </summary>
	/// <remarks>
	/// Ports just enough of WTSAPI32.H for communication to work. The imports,
	/// esp. the use of CharSet.Ansi, are not a mistake.</remarks>
	internal static class RDPVirtualChannel
	{

		// #define CHANNEL_NAME_LEN 7
		private const string VIRTUAL_CHANNEL_NAME = "BEFECOM";

		#region WTSAPI32.H Subset
		private const uint WTS_CURRENT_SESSION = uint.MaxValue;
		private const uint CHANNEL_CHUNK_LENGTH = 1600;

		[Flags]
		private enum WTSVirtualChannelOpenExFlags
		{
			WTS_CHANNEL_OPTION_DYNAMIC = 0x00000001,
			WTS_CHANNEL_OPTION_DYNAMIC_PRI_LOW = 0x00000000,
			WTS_CHANNEL_OPTION_DYNAMIC_PRI_MED = 0x00000002,
			WTS_CHANNEL_OPTION_DYNAMIC_PRI_HIGH = 0x00000004,
			WTS_CHANNEL_OPTION_DYNAMIC_PRI_REAL = 0x00000006,
			WTS_CHANNEL_OPTION_DYNAMIC_NO_COMPRESS = 0x00000008
		}

		[DllImport("wtsapi32.dll", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
		private static extern IntPtr WTSVirtualChannelOpenEx(uint dwSessionID, string pChannelName, WTSVirtualChannelOpenExFlags flags);

		[DllImport("wtsapi32.dll", SetLastError = true)]
		private static extern bool WTSVirtualChannelClose([In] IntPtr channelHandle);

		[DllImport("wtsapi32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool WTSVirtualChannelRead(IntPtr hChannelHandle, uint Timeout, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 3)] byte[] Buffer, uint BufferSize, out uint BytesRead);

		[DllImport("wtsapi32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool WTSVirtualChannelWrite(IntPtr channelHandle, byte[] buffer, uint length, ref uint bytesWritten);
		#endregion

		private static IntPtr s_hChannel;
		static RDPVirtualChannel()
		{
			s_hChannel = IntPtr.Zero;
		}

		public static void OpenChannel()
		{
			Logging.WriteLineToLog("RDPVirtualChannel OPEN: ENTRY");

			if (s_hChannel != IntPtr.Zero)
				return;

			// MSDN:
			// DWORD flags: To open the channel as an SVC, specify zero for this parameter.

			IntPtr hChannel = WTSVirtualChannelOpenEx(WTS_CURRENT_SESSION, VIRTUAL_CHANNEL_NAME, 0);

			if (hChannel == IntPtr.Zero)
				throw new Win32Exception();

			s_hChannel = hChannel;

			Logging.WriteLineToLog("RDPVirtualChannel OPEN: s_hChannel becomes 0x{0:x}", s_hChannel);
		}

		public static void CloseChannel()
		{
			Logging.WriteLineToLog("RDPVirtualChannel CLOSE: ENTRY");

			if (s_hChannel == IntPtr.Zero)
				return;

			bool ok = WTSVirtualChannelClose(s_hChannel);

			if (!ok)
				throw new Win32Exception();

			Logging.WriteLineToLog("RDPVirtualChannel CLOSE: END");
		}

		public static void Reset()
		{
			s_hChannel = IntPtr.Zero;
		}

		public static string ReadChannelUntilResponse(int timeout = 5)
		{
			// Not proud of this, but it will have to do until something better
			// comes along.

			DateTime start = DateTime.Now;

			again:
			string ret = ReadChannel(false);

			if (ret.Length == 0)
			{
				if (timeout != Timeout.Infinite && (DateTime.Now - start).TotalSeconds > timeout)
					throw new TimeoutException();

				goto again;
			}

			return ret;
		}

		public static string ReadChannel(bool noDelay = false)
		{
			Logging.WriteLineToLog("RDPVirtualChannel READ: ENTRY");

			if (s_hChannel == IntPtr.Zero)
				throw new InvalidOperationException("RDP virtual channel is closed.");

			byte[] buffer = new byte[CHANNEL_CHUNK_LENGTH];

			bool ok = WTSVirtualChannelRead(s_hChannel, noDelay ? 0U : 1000U,
				buffer, CHANNEL_CHUNK_LENGTH, out uint bytesRead);

			if (!ok)
				throw new Win32Exception();

			string ret;

			if (bytesRead == 0)
			{
				ret = string.Empty;
				goto done;
			}

			// No, this is not wrong. Strings from YOMIGAERIFE ActiveX
			// control are Unicode string. That's how VB6 works...
			ret = Encoding.Unicode.GetString(buffer, 0, (int)bytesRead);

		done:
			Logging.WriteLineToLog("RDPVirtualChannel READ: \"{0}\"", ret);
			return ret;
		}
		public static void WriteChannel(string message)
		{
			Logging.WriteLineToLog("RDPVirtualChannel WRITE: ENTRY");

			if (s_hChannel == IntPtr.Zero)
				throw new InvalidOperationException("RDP virtual channel is closed.");

			byte[] buf = Encoding.Unicode.GetBytes(message);
			uint written = 0;
			bool ok = WTSVirtualChannelWrite(s_hChannel, buf, (uint)buf.Length, ref written);

			if (!ok)
				throw new Win32Exception();

			if (written != buf.Length)
				throw new IOException(string.Format(CultureInfo.InvariantCulture,
					"Operation failed. Should write {0:n0} bytes, but wrote {1:n0} bytes.", buf.Length, written));

			Logging.WriteLineToLog("RDPVirtualChannel WRITE: \"{0}\"", message);
		}
	}
}