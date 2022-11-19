using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using yomigaeri_shared;

namespace yomigaeri_server
{
	public sealed class Server
	{
		private readonly TcpListener m_HttpServer;
		private readonly Thread m_HttpServerThread;
		private const int HTTP_SERVER_MAX_BACKLOG = 25;

		private readonly string m_ContentDir;

		private readonly CancellationTokenSource m_Cancel;

		public Server(IPAddress listenAddr, int listenPort, string contentDir)
		{
			if (listenAddr == null)
				throw new ArgumentNullException("listenAddr");

			if (contentDir == null)
				throw new ArgumentNullException("contentDir");

			if (listenPort < 0)
				throw new ArgumentOutOfRangeException("listenPort");

			if (!Directory.Exists(contentDir))
				throw new FileNotFoundException("Content directory does not exist.", contentDir);

			m_ContentDir = Path.GetFullPath(contentDir);

			m_HttpServer = new TcpListener(listenAddr, listenPort);
			m_HttpServerThread = new Thread(HttpServerMain);


			m_Cancel = new CancellationTokenSource();
		}

		public void Begin()
		{
			m_HttpServer.Start(HTTP_SERVER_MAX_BACKLOG);
			m_HttpServerThread.Start(m_Cancel.Token);

			m_HttpServerThread.Join();
		}

		private void HttpServerMain(object cancelToken)
		{
			CancellationToken token = (CancellationToken)cancelToken;

			Logging.WriteLineToLog("HttpServerMain: Entry.");

		again:

			try
			{
				Logging.WriteLineToLog("HttpServerMain: Waiting for connection.");
				TcpClient client = m_HttpServer.AcceptTcpClient();

				Logging.WriteLineToLog("HttpServerMain: Accepted a connection.");
				ThreadPool.QueueUserWorkItem(HttpAnswerConnection, client);
			}
			catch (SocketException e)
			{

				if ((e.SocketErrorCode == SocketError.Interrupted))
				{
					Logging.WriteLineToLog("HttpServerMain: HTTP socket shut down.", e.ToString());
				}

				Logging.WriteLineToLog("HttpServerMain: HTTP server socket error: {0}", e.ToString());
			}
			catch (ThreadAbortException)
			{
				Logging.WriteLineToLog("HttpServerMain: HTTP server main thread aborted.");
				return;
			}

			if (token.IsCancellationRequested)
			{
				Logging.WriteLineToLog("HTTP server main thread exiting cleanly.");
				return;
			}

			goto again;
		}

		private void HttpAnswerConnection(object state)
		{
			if (!(state is TcpClient client))
				return;

			if (!client.Connected)
				return;

			IPAddress client_addr = ((IPEndPoint)client.Client.RemoteEndPoint).Address;

			Logging.WriteLineToLog("HttpAnswerConnection: Start processing connection from {0}.", client_addr);

			try
			{
				using (NetworkStream ns = client.GetStream())
				using (StreamReader sr = new StreamReader(ns))
				using (StreamWriter sw = new StreamWriter(ns))
				{
					string request = sr.ReadLine();

					// Deal with the other end hanging up the phone without
					// talking.
					if (string.IsNullOrEmpty(request))
						return;

					string[] parts = request.Split(' ');

					while (!string.IsNullOrEmpty(request))
					{
						// Skip the other crap. I'm not going to implement the
						// entire HTTP specification or create the next nginx
						// just for this project.
						request = sr.ReadLine();
					}

					if (parts.Length != 3 || parts[0] != "GET" || !parts[2].StartsWith("HTTP/1", StringComparison.Ordinal))
					{
						Logging.WriteLineToLog("HttpAnswerConnection: Respond with 401 Bad Request for request: {0}", request);
						ReplyRequest(401, "Bad Request", sw);
						goto done;
					}

					if (parts[1].StartsWith("/dl/", StringComparison.Ordinal))
					{
						if (ReplyDownload(parts[1], sw))
							goto done;
						else
							goto servererror;
					}

					// Arbitrary files. Dangerous. :-(

					if (parts[1] == "/")
						parts[1] = "/index.html";

					/* DANGER DANGER DANGER
					 * 
					 * Do not use Path.Combine here!
					 * 
					 * See MSDN:
					 * 
					 * "If an argument other than the first contains a rooted
					 * path, any previous path components are ignored, and the
					 * returned string begins with that rooted path component."
					 * 
					 * Path.Combine("C:\\test\\", "/Windows/System32/calc.exe")
					 * ==> C:\Windows\System32\calc.exe
					 *
					 * *sigh*
					 */

					string file = m_ContentDir + Path.DirectorySeparatorChar + parts[1];
					file = Path.GetFullPath(file);

					// Prevent path traversal attacks
					if (!file.StartsWith(m_ContentDir, StringComparison.OrdinalIgnoreCase))
					{
						Logging.WriteLineToLog("HttpAnswerConnection: Path traversal attack blocked for request: {0}", request);
						goto notfound;
					}

					if (!File.Exists(file))
					{
						Logging.WriteLineToLog("HttpAnswerConnection: Cannot find file \"{0}\" on disk for request: {1}", file, request);
						goto notfound;
					}

					ReplyRequest(200, "OK", sw, File.ReadAllText(file));
					goto done;

				notfound:
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 404 Not Found for request: {0}", request);

					ReplyRequest(404, "Not Found", sw);
					goto done;

				servererror:
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 500 Internal Server Error for request: {0}", request);

					ReplyRequest(500, "Internal Server Error", sw);
					goto done;

				done:
					ns.Flush();
					client.Close();
					return;

				}
			}
			catch (Exception e)
			{
				Logging.WriteLineToLog("HttpAnswerConnection: Error processing connection: {0}", e);
			}

			Logging.WriteLineToLog("HttpAnswerConnection: End processing connection from {0}.", client_addr);
		}

		private bool ReplyDownload(string url, StreamWriter output)
		{
			string random_id = url.Substring(4);

			Logging.WriteLineToLog("ReplyDownload: Download request for download ID {0}", random_id);

			string download_meta = Path.GetFullPath(Path.Combine(Program.Settings.Get("Downloads", "DownloadTempDir"), random_id)) + ".ini";

			Logging.WriteLineToLog("ReplyDownload: Download {0} meta: {1}", random_id, download_meta);

			if (!File.Exists(download_meta))
			{
				Logging.WriteLineToLog("ReplyDownload: Download ID {0}: download meta not found.", random_id);
				return false;
			}

			IniFileReader reader = new IniFileReader(download_meta);

			/*
			 * "Download", "CEFID", downloadItem.Id
			 * "Download", "TotalBytes", downloadItem.TotalBytes
			 * "Download", "MimeType", downloadItem.MimeType
			 * "Download", "ContentDisposition", downloadItem.ContentDisposition
			 * "Download", "URL", downloadItem.Url
			 * "Download", "SuggestedFileName", downloadItem.SuggestedFileName
			 * "Download", "LocalFile", download_file
			 */

			string download_file = reader.Get("Download", "LocalFile");

			if (!File.Exists(download_file))
			{
				Logging.WriteLineToLog("ReplyDownload: Download ID {0}: download file not found.", random_id);
				return false;
			}

			long total_bytes;

			{
				bool flag = long.TryParse(reader.Get("Download", "TotalBytes"),
					NumberStyles.None, CultureInfo.InvariantCulture, out total_bytes);

				if (!flag)
				{
					Logging.WriteLineToLog("ReplyDownload: Download ID {0}: Could not get file size of download.", random_id);
					return false;
				}
			}

			string mime_type = reader.Get("Download", "MimeType", "application/octet-stream");
			string content_disposition = reader.Get("Download", "ContentDisposition");
			string suggested_file_name = reader.Get("Download", "SuggestedFileName");

			if (string.IsNullOrEmpty(content_disposition))
				content_disposition = string.Format(CultureInfo.InvariantCulture, "attachment; filename=\"{0}\"", suggested_file_name);

			output.WriteLine("HTTP/1.0 200 OK");
			output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Content-Type: {0}", mime_type));

			if (total_bytes != 0)
				output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Content-Length: {0}", total_bytes));

			output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Content-Disposition: {0}", content_disposition));
			output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Date: {0:R}", DateTime.Now));
			output.WriteLine("Connection: Close");
			output.WriteLine("Server: IE6YG Minimal HTTP Server");
			output.WriteLine();
			output.Flush();

			// From here, no longer use StreamWriter; only use its BaseStream.

			short bail = 0;

			using (FileStream fs = new FileStream(download_file, FileMode.Open,
				FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.SequentialScan))
			{
				fs.Seek(0, SeekOrigin.Begin);

				long copied_bytes = 0;
				byte[] buf = new byte[8192];
				int read;

				// This needs to be an infinite loop rather than Stream.CopyTo
				// or whatever other cool way you normally use with streams,
				// because here we are reading a file that Chromium is still
				// actively writing to. The file is still growing while being
				// read from here!

				// It gets even worse when total_bytes is 0, because that means
				// the server Chromium is dowloading from doesn't know the size
				// of the download (e.g. it's being generated on the fly).

				for (; ; )
				{
					read = fs.Read(buf, 0, buf.Length);

					if (read > 0)
					{
						output.BaseStream.Write(buf, 0, read);
						copied_bytes += read;
						bail = 0;
					}
					else
					{
						// Stall while waiting for more data to arrive in the
						// download file.

						Thread.Sleep(500);
						bail++;

						if (bail > short.MaxValue)
						{
							Logging.WriteLineToLog("ReplyDownload: Breaking out of infinite loop.");

							// Can't return false, because that would cause the
							// a HTTP status error to be barfed onto the end of
							// whatever was downloaded so far.
							return true;
						}
					}

					// The first check is obvious. If the size of the download 
					// isn't known, the backend will delete the meta file once
					// it thinks the download is compete. See the comments in
					// MyDownloadHandler.cs about this poor man's IPC.

					if (total_bytes != 0 && copied_bytes == total_bytes)
						return true;
					else if (total_bytes == 0 && !File.Exists(download_meta))
						return true;
				}
			}
		}

		private void ReplyRequest(int response_code, string response_name, StreamWriter output, string content = null)
		{
			output.WriteLine(string.Format(CultureInfo.InvariantCulture, "HTTP/1.0 {0} {1}", response_code, response_name));
			output.WriteLine("Content-Type: text/html");

			if (content == null)
				output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Content-Length: {0}", response_name.Length));
			else
				output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Content-Length: {0}", content.Length));

			output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Date: {0:R}", DateTime.Now));
			output.WriteLine("Server: IE6YG Minimal HTTP Server");
			output.WriteLine("Connection: Close");
			output.WriteLine();

			if (content == null)
				output.Write(response_name);
			else
				output.Write(content);

			output.Flush();
		}
		private class WebClientEx : WebClient
		{
			public CookieContainer CookieContainer { get; set; } = new CookieContainer();

			protected override WebRequest GetWebRequest(Uri uri)
			{
				WebRequest request = base.GetWebRequest(uri);
				if (request is HttpWebRequest)
				{
					(request as HttpWebRequest).CookieContainer = CookieContainer;
				}
				return request;
			}
		}
	}
}