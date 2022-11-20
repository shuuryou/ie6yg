﻿using System;
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

		private enum ReplyDownloadResult
		{
			DownloadSuccess = 0,
			ReplyWithFailure = 1,
			Bail = 2
		}

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

#if DEBUG
			m_HttpServerThread.Join();
#endif
		}

		public void Terminate()
		{
			m_Cancel.Cancel();
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
					Logging.WriteLineToLog("HttpServerMain: HTTP socket shut down.", e);
				}

				Logging.WriteLineToLog("HttpServerMain: HTTP server socket error: {0}", e);
			}
			catch (ThreadAbortException)
			{
				Logging.WriteLineToLog("HttpServerMain: HTTP server main thread aborted.");
				return;
			}

			if (token.IsCancellationRequested)
			{
				Logging.WriteLineToLog("HttpServerMain: HTTP server main thread exiting cleanly.");
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

			Logging.WriteLineToLog("HttpAnswerConnection: Start processing connection from {0}", client_addr);

			try
			{
				using (NetworkStream ns = client.GetStream())
				using (StreamReader sr = new StreamReader(ns))
				using (StreamWriter sw = new StreamWriter(ns))
				{
					string request_uri;

					{
						string header = sr.ReadLine();

						if (string.IsNullOrEmpty(header))
							return; // Browser vanished without sending headers

						if (!header.StartsWith("GET ", StringComparison.Ordinal))
							goto badrequest;

						int idx = header.LastIndexOf(' ');

						if (idx == -1)
							request_uri = header.Substring(4);
						else
							request_uri = header.Substring(4, idx - 4);

						while (!string.IsNullOrEmpty(header))
						{
							// Skip any other headers that the browser sends
							header = sr.ReadLine();
						}
					}

					if (request_uri.StartsWith("/dl/", StringComparison.Ordinal))
					{
						ReplyDownloadResult result = ReplyDownload(request_uri, sw);

						switch (result)
						{
							case ReplyDownloadResult.DownloadSuccess:
								// Everything went well
								goto done;
							case ReplyDownloadResult.Bail:
								// Socket error during download (client vanished/cancelled download)
								goto bail;
							case ReplyDownloadResult.ReplyWithFailure:
								// Bad random_id from the client or bad data from the backend
								goto servererror;
						}
					}

					// Arbitrary files. Dangerous. :-(

					if (string.IsNullOrWhiteSpace(request_uri) || request_uri == "/")
						request_uri = "/index.html";

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

					string request_file = m_ContentDir + Path.DirectorySeparatorChar + request_uri;
					request_file = Path.GetFullPath(request_file);

					// Prevent path traversal attacks
					if (!request_file.StartsWith(m_ContentDir, StringComparison.OrdinalIgnoreCase))
					{
						Logging.WriteLineToLog("HttpAnswerConnection: Path traversal attack blocked for {0}",
							client_addr);

						goto notfound;
					}

					if (!File.Exists(request_file))
					{
						Logging.WriteLineToLog("HttpAnswerConnection: Cannot find file \"{0}\" on disk for {1}",
							request_file, client_addr);

						goto notfound;
					}

					ReplyRequest(200, "OK", sw, request_file);
					goto done;

				badrequest:
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 401 Bad Request for {0}", client_addr);
					ReplyRequest(401, "Bad Request", sw);
					goto done;

				notfound:
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 404 Not Found for {0}", client_addr);

					ReplyRequest(404, "Not Found", sw);
					goto done;

				servererror:
					Logging.WriteLineToLog("HttpAnswerConnection: Respond with 500 Internal Server Error for {0}", client_addr);

					ReplyRequest(500, "Internal Server Error", sw);
					goto done;

				done:
					ns.Flush();
					client.Close();

				bail:
					;
				}
			}
			catch (Exception e)
			{
				Logging.WriteLineToLog("HttpAnswerConnection: Error processing connection for {0}: {1}", client_addr, e);
			}

			Logging.WriteLineToLog("HttpAnswerConnection: End processing connection for {0}", client_addr);
		}

		private ReplyDownloadResult ReplyDownload(string url, StreamWriter output)
		{
			string random_id = url.Substring(4);

			Logging.WriteLineToLog("ReplyDownload: Download request for download ID {0}", random_id);

			string download_meta = Path.GetFullPath(Path.Combine(Program.Settings.Get("Downloads", "DownloadTempDir"), random_id)) + ".ini";

			Logging.WriteLineToLog("ReplyDownload: Download {0} meta: {1}", random_id, download_meta);

			if (!File.Exists(download_meta))
			{
				Logging.WriteLineToLog("ReplyDownload: Download ID {0}: download meta not found.", random_id);
				return ReplyDownloadResult.ReplyWithFailure;
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
			 * 
			 * Not all of them are used (yet) below.
			 */

			string download_file = reader.Get("Download", "LocalFile");

			if (!File.Exists(download_file))
			{
				Logging.WriteLineToLog("ReplyDownload: Download ID {0}: download file not found.", random_id);
				return ReplyDownloadResult.ReplyWithFailure;
			}

			long total_bytes;

			{
				bool flag = long.TryParse(reader.Get("Download", "TotalBytes"),
					NumberStyles.None, CultureInfo.InvariantCulture, out total_bytes);

				if (!flag)
				{
					Logging.WriteLineToLog("ReplyDownload: Download ID {0}: Could not get file size of download.", random_id);
					return ReplyDownloadResult.ReplyWithFailure;
				}
			}

			string mime_type = reader.Get("Download", "MimeType", "application/octet-stream");
			string content_disposition = reader.Get("Download", "ContentDisposition");
			string suggested_file_name = reader.Get("Download", "SuggestedFileName");

			if (string.IsNullOrEmpty(content_disposition))
			{
				if (string.IsNullOrEmpty(suggested_file_name))
					content_disposition = "attachment";
				else
					content_disposition = string.Format(CultureInfo.InvariantCulture, "attachment; filename=\"{0}\"", suggested_file_name);
			}

			/*
			 * Finished with all of the preparations, so send headers.
			 */

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

			/*
			 * Send the actual binary file contents next.
			 */

			// If unable to write to the downloading browser for 30 seconds,
			// will throw an IOException and in the loop below and cause the
			// method to bail.
			output.BaseStream.WriteTimeout = 30000;

			short bail = 0;

			using (FileStream fs = new FileStream(download_file, FileMode.Open,
				FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.SequentialScan))
			{
				fs.Seek(0, SeekOrigin.Begin);

				long copied_bytes = 0;
				byte[] buf = new byte[8192];
				int read;

				// Next needs to be an infinite loop rather than Stream.CopyTo
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
						try
						{
							output.BaseStream.Write(buf, 0, read);
						}
						catch (IOException)
						{
							// Client probably vanished (download cancelled),
							// or timeout writing to client (DOS attack?)

							// Deleting download_meta should make the backend
							// cancel the download and delete download_file.

							File.Delete(download_meta);

							return ReplyDownloadResult.Bail;
						}

						copied_bytes += read;
						bail = 0;
					}
					else
					{
						// Stall while waiting for more data to arrive in the
						// download file.

						Thread.Sleep(500);

						if (++bail == 600)
						{
							// Something is wrong if there is no data arriving
							// in download_file within 5 minutes.

							Logging.WriteLineToLog("ReplyDownload: Breaking out of infinite loop.");

							// Deleting download_meta should make the backend
							// cancel the download and delete download_file.

							File.Delete(download_meta);

							return ReplyDownloadResult.Bail;
						}
					}

					if (total_bytes != 0 && copied_bytes == total_bytes)
					{
						// This will handle the majority of the cases.

						break;
					}
					else if (total_bytes == 0 && !File.Exists(download_meta))
					{
						// If the size of the download isn't known, the backend
						// will delete the meta file once the download is done.
						// See the comments in MyDownloadHandler.cs about this
						// poor man's IPC.

						break;
					}
				}
			}

			/*
			 * Getting here means the download completed successfully.
			 */

			output.BaseStream.Flush();

			File.Delete(download_meta);
			File.Delete(download_file);

			return ReplyDownloadResult.DownloadSuccess;
		}

		private void ReplyRequest(int response_code, string response_name, StreamWriter output)
		{
			output.WriteLine(string.Format(CultureInfo.InvariantCulture, "HTTP/1.0 {0} {1}", response_code, response_name));
			output.WriteLine("Content-Type: text/html");
			output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Content-Length: {0}", response_name.Length));

			output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Date: {0:R}", DateTime.Now));
			output.WriteLine("Server: IE6YG Minimal HTTP Server");
			output.WriteLine("Connection: Close");
			output.WriteLine();

			output.Write(response_name);

			output.Flush();
		}

		private void ReplyRequest(int response_code, string response_name, StreamWriter output, string content_file)
		{
			using (FileStream fs = File.OpenRead(content_file))
			{
				output.WriteLine(string.Format(CultureInfo.InvariantCulture, "HTTP/1.0 {0} {1}", response_code, response_name));
				output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Content-Type: {0}", GetMimeType(content_file)));
				output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Content-Length: {0}", fs.Length));
				output.WriteLine(string.Format(CultureInfo.InvariantCulture, "Date: {0:R}", DateTime.Now));
				output.WriteLine("Server: IE6YG Minimal HTTP Server");
				output.WriteLine("Connection: Close");
				output.WriteLine();
				output.Flush();

				fs.CopyTo(output.BaseStream);

				output.BaseStream.Flush();
			}
		}

		private string GetMimeType(string fileName)
		{
			const string MIME_TYPE_DEFAULT = "application/unknown";

			string ext = Path.GetExtension(fileName).ToLower();

			if (string.IsNullOrEmpty(ext))
				return MIME_TYPE_DEFAULT;

			if (ext[0] == '.')
				ext = ext.Substring(1);

			return ApacheMimeTypes.MimeTypes.ContainsKey(ext) ?
				ApacheMimeTypes.MimeTypes[ext] : MIME_TYPE_DEFAULT;
		}
	}
}