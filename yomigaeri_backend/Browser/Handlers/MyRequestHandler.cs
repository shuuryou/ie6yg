using CefSharp;
using System;
using System.Globalization;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using yomigaeri_shared;
using static yomigaeri_backend.Browser.SynchronizerState;

namespace yomigaeri_backend.Browser.Handlers
{
	internal sealed class MyRequestHandler : CefSharp.Handler.RequestHandler
	{
		private readonly SynchronizerState m_SyncState;
		private readonly Action m_SyncProc;

		public IRequestCallback SSLCertificate_CurrentErrorCallback { get; private set; }

		public IAuthCallback Authentication_CurrentAuthCallback { get; private set; }

		public MyRequestHandler(SynchronizerState syncState, Action syncProc)
		{
			m_SyncState = syncState ?? throw new ArgumentNullException("syncState");
			m_SyncProc = syncProc ?? throw new ArgumentNullException("syncProc");

			SSLCertificate_CurrentErrorCallback = null;
		}

		protected override bool GetAuthCredentials(IWebBrowser chromiumWebBrowser, IBrowser browser, string originUrl, bool isProxy, string host, int port, string realm, string scheme, IAuthCallback callback)
		{
			Logging.WriteLineToLog("GetAuthCredentials: yes");

			Authentication_CurrentAuthCallback = callback;

			m_SyncState.AuthenticationPrompt = string.Format(CultureInfo.InvariantCulture, "{0}\x1{1}\x1{2}", host, realm, originUrl);

			Logging.WriteLineToLog("GetAuthCredentials: {0}", m_SyncState.AuthenticationPrompt);

			m_SyncProc.Invoke();

			return true;
		}

		protected override bool OnCertificateError(IWebBrowser chromiumWebBrowser, IBrowser browser, CefErrorCode errorCode, string requestUrl, ISslInfo sslInfo, IRequestCallback callback)
		{
			ProcessSSLInfo(sslInfo.CertStatus, sslInfo.X509Certificate);

			SSLCertificate_CurrentErrorCallback = callback;

			m_SyncState.CertificatePrompt = true;

			m_SyncProc.Invoke();

			return true;
		}

		protected async override void OnDocumentAvailableInMainFrame(IWebBrowser chromiumWebBrowser, IBrowser browser)
		{
			NavigationEntry current = await Program.WebBrowser.GetVisibleNavigationEntryAsync();

			if (!current.SslStatus.IsSecureConnection)
			{
				m_SyncState.CertificateData = null;
				m_SyncState.CertificateState = FrontendCertificateStates.None;
				m_SyncState.SSLIcon = SynchronizerState.SSLIconState.None;
				m_SyncProc.Invoke();
			}
			else
			{
				ProcessSSLInfo(current.SslStatus.CertStatus, current.SslStatus.X509Certificate);
				m_SyncProc.Invoke();
			}

			base.OnDocumentAvailableInMainFrame(chromiumWebBrowser, browser);
		}
		private void ProcessSSLInfo(CertStatus status, X509Certificate2 certificate)
		{
			FrontendCertificateStates state = FrontendCertificateStates.None;

			bool have_badness = false;

			// Here come the bad things

			if (status.HasFlag(CertStatus.AuthorityInvalid))
			{
				state |= FrontendCertificateStates.UntrustedIssuer;
				have_badness = true;
			}
			else if (status.HasFlag(CertStatus.Revoked))
			{
				state |= FrontendCertificateStates.Revoked;
				have_badness = true;
			}
			else if (status.HasFlag(CertStatus.CtComplianceFailed))
			{
				state |= FrontendCertificateStates.ChromeCTFail;
				have_badness = true;
			}
			else
			{
				state |= FrontendCertificateStates.TrustedIssuer;
			}

			if (status.HasFlag(CertStatus.CommonNameInvalid) ||
				status.HasFlag(CertStatus.NameConstraintViolation) ||
				status.HasFlag(CertStatus.NonUniqueName))
			{
				state |= FrontendCertificateStates.NameInvalid;
				have_badness = true;
			}
			else
			{
				state |= FrontendCertificateStates.NameValid;
			}

			if (status.HasFlag(CertStatus.DateInvalid))
			{
				state |= FrontendCertificateStates.DateInvalid;
				have_badness = true;
			}
			else if (status.HasFlag(CertStatus.ValidityTooLong))
			{
				state |= FrontendCertificateStates.DateInvalidTooLong;
				have_badness = true;
			}
			else
			{
				state |= FrontendCertificateStates.DateValid;
			}


			if (status.HasFlag(CertStatus.Sha1SignaturePresent) ||
			status.HasFlag(CertStatus.WeakSignatureAlgorithm) ||
			status.HasFlag(CertStatus.WeakKey))
			{
				state |= FrontendCertificateStates.WeakCert;
				have_badness = true;
			}
			else
			{
				state |= FrontendCertificateStates.StrongCert;
			}

			if (have_badness)
			{
				state |= FrontendCertificateStates.OverallBad;
				m_SyncState.SSLIcon = SynchronizerState.SSLIconState.SecureBadCert;
			}
			else
			{
				state |= FrontendCertificateStates.OverallOK;
				m_SyncState.SSLIcon = SSLIconState.Secure;
			}

			m_SyncState.CertificateState = state;

			m_SyncState.CertificateData = MakePEM(certificate);
		}
		private byte[] MakePEM(X509Certificate2 certificate)
		{
			return Encoding.ASCII.GetBytes("-----BEGIN CERTIFICATE-----\r\n" + 
				Convert.ToBase64String(certificate.RawData) + 
				"\r\n-----END CERTIFICATE-----\r\n");
		}

	}
}