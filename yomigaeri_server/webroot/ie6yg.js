// This file is actually a workaround for the stupid EOLAS patent issue that
// required ActiveX controls to be clicked on before they start working. This
// behavior was later removed again, but that won't help anyone using IE6YG
// that is still affected by it.

// NOTE: This is a special template file read as text in UTF8 format!

document.write('<object classid="clsid:D322D3BD-AF48-4787-ACA6-2D32F2A59A32" id="yomigaeri" width="100%" height="100%">');
document.write('<param name="DebugLog" value="%FRONTENDDEBUG%">');
document.write('<param name="RDP_Server" value="%RDPSERVER%">');
document.write('<param name="RDP_Port" value="%RDPPORT%">');
document.write('<param name="RDP_Username" value="%RDPUSERNAME%">');
document.write('<param name="RDP_Password" value="%RDPPASSWORD%">');
document.write('<param name="RDP_Shell" value="%RDPSHELL%">');
document.write('<param name="Download_Server" value="%DOWNLOADSERVER%">');
document.write('<param name="Download_Port" value="%DOWNLOADPORT%">');
document.write('</object>');