// This file is actually a workaround for the stupid EOLAS patent issue that
// required ActiveX controls to be clicked on before they start working. This
// behavior was later removed again, but that won't help anyone using IE6YG
// that is still affected by it.

// NOTE: This is a special template file read as text in UTF8 format!

document.write('<object classid="clsid:92E4666C-1BEE-4379-B1B8-14AA72B92407" id="yomigaeri" width="100%" height="100%">');
document.write('<param name="DebugLog" value="%FRONTENDDEBUG%">');
document.write('<param name="RDP_Server" value="%RDPSERVER%">');
document.write('<param name="RDP_Port" value="%RDPPORT%">');
document.write('<param name="RDP_Username" value="%RDPUSERNAME%">');
document.write('<param name="RDP_Password" value="%RDPPASSWORD%">');
document.write('<param name="RDP_Shell" value="%RDPSHELL%">');
document.write('<param name="Download_Server" value="%DOWNLOADSERVER%">');
document.write('<param name="Download_Port" value="%DOWNLOADPORT%">');
document.write('</object>');