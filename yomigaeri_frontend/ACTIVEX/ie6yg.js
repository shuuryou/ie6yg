// Workaround for stupid EOLAS patent that requires ActiveX controls to be clicked
// on before they start working. This was later removed again, but that won't help
// anyone using IE6YG that is affected.

document.write('<object classid="clsid:D322D3BD-AF48-4787-ACA6-2D32F2A59A32" id="yomigaeri" width="100%" height="100%">');
document.write('<param name="DebugLog" value="True">');
document.write('<param name="RDP_Server" value="browservm">');
document.write('<param name="RDP_Username" value="cefshim1">');
document.write('<param name="RDP_Password" value="cefshim">');
document.write('<param name="RDP_Backend" value="Z:\\yomigaeri_backend.exe">');
document.write('</object>');