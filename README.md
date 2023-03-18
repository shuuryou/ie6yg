# Internet Explorer 6の蘇り
(Revival of Internet Explorer 6)

Don't worry, it's not revived.

This was an attempt to make Chromium Embedded Framework integrated into Internet Explorer 6 on ancient Windows 98 and Windows ME using the RDP protocol. The way it works is just as insane as you're probably imaging right now. I tried this to prove it could be done, but I eventually hit some very tough to solve issues that caused me to lose interest. In particular, how to deal with multiple windows (each needs its own RDP session, which has its own WinSta and is fully isolated -- how do I communicate with the other running CEF instances???), how to deal with pop-up windows (JavaScript `window.parent`, etc.)

I might pick this up again if I ever feel like it, but it's highly unlikely. It got annoying and I lost interest around the time I started implementing context menus. Please understand the stuff I put on my GitHub is just a hobby. It's okay to lose interest in hobbies. I have a family and work that take up 80% of my time. What I do with the remaining 20% is my decision alone.

**THE SOURCE CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.**

**This source code uses knowledge gained from the leaked Microsoft Windows 2000 source code and the leaked Microsoft Windows XP source code, which are both widely available on GitHub. However, it does not include any leaked Microsoft source code. Accessing or using this source code may be illegal or in breach of contract for certain individuals. I take no responsibility for your actions.**

**If you want to use any of this stuff, please see LICENSE file.**

What sort of works:
* History recorded in local IE browser
* Travel log working (menu inside back and forward button)
* CSS cursors
* Flexible window resize without disconnecting/reconnecting RDP
* SSL validation with IE-style dialog box
* IE style error messages for Chromium errors
* JavaScript alert, input, prompt, onunload, etc. handlers natively on frontend
* File downloads in native IE downloader (via a small and proxy that streams the download to IE while CEF is still downloading it)
* Find in page in native classic Windows search box + Chrome style hiliting
* Basic/Digest authentication with encrypted password saving on local system
* File uploads

What doesn't work:
* Context menus
* Full screen mode
* Support for Flash (Ruffle)
* Adblock (had some hope of using DistillNET, but that project is dead)
* Open in new window
* JavaScript popups
* Proper error handling if the backend crashes

There is no setup program or anything like that. You need deep knowledge of programming in C#, bypassing RDP restrictions in Windows client releases, classic Win32 API, COM, Terminal Services, etc. to get started. I used VB6 for the frontend code, because I felt like challenging/torturing myself at the time. Visual C++ 6.0 would have been a better choice.

Below is what it looked like when it was running:
![Screenshot](https://user-images.githubusercontent.com/36278767/226141275-afe477ea-523a-4a5e-8bb7-a99866a77c88.png)![b](https://user-images.githubusercontent.com/36278767/226141277-71d925cc-4fae-42e2-9589-d3604c2eb17d.png)
![Screenshot](https://user-images.githubusercontent.com/36278767/226141278-4de6a584-742e-4d4e-9de9-37fb5ec4c35f.png)
![Screenshot](https://user-images.githubusercontent.com/36278767/226141280-b2d40988-c17e-4f35-8b23-95b7388a8290.jpg)
![Screenshot](https://user-images.githubusercontent.com/36278767/226141281-5d87273d-33aa-4e93-9f30-13ff3df2b7e6.jpg)
![Screenshot](https://user-images.githubusercontent.com/36278767/226141282-f5597e3d-8001-4e60-8fdc-015302f0e6ac.jpg)
![Screenshot](https://user-images.githubusercontent.com/36278767/226141283-1108d8b9-c6d9-4345-b70b-ca7944c61bec.jpg)

Good bye.


