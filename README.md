# libTridentUI

**Want to develop a browser-based desktop application for legacy operating systems? Try libTridentUI!**

**-- An extremely lightweight Chrome Embedded Framework alternative for legacy operating systems.**

**-- Works perfectly on Microsoft Windows XP SP3!**

## Why use libTridentUI?

1. Extremely lightweight
   Modern browser based UI frameworks like Electron, CEF, and WebView2 are powerful but extremely heavy.

   They ship a full Chromium browser with your app which would obviously occupy a huge amount of storage space and require modern OS versions.

   If you're building a lightweight desktop utility that needs a rich UI but also needs to run on legacy systems, your options are very limited.

   This is where **libTridentUI** can help you!
   
   Every copy of Windows since Windows 95 ships with mshtml.dll (the Trident/IE engine) that is capable of rendering HTML.

   Not only it's extremely tiny, but it also ships together with Windows so you don't even have to install it additionally! It's already sitting there inside **EVERY** single Windows machine around the world just waiting to be called.

   And that's exactly what **libTridentUI** is doing: An extremely lightweight **single header** library that calls mshtml.dll to render the UI of your software with zero DLL dependencies beyond what Windows already provides!

2. Maximum compatibility ensurance
  libTridentUI is built on Microsoft Internet Explorer’s Trident engine, offering absolute top-tier compatibility!!!

* [x] Windows 2000 SP4
* [x] Windows XP SP3
* [x] Windows Vista SP1
* [x] Windows 7 SP1
* [x] Windows 8
* [x] Windows 8.1
* [x] Windows 10 2004
* [x] Windows 11 21H2
* [ ] Linux with WineHQ (Will be supported in near future)
* [ ] MacOS 10+ with CrossOver (Will be supported in near future)
* [ ] Android 11+ with Winlator (Will be supported in near future)

  You can now create browser based programs that works across all versions of Windows spanning a period of 20 years! Isn't that beautiful?

3. Easy to use
   **libTridentUI** wrapped all of the complex COM interfaces needed to host mshtml.dll into a clean and flat C-Style API that you can quickly learn in just a few hours!
   
   No need to create any class or struct, no need to implement any COM interface yourself, **libTridentUI** has already done all of the heavy lifting for you.

4. Multifunctional
   Other libraries might only provide you a web browser window that can only run the slow as hell JavaScript and nothing else, but **libTridentUI** provides you a fully working IDispatch implementation that allows you to call C++ functions from JS and call JS functions from C++ with zero effort!

   In addition, **libTridentUI** also supports embedding custom ActiveX controls into your webpage! Instead of the 0.1 fps JS Canvas, you now have access to a GDI compatible hWnd that you can render complex graphics with lightning speed using C++ via GDI Plus, DirectX, or OpenGL.
   
   The ActiveX component registration is done completely in RAM, no registry writes, no uninstallation leftover!
