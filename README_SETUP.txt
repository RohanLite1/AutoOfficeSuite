========================================================================
           OFFICE AUTOMATION SUITE - STANDALONE SETUP GUIDE
========================================================================

This package contains:
  1. OfficeAutomation.exe       - Standalone native Windows software
  2. extension/                 - Chrome Extension (unpacked folder)
  3. Office_MIS_Extension.zip   - Chrome Extension (zipped archive)

------------------------------------------------------------------------
STEP 1: RUNNING THE APPLICATION ON ANY MACHINE
------------------------------------------------------------------------
- Double-click "OfficeAutomation.exe".
- It opens immediately in its own native desktop application window.
- No Python, no terminal, and no internet connection required.
- Everything runs 100% locally on your computer.
- Your selected Master MIS Excel sheet and settings are saved automatically.

------------------------------------------------------------------------
STEP 2: INSTALLING THE CHROME EXTENSION (ONE-TIME SETUP PER PC)
------------------------------------------------------------------------
Chrome unpacked extensions are PERMANENT. You only need to load it once:

1. Open Google Chrome and enter in the address bar:
   chrome://extensions

2. In the top-right corner, turn ON the "Developer mode" switch.

3. In the top-left corner, click "Load unpacked".

4. Select the "extension" folder inside this folder:
   [Path to folder]\extension

5. Done! 
   Chrome permanently saves this extension. You do NOT have to reload it
   when closing Chrome or restarting Windows. It will always be there.
   (Tip: Click the puzzle piece icon in Chrome toolbar and pin "Office MIS Assistant").

------------------------------------------------------------------------
STEP 3: EXTRACTING FROM UNIVERSAL SOMPO OR ANY PORTAL
------------------------------------------------------------------------
1. Open any claim page on Universal Sompo (or test at http://127.0.0.1:5055/test-portal).
2. You will see a blue floating badge: "📋 Send Claim to MIS".
3. Click it — the claim details are scraped and sent directly into OfficeAutomation.
4. In the Office Automation desktop window, go to "🌐 MIS from Web" and click
   "✅ Append Web Claims to Master Excel".

========================================================================
