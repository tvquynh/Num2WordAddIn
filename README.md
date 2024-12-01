# Num2WordAddIn
#Num2Word Add-In for Microsoft Excel
===================================

This software is licensed as Free and Open Source. You are free to use, modify, and distribute it in compliance with the terms of open-source licensing. This add-in is provided "as-is" without any warranties or guarantees.

Author: tvquynh@gmail.com
URL: https://github.com/tvquynh/Num2WordAddIn

#Installation Requirements:
=========================
Use the Setup.exe file for installation.
- Before installing, please close all instances of Microsoft Excel.

#What this installation includes:
===============================
Once installed, the add-in will automatically integrate into Microsoft Excel. The installation includes two versions:
- ExcelAddIn2-AddIn (for 32-bit Excel)
- ExcelAddIn2-AddIn64 (for 64-bit Excel)

#Installation Directory:
========================
The following files will be installed in the folder:
%LocalAppData%\Programs\Num2WordAddIn
(typically: C:\Users\<YourUsername>\AppData\Local\Programs\Num2WordAddIn)

#How to Use:
============
1. Open Microsoft Excel after installation.
2. Use the following formulas in your Excel sheet:
   - =usd(number) — Converts a number to words in English.
   - =vnd(number) — Converts a number to words in Vietnamese.

#If the Add-In Does Not Automatically Load:
===========================================
1. Open Microsoft Excel.
2. Go to the File menu and select Options.
3. Navigate to Add-ins in the left panel.
4. At the bottom of the window, choose Excel Add-ins from the dropdown and click Go.
5. Click Browse, then navigate to the folder:
   %LocalAppData%\Programs\Num2WordAddIn
   (typically: C:\Users\<YourUsername>\AppData\Local\Programs\Num2WordAddIn)
6. Select the appropriate file:
   - For 32-bit Excel: ExcelAddIn2-AddIn.xll
   - For 64-bit Excel: ExcelAddIn2-AddIn64.xll
7. Click OK to enable the add-in.

Your Num2Word Add-In is now ready for use!
