ISMS Admin Account Audit Comparator
===================================

What this tool does
-------------------
This desktop tool compares two quarterly admin account Excel submissions and creates a self-contained HTML report. It runs offline on your computer. It does not call any cloud service, API, or website.

Files included
--------------
1. isms_audit.py - the tool
2. requirements.txt - required Python packages
3. README.txt - this guide

Setup
-----
1. Install Python 3.10 or newer from https://www.python.org/downloads/
2. During installation, tick the box that says "Add Python to PATH".
3. Open Command Prompt or PowerShell in this folder.
4. Install the required packages:

   pip install -r requirements.txt

Optional screenshot check setup
-------------------------------
The screenshot check uses OCR to read text from screenshot images. This still runs offline, but it needs the free Tesseract OCR engine installed on the computer.

1. Install Tesseract OCR for Windows.
2. Make sure tesseract.exe is available on PATH.
3. Close and reopen Command Prompt or PowerShell after installing it.

If Tesseract is not installed, the normal Excel comparison still works. The report will say that screenshot OCR could not run.

How to run
----------
In Command Prompt or PowerShell, run:

   python isms_audit.py

The tool window will open.

How to use
----------
1. Click "Browse..." beside "Previous quarter file" and select the older Excel workbook.
2. Click "Browse..." beside "Current quarter file" and select the newer Excel workbook.
3. Optional: tick "Enable screenshot check" and select either one screenshot image or a folder of screenshots.
4. Click "Run comparison and save report".
5. Choose where to save the HTML report.
6. If the tool cannot identify a User ID column for a worksheet, it will ask you to map the columns manually. Select the correct User ID column for both files. You can also skip a worksheet if it is not relevant.
7. When the report is generated, the file selections are cleared automatically so old files are not reused by accident.

Screenshot check tips
---------------------
- A folder works best when screenshot filenames match worksheet/system names, for example "ERP.png" for an "ERP" worksheet.
- Supported image files include PNG, JPG, JPEG, BMP, TIFF, and WEBP.
- The tool checks current-quarter workbook users against the OCR text from the matched screenshot.
- Clear, high-resolution screenshots work best.
- OCR can make mistakes if the screenshot is blurry, zoomed out, cropped, dark, or uses tiny text.
- Treat screenshot misses as items to review, not automatic proof that the user is absent.

What the report shows
---------------------
- New users added in the current quarter
- Whether each added user has a Request ID
- Users removed since the previous quarter
- Users whose tracked fields changed
- Systems that appear only in the current file
- Systems that appear only in the previous file
- Users removed from multiple systems
- Optional screenshot verification showing whether each current workbook user was seen in the matched screenshot OCR text

Opening and sharing the report
------------------------------
The report is a single HTML file with all styling and scripts included inside it. It can be opened in a browser and shared as one file.

Printing
--------
Use your browser's Print command. The report expands all system sections automatically for printing.

Troubleshooting
---------------
If Python is not recognised:
- Reinstall Python and make sure "Add Python to PATH" is selected.
- Or try running: py isms_audit.py

If packages are missing:
- Run: pip install -r requirements.txt

If screenshot OCR does not run:
- Confirm Tesseract OCR is installed.
- Confirm tesseract.exe is available on PATH.
- Close and reopen Command Prompt or PowerShell, then run the tool again.

If a workbook will not open:
- Make sure the file is not password protected.
- Close the workbook in Excel and try again.
- Save legacy .xls files as .xlsx if possible.

If the tool asks for column mapping:
- This is normal when worksheet headers are unusual.
- Select the User ID column first, because it is required.
- Map Request ID if you want additions to be checked for request evidence.

If a worksheet is not relevant:
- Use "Skip worksheet" in the mapping window.

Privacy and offline use
-----------------------
The tool reads the selected Excel files and screenshots locally and writes the report locally. It does not send data anywhere.
