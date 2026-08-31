ADDRESSMEND — SIMPLE DESKTOP INSTRUCTIONS
=========================================

Author and copyright holder: Connor Baird
Copyright (C) 2026 Connor Baird
Licence: GNU General Public License version 3 or later (GPL-3.0-or-later)

The complete licence terms are supplied in LICENSE.txt.

You do not need administrator access. Python 3.10 or newer must already be
available on the computer.

FIRST USE
---------

1. Keep these files together in the same folder:
     Start_AddressMend.cmd
     Start_AddressMend.sh
     addressmend.py
     READ_ME_FIRST.txt
     LICENSE.txt

2. On Windows, double-click Start_AddressMend.cmd.

   On Linux, open a terminal in this folder and run:

       chmod +x Start_AddressMend.sh
       ./Start_AddressMend.sh

   This does not require administrator rights.

3. Choose option 1: "Paste entries into this window".

4. Paste the complete table into the window. Use Ctrl+V on Windows or
   Ctrl+Shift+V in most Linux terminals; right-click and Paste may also work.

5. After the final row, type DONE on a new line and press Enter.

6. The standard procedure uses internet lookup automatically. Address lookup
   sends postcodes only to Doogal and postcodes.io. Uncommon email-domain
   checking sends only the text after @ to Google Public DNS. Names and complete
   email addresses are not sent to those services. Choose main-menu option 9
   before processing if you need a completely local-only batch. An address is
   sent to OpenStreetMap/Nominatim only when it has no usable postcode.

7. The programme explains each correction. When it finishes, select cell A1 in
   Excel or LibreOffice Calc and press Ctrl+V.

   If Calc cannot read the clipboard, open the saved .tsv file and select Tab as
   the separator and Unicode (UTF-8) as the character set.

FILES CREATED
-------------

Results are saved in your Documents folder, inside:

    AddressMend Results

Each batch creates:

* cleaned_entries_....tsv — open this in Excel/Calc or paste its rows there;
* review_report_....tsv — explains changes and marks anything needing review.

The programme does not guess when the evidence is ambiguous. Check every item
marked "unresolved" or "review" in the review report. A "review" result is a
provisional harmonisation, such as the only street in a postcode when the exact
house number was absent from the source.

OTHER MAIN-MENU OPTIONS
-----------------------

2  Clean a table already copied to the clipboard.
3  Clean a saved Markdown, CSV or TSV file; you can drag it into the window.
4  Download offline data, or open an official download/sign-in page.
5  Add an official offline address-data download through a guided menu.
6  Teach the programme corrections that you have already approved.
7  Check Python, SQLite, clipboard support and local databases.
8  Explain the free/open address-data sources and their limitations.
9  Turn internet lookup on or off. It is on by default.

The advanced command-line interface remains available, but it is not required
for everyday use.
