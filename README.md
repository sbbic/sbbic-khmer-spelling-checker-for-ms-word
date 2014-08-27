SBBIC Khmer Spelling Checker for Microsoft Word 2010-2013
=========================================================

A Khmer spelling checker Addin for Microsoft Word 2010-2013 using NHunspell (https://sourceforge.net/projects/nhunspell/).

Download
========
To download the completed Addin, visit: http://www.sbbic.org/2014/08/08/sbbic-khmer-spelling-checker-microsoft-word/

How to Rebuild
==============
To rebuild the Addin with a new Hunspell dictionary (you will need Visual Studio 2010):

1 – Goto “KhmerSpellCheck\bin\Debug” directory and replace the “km_KH.aff” and “km_KH.dic” files. Please make sure you do not change the file names

2 – Double-Click “KhmerSpellCheck.sln” file to open the project in Visual Studio 2010.

3 – The solution explorer looks similar to what is shown below. We now would need to rebuild each project in the solution explorer.

4 – To rebuild a project, right-click on “KhmerSpellCheck” and click Rebuild. Repeat the steps for “KhmerSpellCheckSetup” and “KhmerSpellCheckSetup64”.

5 – Finally, for 32-bit office, go to “KhmerSpellCheckSetup\Debug” folder to get the setup files. And for 64-bit office, go to “KhmerSpellCheckSetup64\Debug” folder to get new setup files.

License
=======
GNU GENERAL PUBLIC LICENSE v2

Special Thanks
==============
Thanks to those who donated to make this a reality (we still need more funds to cover the cost! Donate here: http://goo.gl/6keUJW), as well as to firmusacode for his work on this project: https://www.freelancer.com/u/firmuscode.html
