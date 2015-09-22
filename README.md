SBBIC Khmer Spelling Checker for Microsoft Word 2010-2013
=========================================================

A Khmer spelling checker Addin for Microsoft Word 2010-2013 using NHunspell (https://sourceforge.net/projects/nhunspell/).

Download
========
To download the completed Addin, visit: http://www.sbbic.org/2014/08/08/sbbic-khmer-spelling-checker-microsoft-word/

How to Contribute and Build
==============
With the Open Source Visual Studio Community Edition 2015, you can now contribute and add to the existing features of Khmer Spell Checker for MS Word.

To start contributing, please install the following tools:

1. Visual Studio Community 2015: Download and Install Visual Studio Community 2015 from the below link. https://www.visualstudio.com/products/visual-studio-community-vs

2. Visual Studio 2015 Installer Project Extension: Download and Install the Visual Studio 2015 Installer Project Extension from the below link. https://visualstudiogallery.msdn.microsoft.com/f1cc3f3e-c300-40a7-8797-c509fb8933b9

3. Office Developer Tools: From the below link, download and install the Office Developer Tools by clicking on “2 Get Office Developer Tools” button. https://www.visualstudio.com/en-us/features/office-tools-vs.aspx

Then once these are installed, download the source by cloning in Desktop or downloading as a ZIP.

4. Once you have the source go to “KhmerSpellCheck\bin\Debug” directory and replace the “km_KH.aff” and “km_KH.dic” files. Please make sure you do not change the file names

5. Double-Click “KhmerSpellCheck.sln” file to open the project in Visual Studio 2010.

6. The solution explorer looks similar to what is shown below. We now would need to rebuild each project in the solution explorer.

7. To rebuild a project, right-click on “KhmerSpellCheck” and click Rebuild. Repeat the steps for “KhmerSpellCheckSetup” and “KhmerSpellCheckSetup64”.

8. Finally, for 32-bit office, go to “KhmerSpellCheckSetup\Debug” folder to get the setup files. And for 64-bit office, go to “KhmerSpellCheckSetup64\Debug” folder to get new setup files.

License
=======
GNU GENERAL PUBLIC LICENSE v2

Special Thanks
==============
Thanks to those who donated to make this a reality (we still need more funds to cover the cost! Donate here: http://goo.gl/6keUJW), as well as to firmusacode for his work on this project: https://www.freelancer.com/u/firmuscode.html
