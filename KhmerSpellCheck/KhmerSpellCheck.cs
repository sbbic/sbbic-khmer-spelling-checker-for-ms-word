using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using NHunspell;
using System.Windows.Forms;
using System.IO;

namespace KhmerSpellCheck
{
    public partial class KhmerSpellCheck
    {
        //The word application
        Word.Application wordApp;
        List<IgnoreWord> ignoreWords = new List<IgnoreWord>();
        List<IgnoreWord> ignoreAllWords = new List<IgnoreWord>();
        char[] KhmerCharacters = new char[] { 
            '\u1780', '\u1781', '\u1782', '\u1783', '\u1784', '\u1785', '\u1786', '\u1787', '\u1788', '\u1789', 
            '\u178A', '\u178B', '\u178C', '\u178D', '\u178E', '\u178F',
            '\u1790', '\u1791', '\u1792', '\u1793', '\u1794', '\u1795', '\u1796', '\u1797', '\u1798', '\u1799', 
            '\u179A', '\u179B', '\u179C', '\u179D', '\u179E', '\u179F',
            '\u17A0', '\u17A1', '\u17A2', '\u17A3', '\u17A4', '\u17A5', '\u17A6', '\u17A7', '\u17A8', '\u17A9', 
            '\u17AA', '\u17AB', '\u17AC', '\u17AD', '\u17AE', '\u17AF',
            '\u17B0', '\u17B1', '\u17B2', '\u17B3', '\u17B4', '\u17B5', '\u17B6', '\u17B7', '\u17B8', '\u17B9', 
            '\u17BA', '\u17BB', '\u17BC', '\u17BD', '\u17BE', '\u17BF',
            '\u17C0', '\u17C1', '\u17C2', '\u17C3', '\u17C4', '\u17C5', '\u17C6', '\u17C7', '\u17C8', '\u17C9', 
            '\u17CA', '\u17CB', '\u17CC', '\u17CD', '\u17CE', '\u17CF',
            '\u17D0', '\u17D1', '\u17D2', '\u17D3',
            '\u17DD'};

        char[] KhmerPunctuations = new char[] {
            '\u17D4', '\u17D5', '\u17D6', '\u17D8', '\u17D9',                                     '\u17D7',
            '\u17DA', '\u17DB', '\u17DC',
            '\u17E0', '\u17E1', '\u17E2', '\u17E3', '\u17E4', '\u17E5', '\u17E6', '\u17E7', '\u17E8', '\u17E9', 
            '\u17F0', '\u17F1', '\u17F2', '\u17F3', '\u17F4', '\u17F5', '\u17F6', '\u17F7', '\u17F8', '\u17F9', 
            '\u19E0', '\u19E1', '\u19E2', '\u19E3', '\u19E4', '\u19E5', '\u19E6', '\u19E7', '\u19E8', '\u19E9', 
            '\u19EA', '\u19EB', '\u19EC', '\u19ED', '\u19EE', '\u19EF',
            '\u19F0', '\u19F1', '\u19F2', '\u19F3', '\u19F4', '\u19F5', '\u19F6', '\u19F7', '\u19F8', '\u19F9', 
            '\u19FA', '\u19FB', '\u19FC', '\u19FD', '\u19FE', '\u19FF',
            '\u2013', '\u2014', '\u2018', '\u2019', '\u201A', '\u201C', '\u201D', '\u201E', '\u2020', '\u2021', 
            '\u2022', '\u2026', '\u2030', '\u2039', '\u203A', '\u2044', '\u20AC', '\u2122', '\u2215', '\u25CC'};

        char[] CommonEnglishPunctuations = new char[] {
            '~', '`', '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '=', '+', '{', '[', '}', ']', 
            '|', '\\', ':', ';', '"', '\'', '<', ',', '>', '.', '?', '/', '1', '2', '3', '4', '5', '6', '7',
            '8', '9', '0', '•', '‣', '‥', '…', '«', '»', '£', '¥', '©', '®', '¶', '·'};

        private void KhmerSpellCheck_Load(object sender, RibbonUIEventArgs e)
        {
            //Get the word application object
            wordApp = Globals.ThisAddIn.Application;
        }

        private void btnSpellCheck_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Exit if there is no active document
                if (wordApp.ActiveDocument == null)
                    return;

                //Get all the words from the active document
                Word.Document doc = wordApp.ActiveDocument;
                //Word.Range entirerange = activeDoc.Range();
                //string[] khmerwords = entirerange.Text.Split(new char[] { '\u200B', ' ', '\r', '\a', '.' }, StringSplitOptions.RemoveEmptyEntries);

                #region Locate Dictionary files
                //Get the deployment directory
                System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

                //Location is where the assembly is run from 
                string assemblyLocation = assemblyInfo.Location;

                //CodeBase is the location of the ClickOnce deployment files
                Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
                string InstallationLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());

                //Khmer dictionaries
                string affFile = Path.Combine(InstallationLocation, "km_KH.aff");
                string dictFile = Path.Combine(InstallationLocation, "km_KH.dic");
                #endregion

                //Loop through all the words to find the first mis-spell
                using (Hunspell hunspell = new Hunspell(affFile, dictFile))
                {
                    //Process all Paragraphs in the documents
                    Object oMissing = System.Reflection.Missing.Value;
                    object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine; // change a line; 
                    object moveExtend = Microsoft.Office.Interop.Word.WdMovementType.wdExtend;

                    //bool isInitial = true; //Is first line

                    doc.ActiveWindow.Selection.HomeKey(Word.WdUnits.wdStory, ref oMissing);

                    ////Read the entire document
                    //while (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc") == false)
                    //{
                    //    if (!isInitial)
                    //    {
                    //        doc.ActiveWindow.Selection.MoveDown(ref WdLine, ref oMissing, ref oMissing);
                    //        doc.ActiveWindow.Selection.HomeKey(ref WdLine, ref oMissing);
                    //    }

                    //    isInitial = false;

                    //    //Skiping table content
                    //    if (doc.ActiveWindow.Selection.get_Information(Word.WdInformation.wdEndOfRangeColumnNumber).ToString() != "-1")
                    //    {
                    //        while (doc.ActiveWindow.Selection.get_Information(Word.WdInformation.wdEndOfRangeColumnNumber).ToString() != "-1")
                    //        {
                    //            if (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                    //                break;

                    //            doc.ActiveWindow.Selection.MoveDown(ref WdLine, ref oMissing, ref oMissing);
                    //            doc.ActiveWindow.Selection.HomeKey(ref WdLine, ref oMissing);
                    //        }
                    //        doc.ActiveWindow.Selection.HomeKey(ref WdLine, ref oMissing);
                    //    }

                    //    //Select the line and get all the khmer words from the line
                    //    doc.ActiveWindow.Selection.EndKey(ref WdLine, ref moveExtend);

                    //    //Set the repeatcheck and stopcheck
                    //    bool repeatcheck, stopcheck;

                    //    //Check Khmer Check
                    //    do
                    //    {
                    //        string selectedText = doc.ActiveWindow.Selection.Text;
                    //        string[] khmerwords = GetKhmerWords(selectedText);

                    //        CheckWords(doc, hunspell, selectedText, khmerwords, out stopcheck, out repeatcheck);

                    //        //Break if user selects to exit
                    //        if (stopcheck) return;
                    //    }
                    //    while (repeatcheck);
                    //}

                    foreach (Word.Paragraph para in doc.Paragraphs)
                    {
                        para.Range.Select();

                        //Set the repeatcheck and stopcheck
                        bool repeatcheck, stopcheck;

                        //Check Khmer Check
                        do
                        {
                            string selectedText = doc.ActiveWindow.Selection.Text;
                            string[] khmerwords = GetKhmerWords(selectedText);

                            CheckWords(doc, hunspell, selectedText, khmerwords, out stopcheck, out repeatcheck);

                            //Break if user selects to exit
                            if (stopcheck) return;
                        }
                        while (repeatcheck);
                    }

                    ////Processing all tables in the documents
                    //for (int iCounter = 1; iCounter <= doc.Tables.Count; iCounter++)
                    //{
                    //    foreach (Word.Row aRow in doc.Tables[iCounter].Rows)
                    //    {
                    //        foreach (Word.Cell aCell in aRow.Cells)
                    //        {
                    //            aCell.Select();

                    //            //Set the repeatcheck and stopcheck
                    //            bool repeatcheck, stopcheck;

                    //            //Check Khmer Check
                    //            do
                    //            {
                    //                //Get the selected text and khmer words
                    //                string selectedText = aCell.Range.Text;
                    //                string[] khmerwords = GetKhmerWords(selectedText);

                    //                CheckWords(doc, hunspell, selectedText, khmerwords, out stopcheck, out repeatcheck, "table", aCell);

                    //                //Break if user selects to exit
                    //                if (stopcheck) return;
                    //            }
                    //            while (repeatcheck);
                    //        }
                    //    }
                    //}

                    var shapes = doc.Shapes;
                    //Finds text within textboxes, then changes them
                    foreach (Microsoft.Office.Interop.Word.Shape shape in shapes)
                    {
                        shape.Select();

                        //Set the repeatcheck and stopcheck
                        bool repeatcheck, stopcheck;

                        //Check Khmer Check
                        do
                        {
                            //Get the selected text and khmer words
                            string selectedText = shape.TextFrame.TextRange.Text;
                            string[] khmerwords = GetKhmerWords(selectedText);

                            CheckWords(doc, hunspell, selectedText, khmerwords, out stopcheck, out repeatcheck, "shape", shape);

                            //Break if user selects to exit
                            if (stopcheck) return;
                        }
                        while (repeatcheck);
                    }


                    MessageBox.Show("Khmer Spelling Check is complete", "Khmer Spell Check", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CheckWords(Word.Document doc, Hunspell hunspell, string selectedText, string[] khmerwords, out bool stopcheck, out bool repeatcheck, string objecttype = "", object wordobject = null)
        {
            int startposition = 0;
            Object oMissing = System.Reflection.Missing.Value;
            stopcheck = repeatcheck = false;

            //Check all the Khmer words from the selected line
            foreach (string khmerword in khmerwords)
            {
                DialogResult dialogResult = DialogResult.None;
                frmKhmer frmKhmer = null;
                String newKhmerWord = String.Empty;

                if (!hunspell.Spell(khmerword))
                {
                    if (!ignoreAllWords.Any(ignoreAllWord => ignoreAllWord.khmerword == khmerword))
                    {
                        if (!ignoreWords.Contains(new IgnoreWord { document = doc.Name, khmerword = khmerword, selectedText = selectedText, startposition = startposition, ignoreAll = false }))
                        {
                            Word.Range start = null;
                            Word.WdColorIndex highlightcolorindex = Word.WdColorIndex.wdNoHighlight;
                            Word.WdUnderline fontunderline = Word.WdUnderline.wdUnderlineNone;
                            Word.WdColor fontcolor = Word.WdColor.wdColorBlack;
                            Word.Range selectionRange = null;

                            //Select the erroneous word on the main document
                            if (String.IsNullOrWhiteSpace(objecttype))
                            {
                                //Set the initial selection
                                start = doc.ActiveWindow.Selection.Range;

                                //Set the search area
                                doc.ActiveWindow.Selection.Start += startposition;
                                Word.Selection searchArea = doc.ActiveWindow.Selection;

                                //Set the find object
                                Word.Find findObject = searchArea.Find;
                                findObject.ClearFormatting();
                                findObject.Text = khmerword;


                                //Find the mis-spelled word
                                findObject.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                                //Temp store the current formatting
                                highlightcolorindex = doc.ActiveWindow.Selection.Range.HighlightColorIndex;
                                fontunderline = doc.ActiveWindow.Selection.Range.Font.Underline;
                                fontcolor = doc.ActiveWindow.Selection.Range.Font.UnderlineColor;

                                //Highlight the selection
                                doc.ActiveWindow.Selection.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                                doc.ActiveWindow.Selection.Range.Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                                doc.ActiveWindow.Selection.Range.Font.UnderlineColor = Word.WdColor.wdColorRed;
                                selectionRange = doc.ActiveWindow.Selection.Range;
                                doc.ActiveWindow.Selection.Collapse();
                            }
                            else
                            {
                                if (objecttype == "table")
                                {
                                    start = ((Word.Cell)wordobject).Range;
                                }
                                else if (objecttype == "shape")
                                {
                                    start = ((Word.Shape)wordobject).TextFrame.TextRange;
                                    start.Start += startposition;
                                }

                                //Set the find object
                                Word.Find findObject = start.Find;
                                findObject.ClearFormatting();
                                findObject.Text = khmerword;

                                //Temp store the current formatting
                                highlightcolorindex = start.HighlightColorIndex;
                                fontunderline = start.Font.Underline;
                                fontcolor = start.Font.UnderlineColor;

                                //Find the mis-spelled word
                                findObject.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                                //Highlight the selection
                                start.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                                start.Font.Underline = Word.WdUnderline.wdUnderlineWavy;
                                start.Font.UnderlineColor = Word.WdColor.wdColorRed;
                                start.Select();
                            }

                            bool isObject = !String.IsNullOrWhiteSpace(objecttype);
                            frmKhmer = new frmKhmer(selectedText, khmerword, startposition, hunspell.Suggest(khmerword), isObject);
                            dialogResult = frmKhmer.ShowDialog();

                            //Select the line again
                            if (String.IsNullOrWhiteSpace(objecttype))
                            {
                                //Revert the highlights
                                selectionRange.Select();
                                doc.ActiveWindow.Selection.Range.HighlightColorIndex = highlightcolorindex;
                                doc.ActiveWindow.Selection.Range.Font.Underline = fontunderline;
                                doc.ActiveWindow.Selection.Range.Font.UnderlineColor = fontcolor;

                                if (dialogResult != DialogResult.Abort) start.Select();
                            }
                            else
                            {
                                start.HighlightColorIndex = highlightcolorindex;
                                start.Font.Underline = fontunderline;
                                start.Font.UnderlineColor = fontcolor;

                                if (dialogResult != DialogResult.Abort)
                                {
                                    if (objecttype == "table")
                                    {
                                        ((Word.Cell)wordobject).Select();
                                    }
                                    else if (objecttype == "shape")
                                    {
                                        ((Word.Shape)wordobject).Select();
                                    }
                                }
                            }
                        }
                    }
                }

                #region Cancel Button Clicked
                //Return if the user hits Cancel Button
                if (dialogResult == DialogResult.Cancel || dialogResult == DialogResult.Abort)
                {
                    stopcheck = true;
                    repeatcheck = false;
                    return;
                }
                #endregion

                #region Ignore or Ignore All Clicked
                //Ignore the word
                if (dialogResult == DialogResult.Ignore)
                {
                    if (frmKhmer.ignoreAll)
                    {
                        ignoreAllWords.Add(new IgnoreWord { khmerword = khmerword, ignoreAll = frmKhmer.ignoreAll });
                    }
                    else
                    {
                        ignoreWords.Add(new IgnoreWord { document = doc.Name, khmerword = khmerword, selectedText = selectedText, startposition = startposition });
                    }
                }
                #endregion

                #region Change or Change All Clicked
                if (dialogResult == DialogResult.Yes)
                {
                    if (String.IsNullOrWhiteSpace(objecttype))
                    {
                        //Set the initial selection
                        Word.Range start = doc.ActiveWindow.Selection.Range;

                        //Set the searcharea
                        if (frmKhmer.changeAll)
                        {
                            doc.Content.Select();
                        }
                        Word.Selection searchArea = doc.ActiveWindow.Selection;

                        //Set the find object
                        Word.Find findObject = searchArea.Find;
                        findObject.ClearFormatting();
                        findObject.Text = khmerword;
                        findObject.Replacement.ClearFormatting();
                        findObject.Replacement.Text = frmKhmer.selectedSuggestion;

                        object replaceAll = frmKhmer.changeAll ? Word.WdReplace.wdReplaceAll : Word.WdReplace.wdReplaceOne;

                        findObject.Execute(ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref replaceAll, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        newKhmerWord = frmKhmer.selectedSuggestion;

                        //Set back the selection
                        start.Select();

                        //Set repeatcheck to true
                        if (frmKhmer.changeAll)
                        {
                            stopcheck = false;
                            repeatcheck = true;
                            return;
                        }
                    }
                    else
                    {
                        var resultingText = selectedText.Replace(khmerword, frmKhmer.selectedSuggestion);

                        if (objecttype == "table")
                        {
                            Word.Range range = ((Word.Cell)wordobject).Range;
                            range.Text = resultingText;
                        }
                        else if (objecttype == "shape")
                        {
                            Word.Shape shape = (Word.Shape)wordobject;
                            shape.TextFrame.TextRange.Text = resultingText;
                        }

                        stopcheck = false;
                        repeatcheck = true;
                        return;
                    }
                }
                #endregion

                startposition += String.IsNullOrWhiteSpace(newKhmerWord) ? khmerword.Length : newKhmerWord.Length;
            }
        }

        private string[] GetKhmerWords(string selectedText)
        {
            string[] khmerwords =
                selectedText
                    .Split(new char[] { '\u200B', '\u200C', ' ', '\r', '\a', '.', '\t', '\v' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(khmerword => khmerword.Trim(KhmerPunctuations))
                    .Select(khmerword => khmerword.Trim(CommonEnglishPunctuations))
                    .Where(khmerword => String.IsNullOrWhiteSpace(khmerword) == false)
                    .Where(khmerword => KhmerCharacters.Contains(khmerword[0]))
                    .ToArray();
            return khmerwords;
        }
    }
}
