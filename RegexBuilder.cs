using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Resources;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace Regex_Builder
{
    /// <summary>
    ///		RegexBuilder is a tool for testing out Regular Expressions against sample source text.
    ///		It allows you to easily play with the expression options and examine the hierarchy of
    ///		Matches, Groups, and Captures found.
    ///		
    ///		Copyright (c) Microsoft Corporation.  All rights reserved.
    ///		Author: Scott Louvau
    /// </summary>
    public partial class RegexBuilder : Form
    {
        const RegexOptions TolerantOptions = RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Singleline;

        [STAThread]
        static void Main(string[] args)
        {
            RegexBuilder r = new RegexBuilder();
            r.args = args;

            Application.Run(r);
        }

        public RegexBuilder()
        {
            InitializeComponent();
        }

        #region Member Fields
        public string[] args;

        // Settings
        private int matchLimit;
        private Font fixedFont;
        private string pasteActionString;
        private string baseTitle;

        private string lastStatusWritten;

        // Cross-Thread communication attributes
        private Thread executeThread;
        private delegate void SetExpressionResultDelegate();

        private string expressionForThread;
        private string textForThread;
        private RegexOptions optionsForThread;

        private ArrayList matchesFromThread;
        private string statusFromThread;
        #endregion

        #region RegexOptions Get/Set
        private RegexOptions GetOptions()
        {
            RegexOptions options = RegexOptions.None;

            if (OptionCultureInvariant.Checked) options = options | RegexOptions.CultureInvariant;
            if (OptionEcmaScript.Checked) options = options | RegexOptions.ECMAScript;
            if (OptionExplicitCapture.Checked) options = options | RegexOptions.ExplicitCapture;
            if (OptionIgnoreCase.Checked) options = options | RegexOptions.IgnoreCase;
            if (OptionIgnorePatternWhitespace.Checked) options = options | RegexOptions.IgnorePatternWhitespace;
            if (OptionMultiline.Checked) options = options | RegexOptions.Multiline;
            if (OptionRightToLeft.Checked) options = options | RegexOptions.RightToLeft;
            if (OptionSingleline.Checked) options = options | RegexOptions.Singleline;

            return options;
        }

        private void SetOptions(RegexOptions options)
        {
            OptionCultureInvariant.Checked = (options & RegexOptions.CultureInvariant) != RegexOptions.None;
            OptionEcmaScript.Checked = (options & RegexOptions.ECMAScript) != RegexOptions.None;
            OptionExplicitCapture.Checked = (options & RegexOptions.ExplicitCapture) != RegexOptions.None;
            OptionIgnoreCase.Checked = (options & RegexOptions.IgnoreCase) != RegexOptions.None;
            OptionIgnorePatternWhitespace.Checked = (options & RegexOptions.IgnorePatternWhitespace) != RegexOptions.None;
            OptionMultiline.Checked = (options & RegexOptions.Multiline) != RegexOptions.None;
            OptionRightToLeft.Checked = (options & RegexOptions.RightToLeft) != RegexOptions.None;
            OptionSingleline.Checked = (options & RegexOptions.Singleline) != RegexOptions.None;
        }
        #endregion

        #region Highlight and Scroll Methods
        private void ClearTextHighlight()
        {
            //...clear any previous font and color
            SourceText.SelectAll();
            SourceText.SelectionColor = SourceText.ForeColor;
            SourceText.SelectionBackColor = SourceText.BackColor;
            SourceText.SelectionFont = fixedFont;
        }

        private void HighlightTextRegion(int index, int length)
        {
            if (length == 0)
            {
                // For empty groups, we want to highlight the character before and after.

                int highlightIndex = index;
                int highlightLength = length; // = 0

                //...if the group isn't at the very beginning of the content, we can highlight the previous character
                if (index > 0)
                {
                    highlightIndex--;
                    highlightLength++;
                }

                //...if the group isn't at the very beginning of the content, we can highlight the following character
                if (index + 1 < SourceText.Text.Length)
                {
                    highlightLength++;
                }

                //...make the character before and after (as possible) underlined
                SourceText.Select(highlightIndex, highlightLength);
                SourceText.SelectionBackColor = Color.Red;
            }
            else
            {
                //...make region text bold and underlined
                SourceText.Select(index, length);
                SourceText.SelectionBackColor = Color.Yellow;
            }
        }

        private void ScrollToRegion(int index, int length)
        {
            //..select the right region
            SourceText.Select(index, length);

            //...scroll to the selected region
            SourceText.Select();
            SourceText.ScrollToCaret();
        }
        #endregion

        #region Regex Code and RunExpression method

        private void EscapeExpression()
        {

            Region selection = new Region(RegularExpression.SelectionStart, RegularExpression.SelectionLength);
            string selectedText = RegularExpression.SelectedText;

            string newText = Regex.Escape(selectedText);

            RegularExpression.Text = RegularExpression.Text.Substring(0, selection.Index) + newText + RegularExpression.Text.Substring(selection.Index + selection.Length);
            RegularExpression.Select(selection.Index, newText.Length);
            RegularExpression.Select();
        }

        private void UnescapeExpression()
        {

            Region selection = new Region(RegularExpression.SelectionStart, RegularExpression.SelectionLength);
            string selectedText = RegularExpression.SelectedText;

            string newText = UnescapeString(selectedText);

            RegularExpression.Text = RegularExpression.Text.Substring(0, selection.Index) + newText + RegularExpression.Text.Substring(selection.Index + selection.Length);
            RegularExpression.Select(selection.Index, newText.Length);
            RegularExpression.Select();
        }

        private void RunExpression()
        {
            if (RegularExpression.SelectionLength == 0)
                RunExpression(RegularExpression.Text, SourceText.Text, GetOptions());
            else
                RunExpression(RegularExpression.SelectedText, SourceText.Text, GetOptions());
        }

        private void RunExpression(string expressionText, string sourceText, RegexOptions options)
        {
            //...stop the Thread, if it is running
            if (executeThread != null)
                if (executeThread.IsAlive)
                    executeThread.Abort();

            //...clear the UI and tell the user a new expression is running
            StatusTextBox.Text = "Executing Expression...";

            //...set attributes for the thread to use
            expressionForThread = expressionText;
            textForThread = sourceText;
            optionsForThread = options;

            //...create the Thread object for the thread
            executeThread = new Thread(new ThreadStart(RunExpressionInThread));
            executeThread.Name = "Execute Thread";

            //...start the Thread running the expression (before it stops, it will call back)
            executeThread.Start();
        }

        private void SetExpressionResult()
        {
            // Clear any source text highlight
            ClearTextHighlight();

            // Show the Status message
            StatusTextBox.Text = statusFromThread;

            // Hide the tree (reduce drawing)
            MatchInformation.Hide();

            // Clear the previous Matches
            MatchInformation.SelectedNode = null; // problem - colorization kept after match
            MatchInformation.Nodes.Clear();

            // Build the Tree and highlight the matches
            foreach (TreeNode matchNode in matchesFromThread)
            {
                MatchInformation.Nodes.Add(matchNode);

                Region matchRegion = (Region)matchNode.Tag;
                HighlightTextRegion(matchRegion.Index, matchRegion.Length);
            }

            // Expand the Tree
            MatchInformation.ExpandAll();

            // Show the tree again
            MatchInformation.Show();

            // Write the code which produces the expression
            if (statusFromThread.Length > 0)
            {
                string codeExpression = ExpressionString();

                if (codeExpression.Length > 0)
                    StatusTextBox.Text += "\r\n" + codeExpression;
            }

            lastStatusWritten = StatusTextBox.Text;
        }

        private void RunExpressionInThread()
        {
            RunExpressionInThread(expressionForThread, textForThread, optionsForThread, out matchesFromThread, out statusFromThread);
            this.Invoke(new SetExpressionResultDelegate(SetExpressionResult));
        }

        private void RunExpressionInThread(string expressionText, string sourceText, RegexOptions options, out ArrayList matchNodes, out string status)
        {
            // Set default return values - nothing
            matchNodes = new ArrayList();
            status = "";

            // If there is no expression, just return.
            if (expressionText == "") return;

            // Execute the Expression
            try
            {
                // Get the matches for the expression (up to limit)
                Regex expression = new Regex(expressionText, options);

                // Find the first n Matches
                ArrayList matchList = GetMatches(expression, sourceText);

                if (matchList.Count == 0)
                {
                    // Try with 'tolerant' options to see if there are matches.
                    Regex tExpression = new Regex(expressionText, TolerantOptions);
                    ArrayList tolerantList = GetMatches(tExpression, sourceText);

                    if (tolerantList.Count == 0)
                    {
                        status = "There were no matches.";
                    }
                    else
                    {
                        status = String.Format("There were no matches.\r\nWARNING: {0} matches were found with tolerant options. (See options menu (F6) to see tolerant options)", tolerantList.Count);
                    }

                    return;
                }

                // Determine the number of found matches (to display in status)
                if (matchList.Count == 1)
                    status = "One match found.";
                else if (matchList.Count == matchLimit)
                    status = matchLimit.ToString() + " or more matches found, first " + matchLimit.ToString() + " shown.";
                else
                    status = matchList.Count.ToString() + " matches found.";


                // For every match...
                for (int matchIndex = 0; matchIndex < matchList.Count; matchIndex++)
                {
                    Match matchObject = (Match)matchList[matchIndex];

                    //...add a TreeView entry in Match Information
                    TreeNode matchNode = new TreeNode();
                    matchNode.Text = matchIndex.ToString() + ": [" + matchObject.Value + "]";
                    matchNode.Tag = new Region(matchObject.Index, matchObject.Length, RegionType.Match, matchIndex.ToString(), null, null);

                    //...add a sub-element for each Group (except the first, which is just the full match)
                    for (int groupIndex = 1; groupIndex < matchObject.Groups.Count; groupIndex++)
                    {
                        TreeNode groupNode = new TreeNode();
                        string groupID = expression.GroupNameFromNumber(groupIndex);
                        if (groupID != groupIndex.ToString()) groupID = "\"" + groupID + "\"";
                        groupNode.Text =  groupID + ": [" + matchObject.Groups[groupIndex].Value + "]";
                        groupNode.Tag = new Region(matchObject.Groups[groupIndex].Index, matchObject.Groups[groupIndex].Length, RegionType.Group, matchIndex.ToString(), groupID, null);

                        //...add sub-elements for each Group Capture, if there is more than one capture
                        if (matchObject.Groups[groupIndex].Captures.Count > 1)
                        {
                            for (int captureIndex = 0; captureIndex < matchObject.Groups[groupIndex].Captures.Count; captureIndex++)
                            {
                                TreeNode captureNode = new TreeNode();
                                captureNode.Text = captureIndex.ToString() + ": [" + matchObject.Groups[groupIndex].Captures[captureIndex].Value + "]";
                                captureNode.Tag = new Region(matchObject.Groups[groupIndex].Captures[captureIndex].Index, matchObject.Groups[groupIndex].Captures[captureIndex].Length, RegionType.Capture, matchIndex.ToString(), groupID, captureIndex.ToString());

                                groupNode.Nodes.Add(captureNode);
                            }
                        }

                        matchNode.Nodes.Add(groupNode);
                    }

                    //...matchInformation.Nodes.Add(matchNode) must execute in the main thread
                    matchNodes.Add(matchNode);
                }
            }
            catch (System.ArgumentException ex)
            {
                status = "Error in expression: '" + ex.Message + "'";
            }
        }

        private ArrayList GetMatches(Regex expression, string text)
        {
            ArrayList matchList = new ArrayList();
            int matchCount = 0;

            Match matchObject = expression.Match(text);

            while (matchObject.Success && matchCount < matchLimit)
            {
                matchCount++;
                matchList.Add(matchObject);
                matchObject = matchObject.NextMatch();
            }

            return matchList;
        }
        #endregion

        #region Show Code support
        private string ExpressionString()
        {
            if (mCodeNone.Checked)
                return String.Empty;
            else if (mCodeCS.Checked)
                return String.Format("Regex expression = new Regex(\"{0}\", {1});", CSEscapeString(RegularExpression.Text), OptionsToString("|"));
            else if (mCodeVB.Checked)
                return String.Format("Dim expression As New Regex(\"{0}\", {1})", VBEscapeString(RegularExpression.Text), OptionsToString("Or"));
            else
                return String.Empty;
        }

        private string MatchString(Region region)
        {
            if (mCodeNone.Checked)
                return String.Empty;
            else if (mCodeCS.Checked)
                return CSMatchString(region);
            else if (mCodeVB.Checked)
                return VBMatchString(region);
            else
                return String.Empty;
        }

        private string OptionsToString(string join)
        {
            RegexOptions currentOptions = GetOptions();
            StringBuilder result = new StringBuilder();

            if (currentOptions == RegexOptions.None) return "RegexOptions.None";

            foreach (RegexOptions value in Enum.GetValues(typeof(RegexOptions)))
            {
                if( ((currentOptions & value) != RegexOptions.None))
                {
                    if (result.Length != 0) result.Append(" " + join + " ");
                    result.Append("RegexOptions." + value.ToString());
                }
            }

            return result.ToString();
        }

        private string UnescapeString(string value)
        {
            if (mCodeNone.Checked)
                return Regex.Unescape(value);
            else if (mCodeCS.Checked)
                return CSUnescapeString(value);
            else if (mCodeVB.Checked)
                return VBUnescapeString(value);
            else
                return String.Empty;
        }

        private string CSEscapeString(string value)
        {
            value = value.Replace("\\", "\\\\");
            value = value.Replace("\"", "\\\"");
            value = value.Replace("\r", "\\\r");
            value = value.Replace("\n", "\\\n");
            value = value.Replace("\t", "\\\t");
            return value;
        }

        private string CSUnescapeString(string value)
        {
            value = value.Replace("\\\\", "\\");
            value = value.Replace("\\\"", "\"");
            value = value.Replace("\\\r", "\r");
            value = value.Replace("\\\n", "\n");
            value = value.Replace("\\\t", "\t");
            return value;
        }

        private string VBEscapeString(string value)
        {
            value = value.Replace("\"", "\"\"");
            value = value.Replace("\r\n", "\" + vbCrLf + \"");
            value = value.Replace("\r", "\" + vbCr + \"");
            value = value.Replace("\n", "\" + vbLf + \"");
            value = value.Replace("\t", "\" + vbTab + \"");
            return value;
        }

        private string VBUnescapeString(string value)
        {
            value = value.Replace("\"\"", "\"");
            value = value.Replace("\" + vbCrLf + \"", "\r\n");
            value = value.Replace("\" + vbCr + \"", "\r");
            value = value.Replace("\" + vbLf + \"", "\n");
            value = value.Replace("\" + vbTab + \"", "\t");
            return value;
        }

        private string CSMatchString(Region region)
        {
            if (region.Type == RegionType.Match)
                return String.Format("Match m = expression.Matches(text)[{0}];", region.MatchID);
            else if (region.Type == RegionType.Group)
                return String.Format("Group g = expression.Matches(text)[{0}].Groups[{1}];", region.MatchID, region.GroupID);
            else if (region.Type == RegionType.Capture)
                return String.Format("Capture c = expression.Matches(text)[{0}].Groups[{1}].Captures[{2}];", region.MatchID, region.GroupID, region.CaptureID);
            else
                return String.Empty;
        }

        private string VBMatchString(Region region)
        {
            if (region.Type == RegionType.Match)
                return String.Format("Dim m As Match = expression.Matches(text)({0})", region.MatchID);
            else if (region.Type == RegionType.Group)
                return String.Format("Dim g As Group = expression.Matches(text)({0}).Groups({1})", region.MatchID, region.GroupID);
            else if (region.Type == RegionType.Capture)
                return String.Format("Dim c As Capture = expression.Matches(text)({0}).Groups({1}).Captures({2})", region.MatchID, region.GroupID, region.CaptureID);
            else
                return String.Empty;
        }
        #endregion

        #region RegexFile Load/Save
        private bool LoadRegex(string filePath)
        {
            try
            {
                string text, expression;
                RegexOptions options;

                //...load the file information
                LoadRegexFile(filePath, out text, out expression, out options);

                //...dump the values into the controls
                SourceText.Text = text;
                RegularExpression.Text = expression;
                SetOptions(options);

                //...run the expression
                RunExpression();

                this.Text = baseTitle + " - Loaded '" + args[0] + "'";
                return true;
            }
            catch (Exception ex)
            {
                StatusTextBox.Text = "Error loading RegexFile '" + filePath + "'. Error: " + ex.ToString();
                return false;
            }
        }

        private void LoadRegexFile(string filePath, out string sourceText, out string expression, out RegexOptions options)
        {
            XmlNode e;

            // Read Xml file
            XmlDocument x = new XmlDocument();
            x.Load(filePath);

            // Read the Root element (nothing to do)
            //e = x.DocumentElement;
            //e.GetAttribute("version");

            // Read the Regex Element.
            e = x.SelectSingleNode("//Regex");
            expression = e.InnerText;

            // Read the Options Element
            e = x.SelectSingleNode("//Options");
            options = (RegexOptions)System.Enum.Parse(typeof(RegexOptions), e.Attributes["value"].Value, true);

            // Read the Text Element.
            e = x.SelectSingleNode("//Text");
            sourceText = e.InnerText;
        }

        private bool SaveRegex(string filePath)
        {
            try
            {
                //...save the Regex information to a file
                SaveRegexFile(filePath, SourceText.Text, RegularExpression.Text, GetOptions());

                StatusTextBox.Text = "RegexFile '" + filePath + "' saved.";
                return true;
            }
            catch (Exception ex)
            {
                StatusTextBox.Text = "Error saving RegexFile '" + filePath + "'. Error: " + ex.ToString();
                return false;
            }
        }

        private void SaveRegexFile(string filePath, string sourceText, string expression, RegexOptions options)
        {
            // Create the XmlDocument
            XmlDocument x = new XmlDocument();
            XmlElement e;

            // Create the root element.
            e = x.CreateElement("RegexBuilder");
            e.SetAttribute("version", "1.000");
            x.AppendChild(e);

            // Add the Regex Element.
            e = x.CreateElement("Regex");
            e.InnerText = expression;
            x.DocumentElement.AppendChild(e);

            // Add the Options Element
            e = x.CreateElement("Options");
            e.SetAttribute("value", options.ToString());
            x.DocumentElement.AppendChild(e);

            // Add the Text Element.
            e = x.CreateElement("Text");
            e.InnerText = sourceText;
            x.DocumentElement.AppendChild(e);

            // Write the Regex information file
            x.Save(filePath);
        }

        private void SaveRegexFile(string filePath)
        {
            XmlNode e;

            // Read Xml file
            XmlDocument x = new XmlDocument();
            x.Load(filePath);

            // Read the Root element (nothing to do)
            //e = x.DocumentElement;
            //e.GetAttribute("version");
            //e.GetAttribute("scenario");
            //e.GetAttribute("expectedMatch");

            // Read the Regex Element.
            e = x.SelectSingleNode("//Regex");
            RegularExpression.Text = e.InnerText;

            // Read the Options Element
            e = x.SelectSingleNode("//Options");
            RegexOptions o = (RegexOptions)System.Enum.Parse(typeof(RegexOptions), e.Attributes["value"].Value, true);
            SetOptions(o);

            // Read the Text Element.
            e = x.SelectSingleNode("//Text");
            SourceText.Text = e.InnerText;
        }
        #endregion

        #region Menu Item Helpers (replace selection, wrap selection, ...)
        private void ReplaceSelectionWith(string replaceString)
        {
            //...call the other overload, requesting a 0 char selection after the whole result
            ReplaceSelectionWith(replaceString, new Region(replaceString.Length, 0));
        }

        private void ReplaceSelectionWith(string replaceString, Region finalSelection)
        {
            Region selection = new Region(RegularExpression.SelectionStart, RegularExpression.SelectionLength);
            string expressionText = "";

            //...replace the selection with the replacement text
            expressionText = RegularExpression.Text.Substring(0, selection.Index);
            expressionText += replaceString;
            expressionText += RegularExpression.Text.Substring(selection.Index + selection.Length);

            //...make the requested selection (relative to the start of the replaceString)
            RegularExpression.Text = expressionText;
            RegularExpression.Select();
            RegularExpression.Select(selection.Index + finalSelection.Index, finalSelection.Length);
        }

        private void GroupSelectionAndAdd(string suffixString)
        {
            //...call the other overload, requesting a 0 char selection after the whole result
            GroupSelectionAndAdd(suffixString, new Region(suffixString.Length, 0));
        }

        private void GroupSelectionAndAdd(string suffixString, Region finalSelection)
        {
            Region selection = new Region(RegularExpression.SelectionStart, RegularExpression.SelectionLength);
            string expressionText = "";
            int suffixStartIndex;

            //...surround the selection with parenthesis (if there is a selection), and add the suffix
            expressionText = RegularExpression.Text.Substring(0, selection.Index);

            if (selection.Length > 0)
            {
                expressionText += "(" + RegularExpression.Text.Substring(selection.Index, selection.Length) + ")" + suffixString;
                suffixStartIndex = selection.Index + selection.Length + 2;
            }
            else
            {
                expressionText += suffixString;
                suffixStartIndex = selection.Index;
            }

            expressionText += RegularExpression.Text.Substring(selection.Index + selection.Length);

            //...make the requested selection (relative to the start of the suffixString))
            RegularExpression.Text = expressionText;
            RegularExpression.Select();
            RegularExpression.Select(suffixStartIndex + finalSelection.Index, finalSelection.Length);
        }

        private void InsertSelectionOrPlaceholder(string prefix, string placeholder, string suffix)
        {
            Region selection = new Region(RegularExpression.SelectionStart, RegularExpression.SelectionLength);
            string expressionText = "";
            int selectionLength;

            //...surround the selection (or placeholder, if no selection) with prefix/suffix
            expressionText = RegularExpression.Text.Substring(0, selection.Index);
            expressionText += prefix;

            if (selection.Length > 0)
            {
                expressionText += RegularExpression.Text.Substring(selection.Index, selection.Length);
                selectionLength = selection.Length;
            }
            else
            {
                expressionText += placeholder;
                selectionLength = placeholder.Length;
            }

            expressionText += suffix;
            expressionText += RegularExpression.Text.Substring(selection.Index + selection.Length);

            //...select the placeholder or the old selection
            RegularExpression.Text = expressionText;
            RegularExpression.Select();
            RegularExpression.Select(selection.Index + prefix.Length, selectionLength);
        }

        private void SurroundSelectionWith(string prefix, string suffix)
        {
            //...call the other overload, requesting a 0 char selection after the whole result
            SurroundSelectionWith(prefix, suffix, new Region(prefix.Length + RegularExpression.SelectionLength + suffix.Length, 0));
        }

        private void SurroundSelectionWith(string prefix, string suffix, Region finalSelection)
        {
            Region selection = new Region(RegularExpression.SelectionStart, RegularExpression.SelectionLength);
            string expressionText = "";

            //...surround the selection with prefix/suffix
            expressionText = RegularExpression.Text.Substring(0, selection.Index);
            expressionText += prefix;
            expressionText += RegularExpression.Text.Substring(selection.Index, selection.Length);
            expressionText += suffix;
            expressionText += RegularExpression.Text.Substring(selection.Index + selection.Length);

            //...make the requested selection (relative to the start of the prefix)
            RegularExpression.Text = expressionText;
            RegularExpression.Select();
            RegularExpression.Select(selection.Index + finalSelection.Index, finalSelection.Length);
        }

        private void RunURL(string url)
        {
            try
            {
                System.Diagnostics.Process.Start(url);
            }
            catch (Exception)
            {
                MessageBox.Show(this, "Regex Builder could not browse to a URL due to security restrictions. Please copy Regex Builder to your local drive and run it from there.");
            }
        }
        #endregion

        #region Event Handlers
        private void RegexBuilder_Load(object sender, EventArgs e)
        {
            // Read the base title from the dialog; we'll modify it on load/save.
            baseTitle = this.Text;

            // Define the fixed font to use (normal and for selections)
            fixedFont = new Font("Courier New", 9);

            SourceText.Font = fixedFont;
            RegularExpression.Font = fixedFont;

            // Define the match limit - 500 (will be configurable)
            matchLimit = 500;

            // Determine the string used to identify the "Paste" action in a RichTextBox
            try
            {
                ResourceManager Manager = new ResourceManager("System.Windows.Forms", typeof(RichTextBox).Assembly);
                pasteActionString = Manager.GetString("RichTextBox_IDPaste");
            }
            catch (Exception)
            {
                //...fall back on the English string if we can't read the resource.
                pasteActionString = "Paste";
            }

            // Load the RegexFile given as the application argument, if found. Otherwise, run the sample expression.
            if (args.Length == 1)
                LoadRegex(args[0]);
            else
                RunExpression();
        }

        private void OptionsButton_Click(object sender, EventArgs e)
        {
            OptionsMenu.Show(OptionsButton, new Point(OptionsButton.Width, 0));
        }

        private void MenuButton_Click(object sender, EventArgs e)
        {
            ExpressionMenu.Show(MenuButton, new Point(MenuButton.Width, 0));
        }

        private void ExecuteButton_Click(object sender, EventArgs e)
        { RunExpression(); }

        private void RegularExpression_TextChanged(object sender, EventArgs e)
        {
            // On a Paste, de-RTF the text and move the cursor back to the previous position.
            if (RegularExpression.UndoActionName == pasteActionString)
            {
                int cursorPosition = RegularExpression.SelectionStart;

                //...get the text (not the RTF) and then clear and set it, to clear formatting information
                string text = RegularExpression.Text;
                RegularExpression.Clear();
                RegularExpression.Text = text;

                //...place the cursor at the correct position and scroll it into view
                RegularExpression.Select();
                RegularExpression.Select(cursorPosition, 0);
                RegularExpression.ScrollToCaret();
            }

            RunExpression(); 
        }

        private void OptionChanged(object sender, EventArgs e)
        { RunExpression(); }

        private void MatchInformation_AfterSelect(object sender, TreeViewEventArgs e)
        {
            // If a tag value can be found for the selected node...
            if (MatchInformation.SelectedNode.Tag != null)
            {
                //...make the text corresponding to that region red
                try
                {
                    Region selectedRegion = (Region)(MatchInformation.SelectedNode.Tag);

                    //...clear any previous highlight
                    ClearTextHighlight();

                    //...make region text bold and underlined
                    HighlightTextRegion(selectedRegion.Index, selectedRegion.Length);

                    StatusTextBox.Text = lastStatusWritten + "\r\n" + MatchString(selectedRegion);

                    //...scroll to the selected region
                    ScrollToRegion(selectedRegion.Index, selectedRegion.Length);
                }
                catch (InvalidCastException)
                {
                    StatusTextBox.Text = "Unexpected error - an element in the TreeView had no corresponding capture.";
                }
            }
        }

        private void SourceText_TextChanged(object sender, EventArgs e)
        {
            // On a Paste, de-RTF the text and move the cursor back to the previous position.
            if (SourceText.UndoActionName == pasteActionString)
            {
                int cursorPosition = SourceText.SelectionStart;

                //...get the text (not the RTF) and then clear and set it, to clear formatting information
                string text = SourceText.Text;
                SourceText.Clear();
                SourceText.Text = text;

                //...place the cursor at the correct position and scroll it into view
                SourceText.Select();
                SourceText.Select(cursorPosition, 0);
                SourceText.ScrollToCaret();

                return;
            }

            // We can't run the expression here. When the expression runs we immediately
            // highlight all matches and make a selection, which means each character you type
            // will go after the selection we make instead of the cursor position you had in mind.
        }
        #endregion

        #region MenuItem Events
        private void mTab_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\t"); }

        private void mCarriageReturn_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\r"); }

        private void mNewline_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\n"); }

        private void mAlarm_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\a"); }

        private void mBoundary_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\b"); }

        private void mVerticalTab_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\v"); }

        private void mFormFeed_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\f"); }

        private void mEscape_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\e"); }

        private void mOctalValue_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\0nn", new Region(2, 2)); }

        private void mHexValue_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\xnn", new Region(2, 2)); }

        private void mUnicodeValue_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\unnnn", new Region(2, 4)); }

        private void mSingleCharacter_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"."); }

        private void mWordCharacter_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\w"); }

        private void mNonwordCharacter_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\W"); }

        private void mWhitespaceCharacter_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\s"); }

        private void mNonWhitespaceCharacter_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\S"); }

        private void mDecimalDigit_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\d"); }

        private void mNondigit_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\D"); }

        private void mCharacterInSet_Click(object sender, System.EventArgs e)
        { InsertSelectionOrPlaceholder(@"[", @"abc", @"]"); }

        private void mCharacterNotInSet_Click(object sender, System.EventArgs e)
        { InsertSelectionOrPlaceholder(@"[^", @"abc", @"]"); }

        private void mCharacterInRange_Click(object sender, System.EventArgs e)
        { InsertSelectionOrPlaceholder(@"[", @"0-9", @"]"); }

        private void mCharacterInNamedSet_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\p{name}", new Region(3, 4)); }

        private void mCharacterNotInNamedSet_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\P{name}", new Region(3, 4)); }

        private void mBeginningOfLine_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"^"); }

        private void mEndOfLine_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"$"); }

        private void mBeginningOfString_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\A"); }

        private void mEndOfString_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\Z"); }

        private void mEndOfString2_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\z"); }

        private void mAfterLastMatch_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\G"); }

        private void mWordBoundary_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\b"); }

        private void mNotWordBoundary_Click(object sender, System.EventArgs e)
        { ReplaceSelectionWith(@"\B"); }

        private void mZeroOrMore_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"*"); }

        private void mOneOrMore_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"+"); }

        private void mZeroOrOne_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"?"); }

        private void mExactlyN_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"{n}", new Region(1, 1)); }

        private void mAtLeastN_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"{n,}", new Region(1, 1)); }

        private void mBetweenNAndM_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"{n,m}", new Region(1, 3)); }

        private void mZeroOrMoreMinimal_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"*?"); }

        private void mOneOrMoreMinimal_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"+?"); }

        private void mZeroOrOneMinimal_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"??"); }

        private void mAtLeastNMinimal_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"{n,}?", new Region(1, 1)); }

        private void mBetweenNAndMMinimal_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@"{n,m}?", new Region(1, 3)); }

        private void mCapture_Click(object sender, System.EventArgs e)
        { GroupSelectionAndAdd(@""); }

        private void mNamedCapture_Click(object sender, System.EventArgs e)
        { SurroundSelectionWith(@"(?<name>", @")", new Region(3, 4)); }

        private void mNonCapture_Click(object sender, System.EventArgs e)
        { SurroundSelectionWith(@"(?:", @")"); }

        private void mBalancingGroup_Click(object sender, System.EventArgs e)
        { SurroundSelectionWith(@"(?<name1-name2>", @"", new Region(3, 11)); }

        private void mOptionsGroup_Click(object sender, System.EventArgs e)
        { SurroundSelectionWith(@"(?imnsx-imnsx:", @")", new Region(2, 11)); }

        private void mCurrentOptionsGroup_Click(object sender, System.EventArgs e)
        {
            string currentOptions = "";
            RegexOptions options = GetOptions();

            //...determine which options we can set with a group
            if ((options & RegexOptions.IgnoreCase) != RegexOptions.None) currentOptions += "i";
            if ((options & RegexOptions.Multiline) != RegexOptions.None) currentOptions += "m";
            if ((options & RegexOptions.ExplicitCapture) != RegexOptions.None) currentOptions += "n";
            if ((options & RegexOptions.Singleline) != RegexOptions.None) currentOptions += "s";
            if ((options & RegexOptions.IgnorePatternWhitespace) != RegexOptions.None) currentOptions += "x";

            //...if no pushable options were set, don't do anything
            if (currentOptions.Length == 0) return;

            //...calculate the set options we can't represent in a group
            RegexOptions newOptions = RegexOptions.None;
            newOptions = newOptions | (options & RegexOptions.Compiled);
            newOptions = newOptions | (options & RegexOptions.CultureInvariant);
            newOptions = newOptions | (options & RegexOptions.ECMAScript);
            newOptions = newOptions | (options & RegexOptions.RightToLeft);

            //...set checkboxes for only options not represented in the group
            SetOptions(newOptions);

            //...surround the whole expression with the set options
            RegularExpression.SelectAll();
            SurroundSelectionWith("(?" + currentOptions + ":", ")", new Region(2, currentOptions.Length));
        }

        private void mPositiveLookahead_Click(object sender, System.EventArgs e)
        { SurroundSelectionWith(@"(?=", @")"); }

        private void mNegativeLookahead_Click(object sender, System.EventArgs e)
        { SurroundSelectionWith(@"(?!", @")"); }

        private void mPositiveLookbehind_Click(object sender, System.EventArgs e)
        { SurroundSelectionWith(@"(?<=", @")"); }

        private void mNegativeLookbehind_Click(object sender, System.EventArgs e)
        { SurroundSelectionWith(@"(?<!", @")"); }

        private void mNonBacktracking_Click(object sender, System.EventArgs e)
        { SurroundSelectionWith(@"(?>", @")"); }

        private void mMainHelp_Click(object sender, System.EventArgs e)
        { RunURL("http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconRegularExpressionsLanguageElements.asp"); }

        private void mCharEscapesHelp_Click(object sender, System.EventArgs e)
        { RunURL("http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconcharacterescapes.asp"); }

        private void mCharClassesHelp_Click(object sender, System.EventArgs e)
        { RunURL("http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconcharacterclasses.asp"); }

        private void mAssertionHelp_Click(object sender, System.EventArgs e)
        { RunURL("http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconatomiczero-widthassertions.asp"); }

        private void mQuantifierHelp_Click(object sender, System.EventArgs e)
        { RunURL("http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconquantifiers.asp"); }

        private void mGroupingHelp_Click(object sender, System.EventArgs e)
        { RunURL("http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpcongroupingconstructs.asp"); }

        private void mOptionsHelp_Click(object sender, EventArgs e)
        { RunURL("http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconregularexpressionoptions.asp"); }

        private void mLoadSourceText_Click(object sender, System.EventArgs e)
        {
            OpenFileDialog selector = new OpenFileDialog();
            selector.Multiselect = false;

            if (selector.ShowDialog() == DialogResult.OK)
            {
                StreamReader file = new StreamReader(selector.OpenFile());
                SourceText.Text = file.ReadToEnd();
                file.Close();

                RunExpression();
            }
        }

        private void mLoadFile_Click(object sender, System.EventArgs e)
        {
            OpenFileDialog selector = new OpenFileDialog();
            selector.Filter = "Regex Files (*.xml)|*.xml";
            selector.Multiselect = false;

            if (selector.ShowDialog() == DialogResult.OK)
                LoadRegex(selector.FileName);
        }

        private void mSaveFile_Click(object sender, System.EventArgs e)
        {
            SaveFileDialog selector = new SaveFileDialog();
            selector.Filter = "Regex Files (*.xml)|*.xml";

            if (selector.ShowDialog() == DialogResult.OK)
                SaveRegex(selector.FileName);
        }

        private void mEscapeSelectionExpression_Click(object sender, System.EventArgs e)
        { EscapeExpression(); }

        private void mUnescapeSelectionExpression_Click(object sender, EventArgs e)
        { UnescapeExpression(); }

        private void mExecuteSelectionExpression_Click(object sender, System.EventArgs e)
        { RunExpression(); }

        private void mExecute_Click(object sender, System.EventArgs e)
        { RunExpression(); }

        private void mFeedback_Click(object sender, System.EventArgs e)
        { RunURL("mailto:scottlo@microsoft.com"); }


        private void mSetTolerantOptions_Click(object sender, EventArgs e)
        {
            SetOptions(TolerantOptions);
            RunExpression();
        }

        private void mCodeChange(object sender, EventArgs e)
        {
            mCodeCS.Checked = false;
            mCodeVB.Checked = false;
            mCodeNone.Checked = false;
            ((MenuItem)sender).Checked = true;
            RunExpression();
        }

        #endregion
    }

    #region Region Class
    public enum RegionType
    {
        Match,
        Group,
        Capture
    }

    public class Region
    {
        public int Index, Length;
        public RegionType Type;
        public string MatchID;
        public string GroupID;
        public string CaptureID;

        public Region(int idx, int len)
        {
            Index = idx;
            Length = len;
        }

        public Region(int idx, int len, RegionType type, string matchID, string groupID, string captureID )
        {
            Index = idx;
            Length = len;
            Type = type;
            MatchID = matchID;
            GroupID = groupID;
            CaptureID = captureID;
        }
    }
    #endregion
}