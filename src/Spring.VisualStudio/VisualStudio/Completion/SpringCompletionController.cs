#region License

/*
 * Copyright 2002-2011 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

#endregion

using System;
using System.Linq;
using System.Runtime.InteropServices;

using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.XmlEditor;

namespace Spring.VisualStudio.Completion
{
    /// <summary>
    /// Triggers the IntelliSense completion for the Spring.NET XML configuration file.
    /// </summary>
    /// <author>Bruno Baia</author>
    internal class SpringCompletionController : IOleCommandTarget , IVsExpansionClient
    {
        private ITextView textView;
        private IVsTextView vsTextView;
        private ICompletionBroker completionBroker;
        private ICompletionSession completionSession;
        private IVsExpansionSession expansionSession;
        private ITextStructureNavigatorSelectorService textStructureNavigatorSelectorService;
        private IOleCommandTarget nextCommandTarget;
        private System.IServiceProvider serviceProvider;

        internal SpringCompletionController(System.IServiceProvider serviceProvider, IVsTextView vsTextView, ITextView textView, ICompletionBroker completionBroker, ITextStructureNavigatorSelectorService textStructureNavigatorSelectorService)
        {
            this.serviceProvider = serviceProvider;
            this.vsTextView = vsTextView;
            this.textView = textView;
            this.completionBroker = completionBroker;
            this.textStructureNavigatorSelectorService = textStructureNavigatorSelectorService;

            //add the command to the command chain
            vsTextView.AddCommandFilter(this, out nextCommandTarget);
        }

        private void InsertCodeExpansion(VsExpansion expansion)
        {
            int startLine, startColumn, endLine, endColumn;
            if (completionSession != null)
            {
                // if there is an active completion session we need to use the trigger point of that session
                int position = completionSession.GetTriggerPoint(completionSession.TextView.TextBuffer).GetPosition(textView.TextBuffer.CurrentSnapshot);
                startLine = textView.TextBuffer.CurrentSnapshot.GetLineNumberFromPosition(position);
                startColumn = position - textView.TextBuffer.CurrentSnapshot.GetLineFromPosition(position).Start.Position;

                this.vsTextView.GetCaretPos(out endLine, out endColumn);
            }
            else
            {
                // there is no active completion session so we would use the caret position of the view instead
                this.vsTextView.GetCaretPos(out startLine, out startColumn);
                endColumn = startColumn;
                endLine = startLine;
            }

            InsertCodeExpansion(expansion, startLine, startColumn, endLine, endColumn);
        }

        private void InsertCodeExpansion(VsExpansion expansion, int startLine, int startColumn, int endLine, int endColumn)
        {
            // Insert the selected code snippet and start an expansion session
            IVsTextLines buffer;
            vsTextView.GetBuffer(out buffer);

            // Get the IVsExpansion from the current IVsTextLines
            IVsExpansion vsExpansion = (IVsExpansion)buffer;

            // Call the actual method that performs the snippet insertion
            vsExpansion.InsertNamedExpansion(
                expansion.title,
                expansion.path,
                new TextSpan { iStartIndex = startColumn, iEndIndex = endColumn, iEndLine = endLine, iStartLine = startLine },
                null,
                GuidList.guidSpringLanguage,
                0,
                out expansionSession);
        }

        private void ShowCompletion()
        {
            // If there is no active session
            if (completionSession == null || completionSession.IsDismissed)
            {
                //the caret must be in a non-projection location 
                SnapshotPoint? caretPoint = textView.Caret.Position.Point.GetPoint(
                        textBuffer => (!textBuffer.ContentType.IsOfType("projection")), PositionAffinity.Predecessor);
                if (caretPoint.HasValue)
                {
                    // Trigger the completion session
                    completionSession = completionBroker.TriggerCompletion(textView);

                    // Attach to the active session events
                    if (completionSession != null)
                    {
                        //completionSession.Start();
                        completionSession.Dismissed += new System.EventHandler(OnActiveSessionDismissed);
                        completionSession.Committed += new System.EventHandler(OnActiveSessionCommited);
                        completionSession.Filter();
                    }
                }
            }
        }

        private bool IsCommitKey(char key)
        {
            return char.IsPunctuation(key);
        }

        private bool IsTriggerKey(char key)
        {
            XmlLanguageService xls = (XmlLanguageService)this.serviceProvider.GetService(typeof(XmlLanguageService));
            char quoteChar = xls.XmlPrefs.AutoInsertAttributeQuotes ? '=' : '"';

            if ((!key.Equals(char.MinValue) && char.IsLetterOrDigit(key)) ||
                char.IsPunctuation(key) ||
                key.Equals(quoteChar))
            {
                return true;
            }
            return false;
        }

        private void OnActiveSessionDismissed(object sender, System.EventArgs e)
        {
            completionSession = null;
        }

        private void OnActiveSessionCommited(object sender, System.EventArgs e)
        {
            var selectedCompletion = this.completionSession.SelectedCompletionSet.SelectionStatus.Completion as SpringCompletion;

            // Check if the selected completion is a code snippet
            if (selectedCompletion != null && selectedCompletion.VsExpansion.HasValue)
            {
                InsertCodeExpansion(selectedCompletion.VsExpansion.Value);
            }

            completionSession = null;
        }

        #region IOleCommandTarget

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (pguidCmdGroup == VSConstants.VSStd2K)
            {
                if (nCmdID == (uint)VSConstants.VSStd2KCmdID.INSERTSNIPPET || nCmdID == (uint)VSConstants.VSStd2KCmdID.SURROUNDWITH)
                {
                    IVsTextManager2 textManager = (IVsTextManager2)this.serviceProvider.GetService(typeof(SVsTextManager));
                    IVsExpansionManager expansionManager;
                    if (VSConstants.S_OK == textManager.GetExpansionManager(out expansionManager))
                    {
                        expansionManager.InvokeInsertionUI(
                            vsTextView,
                            this,
                            GuidList.guidSpringLanguage,
                            new string[] { "Expansion" },
                            1,
                            0,
                            null,
                            0,
                            1,
                            "Insert Snippet",
                            string.Empty);
                    }

                    return VSConstants.S_OK;
                }

                if (this.expansionSession != null)
                {
                    // Handle VS Expansion (Code Snippets) keys
                    if ((nCmdID == (uint)VSConstants.VSStd2KCmdID.TAB))
                    {
                        if (expansionSession.GoToNextExpansionField(0) == VSConstants.S_OK)
                            return VSConstants.S_OK;
                    }
                    else if ((nCmdID == (uint)VSConstants.VSStd2KCmdID.BACKTAB))
                    {
                        if (expansionSession.GoToPreviousExpansionField() == VSConstants.S_OK)
                            return VSConstants.S_OK;
                    }
                    else if ((nCmdID == (uint)VSConstants.VSStd2KCmdID.RETURN || nCmdID == (uint)VSConstants.VSStd2KCmdID.CANCEL))
                    {
                        if (expansionSession.EndCurrentExpansion(0) == VSConstants.S_OK)
                        {
                            expansionSession = null;

                            return VSConstants.S_OK;
                        }
                    }
                }

                // Handle Edit.ListMembers or Edit.CompleteWord commands
                if ((nCmdID == (uint)VSConstants.VSStd2KCmdID.SHOWMEMBERLIST || 
                    nCmdID == (uint)VSConstants.VSStd2KCmdID.COMPLETEWORD))
                {
                    if (completionSession != null)
                    {
                        completionSession.Dismiss();
                    }

                    ShowCompletion();

                    return VSConstants.S_OK;
                }

                // Handle Enter/Tab commit keys
                if (completionSession != null && (nCmdID == (uint)VSConstants.VSStd2KCmdID.RETURN || nCmdID == (uint)VSConstants.VSStd2KCmdID.TAB))
                {
                    if (completionSession.SelectedCompletionSet.SelectionStatus.IsSelected)
                    {
                        completionSession.Commit();
                    }
                    else
                    {
                        completionSession.Dismiss();
                    }

                    return VSConstants.S_OK;
                }

                // Handle Code Snippets after pressing the Tab key without completion
                if (completionSession == null && (nCmdID == (uint)VSConstants.VSStd2KCmdID.TAB))
                {
                    IVsTextManager2 expansionManager = (IVsTextManager2)this.serviceProvider.GetService(typeof(SVsTextManager));
                    SnippetsEnumerator snippetsEnumerator = new SnippetsEnumerator(expansionManager, GuidList.guidSpringLanguage);

                    SnapshotPoint currentPoint = (this.textView.Caret.Position.BufferPosition) - 1;
                    ITextStructureNavigator navigator = this.textStructureNavigatorSelectorService.GetTextStructureNavigator(this.textView.TextBuffer);
                    TextExtent extent = navigator.GetExtentOfWord(currentPoint);
                    string shortcut = this.textView.TextSnapshot.GetText(extent.Span);

                    // Search a snippet that matched the token text
                    VsExpansion expansion = snippetsEnumerator.FirstOrDefault(e => e.title == shortcut);
                    if (expansion.title != null)
                    {
                        // Set the location where the snippet will be inserted
                        int startLine, startColumn, endLine, endColumn;

                        this.vsTextView.GetCaretPos(out startLine, out endColumn);
                        startColumn = endColumn - expansion.title.Length;
                        endLine = startLine;

                        // Insert the snippet
                        InsertCodeExpansion(expansion, startLine, startColumn, endLine, endColumn);

                        return VSConstants.S_OK;
                    }
                }

                // Hanlde other keys
                if ((nCmdID == (uint)VSConstants.VSStd2KCmdID.TYPECHAR))
                {
                    char typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);

                    if (completionSession == null)
                    {
                        // Handle trigger keys
                        // Check if the typed char is a trigger
                        if (IsTriggerKey(typedChar))
                        {
                            var result = this.nextCommandTarget.Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

                            ShowCompletion();

                            return result;
                        }
                    }
                    else
                    {
                        // Handle commit keys
                        // Check if the typed char is a commit key
                        if (IsCommitKey(typedChar))
                        {
                            SpringCompletion selectedCompletion = completionSession.SelectedCompletionSet.SelectionStatus.Completion as SpringCompletion;
                            if (completionSession.SelectedCompletionSet.SelectionStatus.IsSelected &&
                                selectedCompletion != null && selectedCompletion.DeclarationType != null && 
                                selectedCompletion.DeclarationType.Value == Declaration.DeclarationType.Namespace)
                            {
                                completionSession.Commit();
                            }
                            else
                            {
                                completionSession.Dismiss();
                            }

                            var result = this.nextCommandTarget.Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

                            // Check we should trigger completion after comitting the previous session (for example, after typing dot '.')
                            if (IsTriggerKey(typedChar))
                            {
                                ShowCompletion();
                            }

                            return result;
                        }
                        else
                        {
                            var result = this.nextCommandTarget.Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                            completionSession.Filter();
                            return result;
                        }
                    }
                }

                // redo the filter if there is a deletion
                if (nCmdID == (uint)VSConstants.VSStd2KCmdID.BACKSPACE || 
                    nCmdID == (uint)VSConstants.VSStd2KCmdID.DELETE)
                {
                    var result = this.nextCommandTarget.Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);

                    if (completionSession != null && !completionSession.IsDismissed)
                    {
                        completionSession.Filter();
                    }

                    return result;
                }
            }

            // we haven't handled this command so pass it onto the next target
            return this.nextCommandTarget.Exec(pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            if (pguidCmdGroup == VSConstants.VSStd2K && cCmds > 0)
            {
                // completion commands should be available
                if (((uint)VSConstants.VSStd2KCmdID.SHOWMEMBERLIST == (uint)prgCmds[0].cmdID || (uint)VSConstants.VSStd2KCmdID.COMPLETEWORD == (uint)prgCmds[0].cmdID))
                {
                    prgCmds[0].cmdf = (int)Constants.MSOCMDF_ENABLED | (int)Constants.MSOCMDF_SUPPORTED;
                    return VSConstants.S_OK;
                }

                // snippet commands should be available
                if ((uint)prgCmds[0].cmdID == (uint)VSConstants.VSStd2KCmdID.INSERTSNIPPET || (uint)prgCmds[0].cmdID == (uint)VSConstants.VSStd2KCmdID.SURROUNDWITH)
                {
                    prgCmds[0].cmdf = (int)Constants.MSOCMDF_ENABLED | (int)Constants.MSOCMDF_SUPPORTED;
                    return VSConstants.S_OK;
                }
            }

            return this.nextCommandTarget.QueryStatus(pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        #endregion

        #region IVsExpansionClient Members

        public int EndExpansion()
        {
            return VSConstants.S_OK;
        }

        public int FormatSpan(IVsTextLines pBuffer, TextSpan[] ts)
        {
            return VSConstants.S_OK;
        }

        public int GetExpansionFunction(MSXML.IXMLDOMNode xmlFunctionNode, string bstrFieldName, out IVsExpansionFunction pFunc)
        {
            pFunc = null;

            return VSConstants.S_OK;
        }

        public int IsValidKind(IVsTextLines pBuffer, TextSpan[] ts, string bstrKind, out int pfIsValidKind)
        {
            pfIsValidKind = 1;

            return VSConstants.S_OK;
        }

        public int IsValidType(IVsTextLines pBuffer, TextSpan[] ts, string[] rgTypes, int iCountTypes, out int pfIsValidType)
        {
            pfIsValidType = 1;

            return VSConstants.S_OK;
        }

        public int OnAfterInsertion(IVsExpansionSession pSession)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeInsertion(IVsExpansionSession pSession)
        {
            return VSConstants.S_OK;
        }

        public int OnItemChosen(string pszTitle, string pszPath)
        {
            InsertCodeExpansion(new VsExpansion { path = pszPath, title = pszTitle });

            return VSConstants.S_OK;
        }

        public int PositionCaretForEditing(IVsTextLines pBuffer, TextSpan[] ts)
        {
            return VSConstants.S_OK;
        }

        #endregion
    }
}
