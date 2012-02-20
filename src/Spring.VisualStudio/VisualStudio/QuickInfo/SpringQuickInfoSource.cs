#region License

/*
 * Copyright 2002-2012 the original author or authors.
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
using System.Collections.Generic;

using Microsoft.XmlEditor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Editor;
using EnvDTE;

namespace Spring.VisualStudio.QuickInfo
{
    /// <summary>
    /// Provides the QuickInfo information for the Spring.NET XML configuration file.
    /// </summary>
    /// <author>Bruno Baia</author>
    public class SpringQuickInfoSource : IQuickInfoSource
    {
        private bool isDisposed;
        private ITextStructureNavigatorSelectorService textStructureNavigatorSelectorService;
        private ITextBuffer textBuffer;
        private IVsEditorAdaptersFactoryService vsEditorAdaptersFactoryService;
        private System.IServiceProvider serviceProvider;

        public SpringQuickInfoSource(ITextBuffer textBuffer, ITextStructureNavigatorSelectorService textStructureNavigatorSelectorService,
            IVsEditorAdaptersFactoryService vsEditorAdaptersFactoryService, System.IServiceProvider serviceProvider)
        {
            this.textBuffer = textBuffer;
            this.textStructureNavigatorSelectorService = textStructureNavigatorSelectorService;
            this.vsEditorAdaptersFactoryService = vsEditorAdaptersFactoryService;
            this.serviceProvider = serviceProvider;
        }

        #region IQuickInfoSource Members

        public void AugmentQuickInfoSession(IQuickInfoSession session, IList<object> quickInfoContent, out ITrackingSpan applicableToSpan)
        {
            // Map the trigger point down to our buffer.
            SnapshotPoint? subjectTriggerPoint = session.GetTriggerPoint(textBuffer.CurrentSnapshot);
            if (subjectTriggerPoint.HasValue)
            {
                XmlLanguageService xls = (XmlLanguageService)this.serviceProvider.GetService(typeof(XmlLanguageService));

                IVsTextView vsTextView = this.vsEditorAdaptersFactoryService.GetViewAdapter(session.TextView);

                int line;
                int column;
                int tmp = vsTextView.GetLineAndColumn(subjectTriggerPoint.Value.Position, out line, out column);

                XmlDocument doc = xls.GetParseTree(xls.GetSource(vsTextView), vsTextView, line, column, Microsoft.VisualStudio.Package.ParseReason.CompleteWord);

                //if (doc.RootNamespaceURI.Equals("http://www.springframework.net", StringComparison.OrdinalIgnoreCase))
                //{
                NodeFinder nf = new NodeFinder(line, column);
                nf.Visit(doc);

                if (nf.Scope is XmlAttribute)
                {
                    XmlAttribute attr = (XmlAttribute)nf.Scope;
                    XmlElement elt = (XmlElement)attr.Parent;

                    if (elt.NamespaceURI.Equals("http://www.springframework.net", StringComparison.OrdinalIgnoreCase))
                    {
                        if (attr.LocalName.EndsWith("type", StringComparison.OrdinalIgnoreCase))
                        {
                            string comment = GetTypeComment(attr.LiteralValue);
                            if (comment != null)
                            {
                                quickInfoContent.Add(comment);

                                applicableToSpan = subjectTriggerPoint.Value.Snapshot.CreateTrackingSpan(
                                    subjectTriggerPoint.Value.Position - (column - attr.Value.SourceContext.StartColumn),
                                    attr.Value.SourceContext.EndColumn - attr.Value.SourceContext.StartColumn,
                                    SpanTrackingMode.EdgeInclusive);

                                return;
                            }
                        }
                        else if (attr.LocalName.Equals("name", StringComparison.OrdinalIgnoreCase) &&
                            elt.LocalName.Equals("property", StringComparison.OrdinalIgnoreCase))
                        {
                            XmlElement elt2 = (XmlElement)elt.Parent;
                            if (elt2.LocalName.Equals("object", StringComparison.OrdinalIgnoreCase))
                            {
                                string typeName = elt2.GetAttribute("type");
                                if (!String.IsNullOrWhiteSpace(typeName))
                                {
                                    string comment = GetPropertyComment(typeName, attr.LiteralValue);
                                    if (comment != null)
                                    {
                                        quickInfoContent.Add(comment);

                                        applicableToSpan = subjectTriggerPoint.Value.Snapshot.CreateTrackingSpan(
                                            subjectTriggerPoint.Value.Position - (column - attr.Value.SourceContext.StartColumn),
                                            attr.Value.SourceContext.EndColumn - attr.Value.SourceContext.StartColumn,
                                            SpanTrackingMode.EdgeInclusive);

                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
                //}
            }

            applicableToSpan = null;
        }

        public void Dispose()
        {
            if (!isDisposed)
            {
                GC.SuppressFinalize(this);
                isDisposed = true;
            }
        }

        #endregion

        private string GetTypeComment(string typeName)
        {
            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project prj = dte.ActiveDocument.ProjectItem.ContainingProject;
            if (prj.CodeModel != null)
            {
                CodeType codeType = prj.CodeModel.CodeTypeFromFullName(typeName.Split(',')[0]);
                if (codeType != null)
                {
                    string comment = String.Empty;
                    try
                    {
                        comment = GetSummaryFromComment(codeType.DocComment);
                    }
                    catch (Exception) { }
                    return comment;
                }
            }
            return null;
        }

        private string GetPropertyComment(string typeName, string propertyName)
        {
            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project prj = dte.ActiveDocument.ProjectItem.ContainingProject;
            if (prj.CodeModel != null)
            {
                CodeType codeType = prj.CodeModel.CodeTypeFromFullName(typeName.Split(',')[0]);
                if (codeType != null)
                {
                    foreach (CodeProperty cp in GetTypeProperties(codeType))
                    {
                        if (cp.Name.Equals(propertyName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            string comment = String.Empty;
                            try
                            {
                                comment = GetSummaryFromComment(cp.DocComment);
                            }
                            catch (Exception) { }
                            return comment;
                        }
                    }
                }
            }
            return null;
        }

        private static IList<CodeProperty> GetTypeProperties(CodeType codeType)
        {
            List<CodeProperty> properties = new List<CodeProperty>();

            foreach (CodeElement ce in codeType.Bases)
            {
                if (ce is CodeType)
                {
                    properties.AddRange(GetTypeProperties((CodeType)ce));
                }
            }
            foreach (CodeElement ce in codeType.Members)
            {
                if (ce is CodeProperty)
                {
                    if (!properties.Any<CodeProperty>(propName => propName.Name.Equals(((CodeProperty)ce).Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        properties.Add((CodeProperty)ce);
                    }
                }
            }

            return properties;
        }

        private static string GetSummaryFromComment(string comment)
        {
            if (!String.IsNullOrWhiteSpace(comment))
            {
                System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                xmlDoc.LoadXml(comment);
                return xmlDoc.SelectSingleNode("//summary").InnerXml.Trim();
            }
            return String.Empty;
        }
    }
}
