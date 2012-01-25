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
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using EnvDTE;
using VSLangProj;
using Microsoft.XmlEditor;

namespace Spring.VisualStudio.Completion
{
    /// <summary>
    /// Provides the completion sets for the Spring.NET XML configuration file.
    /// </summary>
    /// <author>Bruno Baia</author>
    internal class SpringCompletionSource : ICompletionSource
    {
        private bool isDisposed;
        private ITextBuffer textBuffer;
        private IGlyphService glyphService;
        private IVsEditorAdaptersFactoryService vsEditorAdaptersFactoryService;
        private ITextStructureNavigatorSelectorService textStructureNavigatorSelectorService;
        private System.IServiceProvider serviceProvider;

        internal SpringCompletionSource(ITextBuffer textBuffer, IGlyphService glyphService,
            IVsEditorAdaptersFactoryService vsEditorAdaptersFactoryService, 
            ITextStructureNavigatorSelectorService textStructureNavigatorSelectorService, 
            System.IServiceProvider serviceProvider)
        {
            this.textBuffer = textBuffer;
            this.glyphService = glyphService;
            this.vsEditorAdaptersFactoryService = vsEditorAdaptersFactoryService;
            this.textStructureNavigatorSelectorService = textStructureNavigatorSelectorService;
            this.serviceProvider = serviceProvider;
        }

        #region ICompletionSource Members

        public void AugmentCompletionSession(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            XmlLanguageService xls = (XmlLanguageService)this.serviceProvider.GetService(typeof(XmlLanguageService));

            IVsTextView vsTextView = this.vsEditorAdaptersFactoryService.GetViewAdapter(session.TextView);

            int line;
            int column;
            int tmp = vsTextView.GetCaretPos(out line, out column);

            XmlDocument doc = xls.GetParseTree(xls.GetSource(vsTextView), vsTextView, line, column, Microsoft.VisualStudio.Package.ParseReason.CompleteWord);

            //if (doc.RootNamespaceURI.Equals("http://www.springframework.net", StringComparison.OrdinalIgnoreCase))
            //{
            NodeFinder nf = new NodeFinder(line, column);
            nf.Visit(doc);

            if (nf.Scope is XmlElement)
            {
                if (nf.Scope.NamespaceURI.Equals("http://www.springframework.net", StringComparison.OrdinalIgnoreCase))
                {
                    completionSets.Add(CreateCompletionSet(
                           this.GetSnippetsCompletions(),
                           session, true));
                }
            }
            else if (nf.Scope is XmlAttribute)
            {
                //XmlSchemaAttribute sa = nf.Scope.SchemaType as XmlSchemaAttribute;
                //if (sa != null)
                //{
                //if (nf.Scope.Closed)
                //{
                XmlAttribute attr = (XmlAttribute)nf.Scope;
                XmlElement elt = (XmlElement)attr.Parent;
                if (elt.NamespaceURI.Equals("http://www.springframework.net", StringComparison.OrdinalIgnoreCase))
                {
                    if (attr.LocalName.EndsWith("type", StringComparison.OrdinalIgnoreCase))
                    {
                        completionSets.Add(CreateCompletionSet(
                            GetTypeCompletions(attr.Value.Value, false), session, false));
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
                                completionSets.Add(CreateCompletionSet(
                                    GetTypePropertyCompletions(typeName), session, false));
                            }
                        }
                    }
                    else if (attr.LocalName.Equals("name", StringComparison.OrdinalIgnoreCase) &&
                        elt.LocalName.Equals("constructor-arg", StringComparison.OrdinalIgnoreCase))
                    {
                        XmlElement elt2 = (XmlElement)elt.Parent;
                        if (elt2.LocalName.Equals("object", StringComparison.OrdinalIgnoreCase))
                        {
                            string typeName = elt2.GetAttribute("type");
                            if (!String.IsNullOrWhiteSpace(typeName))
                            {
                                completionSets.Add(CreateCompletionSet(
                                    GetTypeCtorsArgsNamesCompletions(typeName), session, false));
                            }
                        }
                    }
                    else if (attr.LocalName.Equals("value", StringComparison.OrdinalIgnoreCase) &&
                        elt.LocalName.Equals("property", StringComparison.OrdinalIgnoreCase))
                    {
                        XmlElement elt2 = (XmlElement)elt.Parent;
                        if (elt2.LocalName.Equals("object", StringComparison.OrdinalIgnoreCase))
                        {
                            string propertyName = elt.GetAttribute("name");
                            string typeName = elt2.GetAttribute("type");
                            if (!String.IsNullOrWhiteSpace(typeName) && !String.IsNullOrWhiteSpace(propertyName))
                            {
                                completionSets.Add(CreateCompletionSet(
                                    GetTypePropertyValueCompletions(typeName, propertyName, attr.Value.Value), session, false));
                            }
                        }
                    }
                }
            }
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
        
        private IList<SpringCompletion> GetSpringAliasesCompletions()
        {
            string[] springAliases = new string[] { "int", "Integer", "int[]", "Integer()", "decimal", "Decimal", "decimal[]", 
                "Decimal()", "char", "Char", "char[]", "Char()", "long", "Long", "long[]", "Long()", "short", "Short", "short[]", 
                "Short()", "uint", "ulong", "ulong[]", "uint[]", "ushort", "ushort[]", "double", "Double", "double[]", "Double()", 
                "float", "Single", "float[]", "Single()", "DateTime", "date", "Date", "DateTime[]", "date[]", "DateTime()", "bool", 
                "Boolean", "bool[]", "Boolean()", "string", "String", "string[]", "String()", "object", "Object", "object[]", "Object()", 
                "int?", "int?[]", "decimal?", "decimal?[]", "char?", "char?[]", "long?", "long?[]", "short?", "short?[]", "uint?", 
                "ulong?", "ulong?[]", "uint?[]", "ushort?", "ushort?[]", "double?", "double?[]", "float?", "float?[]", "bool?", "bool?[]" };

            List<SpringCompletion> completions = new List<SpringCompletion>();

            completions.AddRange(springAliases.Select(
                alias => new SpringCompletion(this.glyphService, alias, alias, "Spring.NET standard type alias", SpringCompletionType.Alias)));

            completions.Sort();
            return completions;
        }

        private IEnumerable<SpringCompletion> GetTypeCompletions(string text, bool addInterfaces)
        {
            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project project = dte.ActiveDocument.ProjectItem.ContainingProject;

            List<SpringCompletion> completions = new List<SpringCompletion>();

            if (!text.Contains('.'))
            {
                completions.AddRange(GetSpringAliasesCompletions());
            }

            if (project.CodeModel != null)
            {
                CodeElements codeElements = GetCodeElements(text, project.CodeModel.CodeElements);
                if (codeElements != null)
                {
                    completions.AddRange(GetCodeElementsCompletions(project, codeElements, addInterfaces));
                }
            }

            completions.Sort();
            return completions;
                }

        private IEnumerable<SpringCompletion> GetTypeCtorsArgsNamesCompletions(string typeName)
        {
            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project project = dte.ActiveDocument.ProjectItem.ContainingProject;

            List<SpringCompletion> completions = new List<SpringCompletion>();

            if (project.CodeModel != null)
            {
                CodeType codeType = project.CodeModel.CodeTypeFromFullName(typeName.Split(',')[0]);

                if (codeType != null)
                {
                    foreach (CodeParameter cp in GetTypeCtorsArgs(codeType))
                    {
                        string comment = String.Empty;
                        try
                        {
                            comment = cp.DocComment;
                        }
                        catch (COMException) { }
                        completions.Add(new SpringCompletion(
                            this.glyphService,
                            cp.Name,
                            cp.Name,
                            String.Format("{0} {1}{2}{3}", cp.Type.AsString, cp.FullName, Environment.NewLine, GetSummaryFromComment(comment)), 
                            SpringCompletionType.ConstructorArg));
                    }
                }
            }

            completions.Sort();
            return completions;
        }

        private IEnumerable<SpringCompletion> GetTypePropertyCompletions(string typeName)
        {
            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project project = dte.ActiveDocument.ProjectItem.ContainingProject;

            List<SpringCompletion> completions = new List<SpringCompletion>();

            if (project.CodeModel != null)
            {
                CodeType codeType = project.CodeModel.CodeTypeFromFullName(typeName.Split(',')[0]);

                if (codeType != null)
                {
                    foreach (CodeProperty cp in GetTypeProperties(codeType))
                    {
                        string comment = String.Empty;
                        try
                        {
                            comment = cp.DocComment;
                        }
                        catch (COMException) { }
                        completions.Add(new SpringCompletion(
                            this.glyphService,
                            cp.Name,
                            cp.Name,
                            String.Format("{0} {1}.{2}{3}{4}", cp.Type.AsString, cp.Parent.Name, cp.Name, Environment.NewLine, GetSummaryFromComment(comment)), 
                            SpringCompletionType.Property));
                    }
                }
            }

            completions.Sort();
            return completions;
        }

        private IEnumerable<SpringCompletion> GetTypePropertyValueCompletions(string typeName, string propertyName, string text)
        {
            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project project = dte.ActiveDocument.ProjectItem.ContainingProject;

            if (project.CodeModel != null)
            {
                CodeType codeType = project.CodeModel.CodeTypeFromFullName(typeName.Split(',')[0]);

                if (codeType != null)
                {
                    CodeProperty cp = GetTypeProperties(codeType).First<CodeProperty>(codeProp => codeProp.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));
                    if (cp.Type.CodeType is CodeEnum)
                    {
                        List<SpringCompletion> completions = new List<SpringCompletion>();
                        CodeEnum ce = cp.Type.CodeType as CodeEnum;
                        foreach (CodeElement ce2 in ce.Members)
                        {
                            if (ce2 is CodeVariable)
                            {
                                CodeVariable cv = ce2 as CodeVariable;
                                string comment = String.Empty;
                                try
                                {
                                    comment = cv.DocComment;
                                }
                                catch (COMException) { }
                                completions.Add(new SpringCompletion(
                                    this.glyphService, 
                                    cv.Name,
                                    cv.Name,
                                    String.Format("{0}.{1}{2}{3}", ce.Name, cv.Name, Environment.NewLine, GetSummaryFromComment(comment)), 
                                    SpringCompletionType.EnumMember));
                            }
                        }
                        completions.Sort();
                        return completions;
                    }
                    else if (cp.Type.CodeType is CodeClass)
                    {
                        if (cp.Type.AsString.Equals("System.Type", StringComparison.OrdinalIgnoreCase) ||
                            cp.Type.AsString.Equals("System.RuntimeType", StringComparison.OrdinalIgnoreCase))
                        {
                            return GetTypeCompletions(text, true);
                        }
                    }
                    else if (cp.Type.AsString.Equals("bool", StringComparison.OrdinalIgnoreCase))
                    {
                        List<SpringCompletion> completions = new List<SpringCompletion>();
                        completions.Add(new SpringCompletion(this.glyphService, "true", "true", "true", SpringCompletionType.Boolean));
                        completions.Add(new SpringCompletion(this.glyphService, "false", "false", "false", SpringCompletionType.Boolean));
                        return completions;
                    }
                    else if (cp.Type.AsString.Equals("System.Type", StringComparison.OrdinalIgnoreCase) || 
                        cp.Type.AsString.Equals("System.RuntimeType", StringComparison.OrdinalIgnoreCase))
                    {
                        return GetTypeCompletions(text, true);
                    }
                }
            }

            return new List<SpringCompletion>();
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

        private static IList<CodeParameter> GetTypeCtorsArgs(CodeType codeType)
        {
            List<CodeParameter> args = new List<CodeParameter>();

            foreach (CodeElement ce in codeType.Bases)
            {
                if (ce is CodeType)
                {
                    args.AddRange(GetTypeCtorsArgs((CodeType)ce));
                }
            }
            foreach (CodeElement ce in codeType.Members)
            {
                if (ce is CodeFunction)
                {
                    CodeFunction cf = ce as CodeFunction;
                    if (cf.FunctionKind == vsCMFunction.vsCMFunctionConstructor)
                    {
                        foreach (CodeElement cp in cf.Parameters)
                        {
                            if (cp is CodeParameter)
                            {
                                if (!args.Any<CodeParameter>(argName => argName.Name.Equals(((CodeParameter)cp).Name, StringComparison.OrdinalIgnoreCase)))
                                {
                                    args.Add((CodeParameter)cp);
                                }
                            }
                        }
                    }
                }
            }

            return args;
        }

        private static CodeElements GetCodeElements(string text, CodeElements codeElements)
        {
            int index = text.IndexOf('.');
            if (index < 0)
            {
                return codeElements;
            }
            foreach (CodeElement ce in codeElements)
            {
                if (ce is CodeNamespace)
                {
                    if (text.Substring(0, index).Equals(ce.Name, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return GetCodeElements(text.Substring(index + 1), ((CodeNamespace)ce).Members);
                    }
                }
            }
            return null;
        }

        private IList<SpringCompletion> GetCodeElementsCompletions(Project prj, CodeElements codeElements, bool addInterfaces)
        {
            List<SpringCompletion> completions = new List<SpringCompletion>();

            foreach (CodeElement ce in codeElements)
            {
                if (ce is CodeNamespace)
                {
                    completions.Add(new SpringCompletion(
                        this.glyphService, 
                        ce.Name,
                        ce.Name,
                        ce.FullName, 
                        SpringCompletionType.Namespace));
                }
                else if (ce is CodeInterface)
                {
                    if (addInterfaces)
                    {
                        CodeInterface ci = ce as CodeInterface;
                        if (ci.Access == vsCMAccess.vsCMAccessPublic)
                        {
                            string comment = String.Empty;
                            try
                            {
                                comment = ci.DocComment;
                            }
                            catch (COMException) { }
                            string assemblyName = GetAssemblyName(prj, ce);
                            completions.Add(new SpringCompletion(
                                this.glyphService,
                                ci.Name,
                                String.Format("{0}, {1}", ci.Name, assemblyName),
                                String.Format("{0}, {1}{2}{3}", ci.FullName, assemblyName, Environment.NewLine, GetSummaryFromComment(comment)), 
                                SpringCompletionType.Interface));
                        }
                    }
                }
                else if (ce is CodeClass)
                {
                    CodeClass cc = ce as CodeClass;
                    if (cc.Access == vsCMAccess.vsCMAccessPublic)
                    {
                        string comment = String.Empty;
                        try
                        {
                            comment = cc.DocComment;
                        }
                        catch (COMException) { }
                        string assemblyName = GetAssemblyName(prj, ce);
                        completions.Add(new SpringCompletion(
                            this.glyphService, 
                            cc.Name,
                            String.Format("{0}, {1}", cc.Name, assemblyName),
                            String.Format("{0}, {1}{2}{3}", cc.FullName, assemblyName, Environment.NewLine, GetSummaryFromComment(comment)), 
                            SpringCompletionType.Class));

                    }
                }
            }

            completions.Sort();
            return completions;
        }

        private static string GetAssemblyName(Project prj, CodeElement ce)
        {
            if (ce.InfoLocation == vsCMInfoLocation.vsCMInfoLocationProject)
            {
                return ce.ProjectItem.ContainingProject.Properties.Item("AssemblyName").Value.ToString();
            }
            if (ce.InfoLocation == vsCMInfoLocation.vsCMInfoLocationExternal)
            {
                // Does not work
                //dynamic exloc = ce.Extender["ExternalLocation"];
                //return exloc.ExternalLocation;

                if (prj.Object is VSProject)
                {
                    VSProject vsPrj = prj.Object as VSProject;

                    foreach (Reference reference in vsPrj.References)
                    {
                        Assembly assembly = Assembly.UnsafeLoadFrom(reference.Path);
                        Type type = assembly.GetType(ce.FullName, false, true);
                        if (type != null)
                        {
                            return assembly.GetName().Name;
                        }
                    }
                }
            }
            return "?";
        }

        private static string GetSummaryFromComment(string comment)
        {
            if (!String.IsNullOrWhiteSpace(comment))
            {
                try
                {
                    System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                    xmlDoc.LoadXml(comment);
                    return xmlDoc.SelectSingleNode("//summary").InnerXml.Trim();
                }
                catch (Exception)
                {
                    return comment;
                }
            }
            return String.Empty;
        }

        private IEnumerable<SpringCompletion> GetSnippetsCompletions()
        {
            List<SpringCompletion> completions = new List<SpringCompletion>();

            IVsTextManager2 expansionManager = (IVsTextManager2)this.serviceProvider.GetService(typeof(SVsTextManager));
            SnippetsEnumerable snippetsEnumerator = new SnippetsEnumerable(expansionManager, GuidList.guidSpringLanguage);
            completions.AddRange(snippetsEnumerator.Select(vsExpansion => new SpringCompletion(this.glyphService, vsExpansion)));

            completions.Sort();
            return completions;
        }

        private CompletionSet CreateCompletionSet(IEnumerable<SpringCompletion> completions, ICompletionSession session, bool isSnippet)
        {
            // Create Spring.NET completion
            return new SpringCompletionSet("Spring.NET",
                "Spring.NET",
                isSnippet ? CreateTrackingSpanForSnippet(session) : CreateTrackingSpan(session),
                completions,
                null);
        }

        private ITrackingSpan CreateTrackingSpan(ICompletionSession session)
        {
            int position = session.GetTriggerPoint(session.TextView.TextBuffer).GetPosition(textBuffer.CurrentSnapshot);

            char[] separators = new[] { '"', '\'', '.' };

            string text = textBuffer.CurrentSnapshot.GetText();
            int first = text.Substring(0, position).LastIndexOfAny(separators) + 1;

            return textBuffer.CurrentSnapshot.CreateTrackingSpan(new Span(first, position - first), SpanTrackingMode.EdgeInclusive);
        }

        private ITrackingSpan CreateTrackingSpanForSnippet(ICompletionSession session)
        {
            ITrackingPoint point = session.GetTriggerPoint(textBuffer);
            SnapshotPoint currentPoint = (session.TextView.Caret.Position.BufferPosition) - 1;
            ITextStructureNavigator navigator = textStructureNavigatorSelectorService.GetTextStructureNavigator(textBuffer);
            TextExtent extent = navigator.GetExtentOfWord(currentPoint);
            return currentPoint.Snapshot.CreateTrackingSpan(extent.Span, SpanTrackingMode.EdgeInclusive);
        }
    }
}