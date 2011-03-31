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
                    this.AddSnippetsCompletionSets(session, completionSets);
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
                        this.AddTypeCompletionSets(session, completionSets, attr.Value.Value, false);
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
                                this.AddTypePropertiesCompletionSets(session, completionSets, typeName);
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
                                this.AddTypeCtorNamesCompletionSets(session, completionSets, typeName);
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
                                this.AddPropertyValuesCompletionSets(session, completionSets, attr.Value.Value, typeName, propertyName);
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

        private void AddSnippetsCompletionSets(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            List<Declaration> declarations = new List<Declaration>();

            IVsTextManager2 expansionManager = (IVsTextManager2)this.serviceProvider.GetService(typeof(SVsTextManager));
            SnippetsEnumerable snippetsEnumerator = new SnippetsEnumerable(expansionManager, GuidList.guidSpringLanguage);
            declarations.AddRange(snippetsEnumerator.Select(expansion => new Declaration(expansion.shortcut, expansion.title, Declaration.DeclarationType.Snippet, expansion.description)));

            completionSets.Add(GetCompletions(declarations, session, true));
        }

        private void AddTypeCompletionSets(ICompletionSession session, IList<CompletionSet> completionSets, string attrValue, bool addInterfaces)
        {
            List<Declaration> declarations = new List<Declaration>();

            if (!attrValue.Contains('.'))
            {
                declarations.AddRange(GetSpringAliasesDeclarations());
            }
            declarations.AddRange(GetProjectDeclarations(attrValue, addInterfaces));

            completionSets.Add(GetCompletions(declarations, session, false));
        }

        private void AddTypeCtorNamesCompletionSets(ICompletionSession session, IList<CompletionSet> completionSets, string typeName)
        {
            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project prj = dte.ActiveDocument.ProjectItem.ContainingProject;
            if (prj.CodeModel != null)
            {
                CodeType codeType = prj.CodeModel.CodeTypeFromFullName(typeName.Split(',')[0]);

                if (codeType != null)
                {
                    IList<Declaration> declarations = new List<Declaration>();

                    foreach (CodeParameter cp in GetTypeContructorsArgs(codeType))
                    {
                        declarations.Add(new Declaration(
                            cp.Name,
                            cp.Name,
                            Declaration.DeclarationType.ConstructorArg,
                            String.Format("{0} {1}{2}{3}", cp.Type.AsString, cp.FullName, Environment.NewLine, GetSummaryFromComment(cp.DocComment))));
                    }

                    completionSets.Add(GetCompletions(declarations, session, false));
                }
            }
        }

        private void AddTypePropertiesCompletionSets(ICompletionSession session, IList<CompletionSet> completionSets, string typeName)
        {
            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project prj = dte.ActiveDocument.ProjectItem.ContainingProject;
            if (prj.CodeModel != null)
            {
                CodeType codeType = prj.CodeModel.CodeTypeFromFullName(typeName.Split(',')[0]);

                if (codeType != null)
                {
                    IList<Declaration> declarations = new List<Declaration>();

                    foreach (CodeProperty cp in GetTypeProperties(codeType))
                    {
                        declarations.Add(new Declaration(
                            cp.Name,
                            cp.Name,
                            Declaration.DeclarationType.Property,
                            String.Format("{0} {1}{2}{3}", cp.Type.AsString, cp.FullName, Environment.NewLine, GetSummaryFromComment(cp.DocComment))));
                    }

                    completionSets.Add(GetCompletions(declarations, session, false));
                }
            }
        }

        private void AddPropertyValuesCompletionSets(ICompletionSession session, IList<CompletionSet> completionSets, string attrValue, string typeName, string propertyName)
        {
            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project prj = dte.ActiveDocument.ProjectItem.ContainingProject;
            if (prj.CodeModel != null)
            {
                CodeType codeType = prj.CodeModel.CodeTypeFromFullName(typeName.Split(',')[0]);

                if (codeType != null)
                {
                    CodeProperty cp = GetTypeProperties(codeType).First<CodeProperty>(codeProp => codeProp.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));
                    if (cp.Type.CodeType is CodeEnum)
                    {
                        List<Declaration> declarations = new List<Declaration>();
                        CodeEnum ce = cp.Type.CodeType as CodeEnum;
                        foreach (CodeElement ce2 in ce.Members)
                        {
                            if (ce2 is CodeVariable)
                            {
                                CodeVariable cv = ce2 as CodeVariable;
                                declarations.Add(new Declaration(
                                    cv.Name,
                                    cv.Name,
                                    Declaration.DeclarationType.EnumMember,
                                    String.Format("{0}{1}{2}", cv.FullName, Environment.NewLine, GetSummaryFromComment(cv.DocComment))));
                            }
                        }
                        completionSets.Add(GetCompletions(declarations, session, false));
                    }
                    else if (cp.Type.AsString.Equals("bool", StringComparison.OrdinalIgnoreCase))
                    {
                        List<Declaration> declarations = new List<Declaration>();
                        declarations.Add(new Declaration("true", "true", Declaration.DeclarationType.Boolean, "true"));
                        declarations.Add(new Declaration("false", "false", Declaration.DeclarationType.Boolean, "false"));
                        completionSets.Add(GetCompletions(declarations, session, false));
                    }
                    else if (cp.Type.AsString.Equals("System.Type", StringComparison.OrdinalIgnoreCase) ||
                        cp.Type.AsString.Equals("System.RuntimeType", StringComparison.OrdinalIgnoreCase))
                    {
                        AddTypeCompletionSets(session, completionSets, attrValue, true);
                    }
                }
            }
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

        private static IList<CodeParameter> GetTypeContructorsArgs(CodeType codeType)
        {
            List<CodeParameter> args = new List<CodeParameter>();

            foreach (CodeElement ce in codeType.Bases)
            {
                if (ce is CodeType)
                {
                    args.AddRange(GetTypeContructorsArgs((CodeType)ce));
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

        private static IList<Declaration> GetSpringAliasesDeclarations()
        {
            string[] springAliases = new string[] { "int", "Integer", "int[]", "Integer()", "decimal", "Decimal", "decimal[]", 
                "Decimal()", "char", "Char", "char[]", "Char()", "long", "Long", "long[]", "Long()", "short", "Short", "short[]", 
                "Short()", "uint", "ulong", "ulong[]", "uint[]", "ushort", "ushort[]", "double", "Double", "double[]", "Double()", 
                "float", "Single", "float[]", "Single()", "DateTime", "date", "Date", "DateTime[]", "date[]", "DateTime()", "bool", 
                "Boolean", "bool[]", "Boolean()", "string", "String", "string[]", "String()", "object", "Object", "object[]", "Object()", 
                "int?", "int?[]", "decimal?", "decimal?[]", "char?", "char?[]", "long?", "long?[]", "short?", "short?[]", "uint?", 
                "ulong?", "ulong?[]", "uint?[]", "ushort?", "ushort?[]", "double?", "double?[]", "float?", "float?[]", "bool?", "bool?[]" };

            List<Declaration> declarations = new List<Declaration>();

            declarations.AddRange(springAliases.Select(
                alias => new Declaration(alias, alias, Declaration.DeclarationType.Alias, "Spring.NET standard type alias")));
            
            return declarations;
        }

        private IList<Declaration> GetProjectDeclarations(string text, bool addInterfaces)
        {
            List<Declaration> declarations = new List<Declaration>();

            DTE dte = this.serviceProvider.GetService(typeof(DTE)) as DTE;
            Project prj = dte.ActiveDocument.ProjectItem.ContainingProject;
            if (prj.CodeModel != null)
            {
                CodeElements codeElements = GetCodeElements(text, prj.CodeModel.CodeElements);
                if (codeElements != null)
                {
                    declarations.AddRange(GetCodeElementsDeclarations(prj, codeElements, addInterfaces));
                }
            }

            return declarations;
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

        private static IList<Declaration> GetCodeElementsDeclarations(Project prj, CodeElements codeElements, bool addInterfaces)
        {
            List<Declaration> declarations = new List<Declaration>();

            foreach (CodeElement ce in codeElements)
            {
                if (ce is CodeNamespace)
                {
                    declarations.Add(new Declaration(
                            ce.Name,
                            ce.Name,
                            Declaration.DeclarationType.Namespace,
                            ce.FullName));
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
                            declarations.Add(new Declaration(
                                ci.Name,
                                String.Format("{0}, {1}", ci.Name, assemblyName),
                                Declaration.DeclarationType.Interface,
                                String.Format("{0}, {1}{2}{3}", ci.FullName, assemblyName, Environment.NewLine, GetSummaryFromComment(comment))));
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
                        declarations.Add(new Declaration(
                            cc.Name,
                            String.Format("{0}, {1}", cc.Name, assemblyName),
                            Declaration.DeclarationType.Class,
                            String.Format("{0}, {1}{2}{3}", cc.FullName, assemblyName, Environment.NewLine, GetSummaryFromComment(comment))));

                    }
                }
            }

            return declarations;
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
            //return comment;
            if (!String.IsNullOrWhiteSpace(comment))
            {
                try
                {
                    System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                    xmlDoc.LoadXml(comment);
                    return xmlDoc.SelectSingleNode("//summary").InnerText.Trim();
                }
                catch (Exception)
                {
                    return comment;
                }
            }
            return String.Empty;
        }

        private CompletionSet GetCompletions(IList<Declaration> declarations, ICompletionSession session, bool isSnippet)
        {
            // Add Spring.NET completion
            List<SpringCompletion> completions = new List<SpringCompletion>();
            completions.AddRange(declarations.Select(declaration => new SpringCompletion(declaration, this.glyphService)));

            completions.Sort();

            return
                new SpringCompletionSet("Spring.NET",
                    "Spring.NET",
                    isSnippet ? CreateTrackingSpanForSnippet(session) : CreateTrackingSpan(session),
                    completions,
                    null)
            ;
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