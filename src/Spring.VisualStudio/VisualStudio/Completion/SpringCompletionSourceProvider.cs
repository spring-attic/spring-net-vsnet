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

using System.ComponentModel.Composition;

using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Utilities;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;

namespace Spring.VisualStudio.Completion
{
    /// <summary>
    /// Provider for Spring.NET completion sources. 
    /// </summary>
    /// <author>Bruno Baia</author>
    [Export(typeof(ICompletionSourceProvider))]
    [ContentType("xml")]
    [Name("Spring.NET Completion Source")]
    internal class SpringCompletionSourceProvider : ICompletionSourceProvider
    {
        [Import]
        internal IGlyphService GlyphService { get; set; }

        [Import]
        internal IVsEditorAdaptersFactoryService VsEditorAdaptersFactoryService { get; set; }

        [Import]
        internal ITextStructureNavigatorSelectorService TextStructureNavigatorSelectorService { get; set; }

        [Import(typeof(SVsServiceProvider))]
        internal System.IServiceProvider ServiceProvider { get; set; }

        public ICompletionSource TryCreateCompletionSource(ITextBuffer textBuffer)
        {
            return new SpringCompletionSource(textBuffer, GlyphService, VsEditorAdaptersFactoryService, TextStructureNavigatorSelectorService, ServiceProvider);
        }
    }
}
