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
using System.ComponentModel.Composition;

using Microsoft.VisualStudio.Utilities;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;

namespace Spring.VisualStudio.Completion
{
    /// <summary>
    /// Provider for Spring.NET Completion Controller.
    /// </summary>
    /// <author>Bruno Baia</author>
    [Export(typeof(IVsTextViewCreationListener))]
    [Name("Spring.NET Completion Controller")]
    [ContentType("xml")]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    internal class SpringCompletionControllerProvider : IVsTextViewCreationListener
    {
        [Import]
        internal IVsEditorAdaptersFactoryService VsEditorAdaptersFactoryService { get; set; }

        [Import]
        internal ITextStructureNavigatorSelectorService TextStructureNavigatorSelectorService { get; set; }

        [Import]
        internal ICompletionBroker CompletionBroker { get; set; }

        [Import(typeof(SVsServiceProvider))]
        internal SVsServiceProvider ServiceProvider { get; set; }

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            ITextView textView = VsEditorAdaptersFactoryService.GetWpfTextView(textViewAdapter);
            if (textView == null)
            {
                return;
            }

            Func<SpringCompletionController> controllerCreator = delegate() { return new SpringCompletionController(ServiceProvider, textViewAdapter, textView, CompletionBroker, TextStructureNavigatorSelectorService); };
            textView.Properties.GetOrCreateSingletonProperty(controllerCreator);
        }
    }
}

