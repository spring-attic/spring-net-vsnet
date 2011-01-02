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
using System.Text;
using System.Collections.Generic;

using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;

namespace Spring.VisualStudio.Completion
{
    /// <summary>
    /// Represents a Spring.NET set of completions.
    /// </summary>
    /// <author>Bruno Baia</author>
    public class SpringCompletionSet : CompletionSet
    {
        private string _filterBufferText;
        private FilteredObservableCollection<Microsoft.VisualStudio.Language.Intellisense.Completion> _filteredCompletionBuilders;
        private FilteredObservableCollection<Microsoft.VisualStudio.Language.Intellisense.Completion> _filteredCompletions;

        public override IList<Microsoft.VisualStudio.Language.Intellisense.Completion> CompletionBuilders
        {
            get
            {
                return this._filteredCompletionBuilders;
            }
        }

        public override IList<Microsoft.VisualStudio.Language.Intellisense.Completion> Completions
        {
            get
            {
                return this._filteredCompletions;
            }
        }

        public SpringCompletionSet(string moniker, string displayName, ITrackingSpan applicableTo, IEnumerable<Microsoft.VisualStudio.Language.Intellisense.Completion> completions, IEnumerable<Microsoft.VisualStudio.Language.Intellisense.Completion> completionBuilders)
            : base(moniker, displayName, applicableTo, completions, completionBuilders)
        {
            this._filteredCompletions = new FilteredObservableCollection<Microsoft.VisualStudio.Language.Intellisense.Completion>(this.WritableCompletions);
            this._filteredCompletionBuilders = new FilteredObservableCollection<Microsoft.VisualStudio.Language.Intellisense.Completion>(this.WritableCompletionBuilders);
        }

        public override void Filter()
        {
            ITextSnapshot currentSnapshot = this.ApplicableTo.TextBuffer.CurrentSnapshot;
            this._filterBufferText = this.ApplicableTo.GetText(currentSnapshot);
            if (String.IsNullOrEmpty(this._filterBufferText))
            {
                this._filteredCompletions.StopFiltering();
                this._filteredCompletionBuilders.StopFiltering();
            }
            else
            {
                this._filteredCompletions.Filter(new Predicate<Microsoft.VisualStudio.Language.Intellisense.Completion>(this.DoesCompletionMatchApplicabilityText));
                this._filteredCompletionBuilders.Filter(new Predicate<Microsoft.VisualStudio.Language.Intellisense.Completion>(this.DoesCompletionMatchApplicabilityText));
            }
        }

        private bool DoesCompletionMatchApplicabilityText(Microsoft.VisualStudio.Language.Intellisense.Completion completion)
        {
            if (IsUpper(this._filterBufferText))
            {
                return GetUpperString(completion.DisplayText).StartsWith(this._filterBufferText);
            }

            return completion.DisplayText.ToLowerInvariant().Contains(this._filterBufferText.ToLowerInvariant());
        }

        public static bool IsUpper(string value)
        {
            for (int i = 0; i < value.Length; i++)
            {
                if (char.IsLower(value[i]))
                {
                    return false;
                }
            }
            return true;
        }

        public static string GetUpperString(string value)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < value.Length; i++)
            {
                if (char.IsUpper(value[i]))
                {
                    sb.Append(value[i]);
                }
            }
            return sb.ToString();
        } 
    }
}
