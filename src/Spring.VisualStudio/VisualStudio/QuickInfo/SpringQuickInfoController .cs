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
using System.Collections.Generic;

using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;

namespace Spring.VisualStudio.QuickInfo
{
    /// <summary>
    /// Triggers the QuickInfo process for the Spring.NET XML configuration file.
    /// </summary>
    /// <author>Bruno Baia</author>
    public class SpringQuickInfoController : IIntellisenseController
    {
        private ITextView textView;
        private IList<ITextBuffer> textBuffers;
        private IQuickInfoBroker quickInfoBroker;
        private IQuickInfoSession quickInfoSession;

        internal SpringQuickInfoController(ITextView textView, IList<ITextBuffer> textBuffers, IQuickInfoBroker quickInfoBroker)
        {
            this.textView = textView;
            this.textBuffers = textBuffers;
            this.quickInfoBroker = quickInfoBroker;

            this.textView.MouseHover += new EventHandler<MouseHoverEventArgs>(OnTextViewMouseHover);
        }

        private void OnTextViewMouseHover(object sender, MouseHoverEventArgs e)
        {
            //find the mouse position by mapping down to the subject buffer
            SnapshotPoint? point = this.textView.BufferGraph.MapDownToFirstMatch(
                new SnapshotPoint(this.textView.TextSnapshot, e.Position), 
                PointTrackingMode.Positive,
                snapshot => this.textBuffers.Contains(snapshot.TextBuffer), 
                PositionAffinity.Predecessor);

            if (point != null)
            {
                ITrackingPoint triggerPoint = point.Value.Snapshot.CreateTrackingPoint(
                    point.Value.Position, PointTrackingMode.Positive);

                if (!this.quickInfoBroker.IsQuickInfoActive(this.textView))
                {
                    this.quickInfoSession = this.quickInfoBroker.TriggerQuickInfo(this.textView, triggerPoint, true);
                }
            }
        }

        #region IIntellisenseController Members

        public void ConnectSubjectBuffer(ITextBuffer subjectBuffer)
        {
        }

        public void Detach(ITextView textView)
        {
            if (this.textView == textView)
            {
                this.textView.MouseHover -= this.OnTextViewMouseHover;
                this.textView = null;
            }
        }

        public void DisconnectSubjectBuffer(ITextBuffer subjectBuffer)
        {
        }

        #endregion
    }
}
