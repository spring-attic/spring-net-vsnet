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
using System.IO;
using System.Windows.Media;

using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.TextManager.Interop;
using System.Windows.Media.Imaging;

namespace Spring.VisualStudio.Completion
{
    /// <summary>
    /// Defines an item in a Spring.NET completion set.
    /// </summary>
    /// <author>Bruno Baia</author>
    internal class SpringCompletion : Microsoft.VisualStudio.Language.Intellisense.Completion, IComparable
    {
        private static ImageSource SpringLeafGlyph;

        static SpringCompletion()
        {
            using (Stream stream = typeof(SpringCompletion).Assembly.GetManifestResourceStream("Spring.Resources.SpringGlyph.ico"))
            {
                SpringLeafGlyph = BitmapFrame.Create(stream);
            }
        }

        internal SpringCompletion(IGlyphService glyphService, string shortcut, string title, string description, SpringCompletionType type)
            : base(title)
        {
            this.DisplayText = shortcut;
            this.InsertionText = title;
            this.Description = description;
            this.IconSource = this.GetIconSource(glyphService, type);

            this.Type = type;
        }

        internal SpringCompletion(IGlyphService glyphService, VsExpansion vsExpansion)
            : base(vsExpansion.title)
        {
            this.DisplayText = vsExpansion.shortcut;
            this.InsertionText = vsExpansion.title;
            this.Description = vsExpansion.description;
            this.IconSource = this.GetIconSource(glyphService, SpringCompletionType.Snippet);

            this.VsExpansion = vsExpansion;
        }

        private ImageSource GetIconSource(IGlyphService glyphService, SpringCompletionType type)
        {
            switch (type)
            {
                case SpringCompletionType.Namespace:
                case SpringCompletionType.Class:
                case SpringCompletionType.Interface:
                case SpringCompletionType.Property:
                case SpringCompletionType.ConstructorArg:
                case SpringCompletionType.EnumMember:
                case SpringCompletionType.Boolean:
                    return glyphService.GetGlyph(GetGroupFromDeclaration(type), GetScopeFromDeclaration(type));
                case SpringCompletionType.Snippet:
                case SpringCompletionType.Alias:
                    return SpringLeafGlyph;
                default:
                    return glyphService.GetGlyph(StandardGlyphGroup.GlyphGroupClass, StandardGlyphItem.GlyphItemPublic);
            }
        }

        private StandardGlyphItem GetScopeFromDeclaration(SpringCompletionType declarationType)
        {
            return StandardGlyphItem.GlyphItemPublic;
        }

        private StandardGlyphGroup GetGroupFromDeclaration(SpringCompletionType declarationType)
        {
            switch (declarationType)
            {
                case SpringCompletionType.Namespace:
                    return StandardGlyphGroup.GlyphGroupNamespace;
                case SpringCompletionType.Class:
                    return StandardGlyphGroup.GlyphGroupClass;
                case SpringCompletionType.Interface:
                    return StandardGlyphGroup.GlyphGroupInterface;
                case SpringCompletionType.Property:
                    return StandardGlyphGroup.GlyphGroupProperty;
                case SpringCompletionType.ConstructorArg:
                    return StandardGlyphGroup.GlyphGroupField;
                case SpringCompletionType.EnumMember:
                    return StandardGlyphGroup.GlyphGroupEnumMember;
                case SpringCompletionType.Boolean:
                    return StandardGlyphGroup.GlyphGroupValueType;
                case SpringCompletionType.Alias:
                    return StandardGlyphGroup.GlyphGroupConstant;
                case SpringCompletionType.Snippet:
                    return StandardGlyphGroup.GlyphCSharpExpansion;
                default:
                    return StandardGlyphGroup.GlyphGroupClass;
            }
        }

        internal SpringCompletionType? Type { get; private set; }

        internal VsExpansion? VsExpansion { get; private set; }

        public int CompareTo(object other)
        {
            SpringCompletion otherCompletion = other as SpringCompletion;
            return this.DisplayText.CompareTo(otherCompletion.DisplayText);
        }
    }
}
