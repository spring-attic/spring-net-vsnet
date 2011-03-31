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
    /// Provides an item in a Spring.NET completion set.
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

        internal SpringCompletion(Declaration declaration, IGlyphService glyphService)
            : base(declaration.Title)
        {
            this.DisplayText = declaration.Shortcut;
            this.InsertionText = declaration.Title;
            this.Description = declaration.Description;
            this.IconSource = this.GetIconSource(glyphService, declaration.Type);

            this.DeclarationType = declaration.Type;
        }

        internal SpringCompletion(VsExpansion vsExpansion, IGlyphService glyphService)
            : base(vsExpansion.title)
        {
            this.DisplayText = vsExpansion.shortcut;
            this.InsertionText = vsExpansion.title;
            this.Description = vsExpansion.description;
            this.IconSource = this.GetIconSource(glyphService, Declaration.DeclarationType.Snippet);

            this.VsExpansion = vsExpansion;
        }

        private ImageSource GetIconSource(IGlyphService glyphService, Declaration.DeclarationType declarationType)
        {
            switch (declarationType)
            {
                case Declaration.DeclarationType.Namespace:
                case Declaration.DeclarationType.Class:
                case Declaration.DeclarationType.Interface:
                case Declaration.DeclarationType.Property:
                case Declaration.DeclarationType.ConstructorArg:
                case Declaration.DeclarationType.EnumMember:
                case Declaration.DeclarationType.Boolean:
                    return glyphService.GetGlyph(GetGroupFromDeclaration(declarationType), GetScopeFromDeclaration(declarationType));
                case Declaration.DeclarationType.Snippet:
                case Declaration.DeclarationType.Alias:
                    return SpringLeafGlyph;
                default:
                    return glyphService.GetGlyph(StandardGlyphGroup.GlyphGroupClass, StandardGlyphItem.GlyphItemPublic);
            }
        }

        private StandardGlyphItem GetScopeFromDeclaration(Declaration.DeclarationType declarationType)
        {
            return StandardGlyphItem.GlyphItemPublic;
        }

        private StandardGlyphGroup GetGroupFromDeclaration(Declaration.DeclarationType declarationType)
        {
            switch (declarationType)
            {
                case Declaration.DeclarationType.Namespace:
                    return StandardGlyphGroup.GlyphGroupNamespace;
                case Declaration.DeclarationType.Class:
                    return StandardGlyphGroup.GlyphGroupClass;
                case Declaration.DeclarationType.Interface:
                    return StandardGlyphGroup.GlyphGroupInterface;
                case Declaration.DeclarationType.Property:
                    return StandardGlyphGroup.GlyphGroupProperty;
                case Declaration.DeclarationType.ConstructorArg:
                    return StandardGlyphGroup.GlyphGroupField;
                case Declaration.DeclarationType.EnumMember:
                    return StandardGlyphGroup.GlyphGroupEnumMember;
                case Declaration.DeclarationType.Boolean:
                    return StandardGlyphGroup.GlyphGroupValueType;
                case Declaration.DeclarationType.Alias:
                    return StandardGlyphGroup.GlyphGroupConstant;
                case Declaration.DeclarationType.Snippet:
                    return StandardGlyphGroup.GlyphCSharpExpansion;
                default:
                    return StandardGlyphGroup.GlyphGroupClass;
            }
        }

        internal Declaration.DeclarationType? DeclarationType { get; private set; }

        internal VsExpansion? VsExpansion { get; private set; }

        public int CompareTo(object other)
        {
            var otherCompletion = other as SpringCompletion;
            return this.DisplayText.CompareTo(otherCompletion.DisplayText);
        }
    }
}
