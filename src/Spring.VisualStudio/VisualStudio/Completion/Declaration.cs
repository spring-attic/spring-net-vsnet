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

namespace Spring.VisualStudio.Completion
{
    /// <summary>
    /// Provides information about an item in a Spring.NET completion set.
    /// </summary>
    /// <author>Bruno Baia</author>
    public class Declaration : IComparable
    {
        public enum DeclarationType
        {
            Namespace,
            Class,
            Interface,
            Property,
            ConstructorArg,
            EnumMember,
            Alias,
            Snippet,
            Unknown
        }

        public Declaration(string shortcut, string title, DeclarationType type, string description)
        {
            this.Type = type;
            this.shortcut = shortcut;
            this.title = title;
            this.description = description;
        }

        protected Declaration()
        {
        }

        public int CompareTo(object obj)
        {
            Declaration decl = (Declaration)obj;
            return (this.title.CompareTo(decl.title));
        }

        public override bool Equals(Object obj)
        {
            if (!(obj is Declaration))
                return false;
            return (this.CompareTo(obj) == 0);
        }

        public override int GetHashCode()
        {
            return this.title.GetHashCode();
        }

        // Disable the "IdentifiersShouldNotMatchKeywords" warning.
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062")]
        public static bool operator ==(Declaration d1, Declaration d2)
        {
            return d1.Equals(d2);
        }

        // Disable the "IdentifiersShouldNotMatchKeywords" warning.
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062")]
        public static bool operator !=(Declaration d1, Declaration d2)
        {
            return !(d1 == d2);
        }

        // Disable the "IdentifiersShouldNotMatchKeywords" warning.
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062")]
        public static bool operator <(Declaration d1, Declaration d2)
        {
            return (d1.CompareTo(d2) < 0);
        }

        // Disable the "IdentifiersShouldNotMatchKeywords" warning.
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062")]
        public static bool operator >(Declaration d1, Declaration d2)
        {
            return (d1.CompareTo(d2) > 0);
        }

        public DeclarationType Type
        {
            get { return type; }
            set { type = value; }
        }

        public string Description
        {
            get { return description; }
            set { description = value; }
        }

        public string Shortcut
        {
            get { return shortcut; }
            set { shortcut = value; }
        }

        public string Title
        {
            get { return title; }
            set { title = value; }
        }

        private DeclarationType type;
        private string description;
        private string shortcut;
        private string title;
    }
}
