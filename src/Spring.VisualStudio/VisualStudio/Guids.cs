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

// Guids.cs
// MUST match guids.h
using System;

namespace Spring.VisualStudio
{
    static class GuidList
    {
        public const string guidPackagePkgString = "11d948f0-ba16-4557-920e-c37436a0204b";
        public const string guidPackageCmdSetString = "9eb89028-d74c-42fe-a8a7-eb0cc4e7f6d5";
        public const string guidSpringLanguageString = "418f0bdf-4d69-4581-b97c-84fbbf099dbe";

        public static readonly Guid guidPackageCmdSet = new Guid(guidPackageCmdSetString);
        public static readonly Guid guidSpringLanguage = new Guid(guidSpringLanguageString);
    };
}