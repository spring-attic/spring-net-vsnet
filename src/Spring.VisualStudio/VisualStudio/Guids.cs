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

// Guids.cs
// MUST match guids.h
using System;

namespace Spring.VisualStudio
{
    static class GuidList
    {
        public const string guidPackagePkgString = "2229c8dc-a846-42d0-8bce-33f5f23edf27";
        public const string guidPackageCmdSetString = "9a38fd0c-d5c6-4509-90b5-e3fa95c047af";
        public const string guidSpringLanguageString = "325c7a3a-2bf9-4681-aa91-9c353262c5d4";

        public static readonly Guid guidPackageCmdSet = new Guid(guidPackageCmdSetString);
        public static readonly Guid guidSpringLanguage = new Guid(guidSpringLanguageString);
    };
}