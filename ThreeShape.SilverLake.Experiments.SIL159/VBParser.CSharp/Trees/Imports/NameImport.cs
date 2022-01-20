﻿// 
// Visual Basic .NET Parser
// 
// Copyright (C) 2005, Microsoft Corporation. All rights reserved.
// 
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
// EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
// MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
// 

/// <summary>
/// A parse tree for an Imports statement for a name.
/// </summary>
using System;
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class NameImport : Import
    {
        private readonly TypeName _TypeName;

        /// <summary>
    /// The imported name.
    /// </summary>
        public TypeName TypeName
        {
            get
            {
                return _TypeName;
            }
        }

        /// <summary>
    /// Constructs a new name import parse tree.
    /// </summary>
    /// <param name="typeName">The name to import.</param>
    /// <param name="span">The location of the parse tree.</param>
        public NameImport(TypeName typeName, Span span) : base(TreeType.NameImport, span)
        {
            if (typeName is null)
            {
                throw new ArgumentNullException("typeName");
            }

            SetParent(typeName);
            _TypeName = typeName;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, TypeName);
        }
    }
}