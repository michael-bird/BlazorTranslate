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
/// A parse tree for a label declaration statement.
/// </summary>
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class LabelStatement : LabelReferenceStatement
    {

        /// <summary>
    /// Constructs a parse tree for a label declaration statement.
    /// </summary>
    /// <param name="name">The label to branch to, if any.</param>
    /// <param name="isLineNumber">Whether the label is a line number.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments for the parse tree.</param>
        public LabelStatement(SimpleName name, bool isLineNumber, Span span, IList<Comment> comments) : base(TreeType.LabelStatement, name, isLineNumber, span, comments)
        {
        }
    }
}