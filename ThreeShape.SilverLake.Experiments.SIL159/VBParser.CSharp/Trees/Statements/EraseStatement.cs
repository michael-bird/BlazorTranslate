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
/// A parse tree for an Erase statement.
/// </summary>
using System;
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class EraseStatement : Statement
    {
        private readonly ExpressionCollection _Variables;

        /// <summary>
    /// The variables to erase.
    /// </summary>
        public ExpressionCollection Variables
        {
            get
            {
                return _Variables;
            }
        }

        /// <summary>
    /// Constructs a new parse tree for an Erase statement.
    /// </summary>
    /// <param name="variables">The variables to erase.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments for the parse tree.</param>
        public EraseStatement(ExpressionCollection variables, Span span, IList<Comment> comments) : base(TreeType.EraseStatement, span, comments)
        {
            if (variables is null)
            {
                throw new ArgumentNullException("variables");
            }

            SetParent(variables);
            _Variables = variables;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, Variables);
        }
    }
}