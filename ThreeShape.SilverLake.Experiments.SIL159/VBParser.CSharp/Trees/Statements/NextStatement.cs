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
/// A parse tree for a Next statement.
/// </summary>
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class NextStatement : Statement
    {
        private readonly ExpressionCollection _Variables;

        /// <summary>
    /// The loop control variables.
    /// </summary>
        public ExpressionCollection Variables
        {
            get
            {
                return _Variables;
            }
        }

        /// <summary>
    /// Constructs a parse tree for a Next statement.
    /// </summary>
    /// <param name="variables">The loop control variables.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments for the parse tree.</param>
        public NextStatement(ExpressionCollection variables, Span span, IList<Comment> comments) : base(TreeType.NextStatement, span, comments)
        {
            SetParent(variables);
            _Variables = variables;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, Variables);
        }
    }
}