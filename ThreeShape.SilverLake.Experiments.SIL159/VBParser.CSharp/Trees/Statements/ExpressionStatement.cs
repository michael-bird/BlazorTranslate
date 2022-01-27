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
/// A parse tree for an expression statement.
/// </summary>
using System.Collections.Generic;
using System.Diagnostics;

namespace Dlrsoft.VBScript.Parser
{
    public abstract class ExpressionStatement : Statement
    {
        private readonly Expression _Expression;

        /// <summary>
    /// The expression.
    /// </summary>
        public Expression Expression
        {
            get
            {
                return _Expression;
            }
        }

        /// <summary>
    /// Constructs a new parse tree for an expression statement.
    /// </summary>
    /// <param name="type">The type of the parse tree.</param>
    /// <param name="expression">The expression.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments for the parse tree.</param>
        protected ExpressionStatement(TreeType type, Expression expression, Span span, IList<Comment> comments) : base(type, span, comments)
        {
            Debug.Assert(type == TreeType.ReturnStatement || type == TreeType.ErrorStatement || type == TreeType.ThrowStatement);
            SetParent(expression);
            _Expression = expression;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, Expression);
        }
    }
}