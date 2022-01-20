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
/// A parse tree for a Using block statement.
/// </summary>
using System;
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class UsingBlockStatement : ExpressionBlockStatement
    {
        private readonly VariableDeclaratorCollection _VariableDeclarators;

        /// <summary>
    /// The variable declarators, if no expression.
    /// </summary>
        public VariableDeclaratorCollection VariableDeclarators
        {
            get
            {
                return _VariableDeclarators;
            }
        }

        /// <summary>
    /// Constructs a new parse tree for a Using statement block with an expression.
    /// </summary>
    /// <param name="expression">The expression.</param>
    /// <param name="statements">The statements in the block.</param>
    /// <param name="endStatement">The End statement for the block, if any.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments for the parse tree.</param>
        public UsingBlockStatement(Expression expression, StatementCollection statements, EndBlockStatement endStatement, Span span, IList<Comment> comments) : base(TreeType.UsingBlockStatement, expression, statements, endStatement, span, comments)
        {
        }

        /// <summary>
    /// Constructs a new parse tree for a Using statement block with variable declarators.
    /// </summary>
    /// <param name="variableDeclarators">The variable declarators.</param>
    /// <param name="statements">The statements in the block.</param>
    /// <param name="endStatement">The End statement for the block, if any.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments for the parse tree.</param>
        public UsingBlockStatement(VariableDeclaratorCollection variableDeclarators, StatementCollection statements, EndBlockStatement endStatement, Span span, IList<Comment> comments) : base(TreeType.UsingBlockStatement, null, statements, endStatement, span, comments)
        {
            if (variableDeclarators is null)
            {
                throw new ArgumentNullException("variableDeclarators");
            }

            SetParent(variableDeclarators);
            _VariableDeclarators = variableDeclarators;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, VariableDeclarators);
            base.GetChildTrees(childList);
        }
    }
}