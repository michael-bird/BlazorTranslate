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
/// A parse tree for an expression initializer.
/// </summary>
using System;
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class ExpressionInitializer : Initializer
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
    /// Constructs a new expression initializer parse tree.
    /// </summary>
    /// <param name="expression">The expression.</param>
    /// <param name="span">The location of the parse tree.</param>
        public ExpressionInitializer(Expression expression, Span span) : base(TreeType.ExpressionInitializer, span)
        {
            if (expression is null)
            {
                throw new ArgumentNullException("expression");
            }

            SetParent(expression);
            _Expression = expression;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, Expression);
        }
    }
}