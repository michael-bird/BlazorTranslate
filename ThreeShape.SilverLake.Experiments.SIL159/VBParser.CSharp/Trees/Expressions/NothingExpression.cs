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
/// A parse tree for Nothing.
/// </summary>

namespace Dlrsoft.VBScript.Parser
{
    public sealed class NothingExpression : Expression
    {
        public override bool IsConstant
        {
            get
            {
                return true;
            }
        }

        /// <summary>
    /// Constructs a new parse tree for Nothing.
    /// </summary>
    /// <param name="span">The location of the parse tree.</param>
        public NothingExpression(Span span) : base(TreeType.NothingExpression, span)
        {
        }
    }
}