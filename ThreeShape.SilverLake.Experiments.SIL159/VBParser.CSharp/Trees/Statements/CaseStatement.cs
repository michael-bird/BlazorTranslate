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
/// A parse tree for a Case statement.
/// </summary>
using System;
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class CaseStatement : Statement
    {
        private readonly CaseClauseCollection _CaseClauses;

        /// <summary>
    /// The clauses in the Case statement.
    /// </summary>
        public CaseClauseCollection CaseClauses
        {
            get
            {
                return _CaseClauses;
            }
        }

        /// <summary>
    /// Constructs a new parse tree for a Case statement.
    /// </summary>
    /// <param name="caseClauses">The clauses in the Case statement.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments on the parse tree.</param>
        public CaseStatement(CaseClauseCollection caseClauses, Span span, IList<Comment> comments) : base(TreeType.CaseStatement, span, comments)
        {
            if (caseClauses is null)
            {
                throw new ArgumentNullException("caseClauses");
            }

            SetParent(caseClauses);
            _CaseClauses = caseClauses;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, CaseClauses);
        }
    }
}