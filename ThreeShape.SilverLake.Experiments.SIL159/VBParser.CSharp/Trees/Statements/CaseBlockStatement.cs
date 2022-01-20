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
/// A parse tree for the block of a Case statement.
/// </summary>
using System;
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class CaseBlockStatement : BlockStatement
    {
        private readonly CaseStatement _CaseStatement;

        /// <summary>
    /// The Case statement that started the block.
    /// </summary>
        public CaseStatement CaseStatement
        {
            get
            {
                return _CaseStatement;
            }
        }

        /// <summary>
    /// Constructs a new parse tree for the block of a Case statement.
    /// </summary>
    /// <param name="caseStatement">The Case statement that started the block.</param>
    /// <param name="statements">The statements in the block.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments of the tree.</param>
        public CaseBlockStatement(CaseStatement caseStatement, StatementCollection statements, Span span, IList<Comment> comments) : base(TreeType.CaseBlockStatement, statements, span, comments)
        {
            if (caseStatement is null)
            {
                throw new ArgumentNullException("caseStatement");
            }

            SetParent(caseStatement);
            _CaseStatement = caseStatement;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            AddChild(childList, CaseStatement);
            base.GetChildTrees(childList);
        }
    }
}