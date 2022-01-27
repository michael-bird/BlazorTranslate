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
/// A parse tree for a RemoveHandler property accessor.
/// </summary>
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class RemoveHandlerAccessorDeclaration : ModifiedDeclaration
    {
        private readonly Location _RemoveHandlerLocation;
        private readonly ParameterCollection _Parameters;
        private readonly StatementCollection _Statements;
        private readonly EndBlockDeclaration _EndStatement;

        /// <summary>
    /// The location of the 'RemoveHandler'.
    /// </summary>
        public Location RemoveHandlerLocation
        {
            get
            {
                return _RemoveHandlerLocation;
            }
        }

        /// <summary>
    /// The accessor's parameters.
    /// </summary>
        public ParameterCollection Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        /// <summary>
    /// The statements in the accessor.
    /// </summary>
        public StatementCollection Statements
        {
            get
            {
                return _Statements;
            }
        }

        /// <summary>
    /// The End declaration for the accessor.
    /// </summary>
        public EndBlockDeclaration EndStatement
        {
            get
            {
                return _EndStatement;
            }
        }

        /// <summary>
    /// Constructs a new parse tree for a property accessor.
    /// </summary>
    /// <param name="attributes">The attributes for the parse tree.</param>
    /// <param name="removeHandlerLocation">The location of the 'RemoveHandler'.</param>
    /// <param name="parameters">The parameters of the declaration.</param>
    /// <param name="statements">The statements in the declaration.</param>
    /// <param name="endStatement">The end block declaration, if any.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments for the parse tree.</param>
        public RemoveHandlerAccessorDeclaration(AttributeBlockCollection attributes, Location removeHandlerLocation, ParameterCollection parameters, StatementCollection statements, EndBlockDeclaration endStatement, Span span, IList<Comment> comments) : base(TreeType.RemoveHandlerAccessorDeclaration, attributes, null, span, comments)
        {
            SetParent(parameters);
            SetParent(statements);
            SetParent(endStatement);
            _Parameters = parameters;
            _RemoveHandlerLocation = removeHandlerLocation;
            _Statements = statements;
            _EndStatement = endStatement;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            base.GetChildTrees(childList);
            AddChild(childList, Parameters);
            AddChild(childList, Statements);
            AddChild(childList, EndStatement);
        }
    }
}