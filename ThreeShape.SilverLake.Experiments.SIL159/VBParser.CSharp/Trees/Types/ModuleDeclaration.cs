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
/// A parse tree for a Module declaration.
/// </summary>
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class ModuleDeclaration : BlockDeclaration
    {

        /// <summary>
    /// Constructs a new parse tree for a Module declaration.
    /// </summary>
    /// <param name="attributes">The attributes for the parse tree.</param>
    /// <param name="modifiers">The modifiers for the parse tree.</param>
    /// <param name="keywordLocation">The location of the keyword.</param>
    /// <param name="name">The name of the declaration.</param>
    /// <param name="declarations">The declarations in the block.</param>
    /// <param name="endStatement">The end block declaration, if any.</param>
    /// <param name="span">The location of the parse tree.</param>
    /// <param name="comments">The comments for the parse tree.</param>
        public ModuleDeclaration(AttributeBlockCollection attributes, ModifierCollection modifiers, Location keywordLocation, SimpleName name, DeclarationCollection declarations, EndBlockDeclaration endStatement, Span span, IList<Comment> comments) : base(TreeType.ModuleDeclaration, attributes, modifiers, keywordLocation, name, declarations, endStatement, span, comments)
        {
        }
    }
}