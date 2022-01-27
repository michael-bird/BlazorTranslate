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
/// A read-only collection of variable declarators.
/// </summary>
using System;
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class VariableDeclaratorCollection : CommaDelimitedTreeCollection<VariableDeclarator>
    {

        /// <summary>
    /// Constructs a new collection of variable declarators.
    /// </summary>
    /// <param name="variableDeclarators">The variable declarators in the collection.</param>
    /// <param name="commaLocations">The locations of the commas in the list.</param>
    /// <param name="span">The location of the parse tree.</param>
        public VariableDeclaratorCollection(IList<VariableDeclarator> variableDeclarators, IList<Location> commaLocations, Span span) : base(TreeType.VariableDeclaratorCollection, variableDeclarators, commaLocations, span)
        {
            if (variableDeclarators is null || variableDeclarators.Count == 0)
            {
                throw new ArgumentException("VariableDeclaratorCollection cannot be empty.");
            }
        }
    }
}