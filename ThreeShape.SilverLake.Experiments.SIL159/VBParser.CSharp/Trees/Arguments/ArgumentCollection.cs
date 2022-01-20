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
/// A read-only collection of arguments.
/// </summary>
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class ArgumentCollection : CommaDelimitedTreeCollection<Argument>
    {
        private readonly Location _RightParenthesisLocation;

        /// <summary>
    /// The location of the ')'.
    /// </summary>
        public Location RightParenthesisLocation
        {
            get
            {
                return _RightParenthesisLocation;
            }
        }

        /// <summary>
    /// Constructs a new argument collection.
    /// </summary>
    /// <param name="arguments">The arguments in the collection.</param>
    /// <param name="commaLocations">The location of the commas in the collection.</param>
    /// <param name="rightParenthesisLocation">The location of the ')'.</param>
    /// <param name="span">The location of the parse tree.</param>
        public ArgumentCollection(IList<Argument> arguments, IList<Location> commaLocations, Location rightParenthesisLocation, Span span) : base(TreeType.ArgumentCollection, arguments, commaLocations, span)
        {
            _RightParenthesisLocation = rightParenthesisLocation;
        }
    }
}