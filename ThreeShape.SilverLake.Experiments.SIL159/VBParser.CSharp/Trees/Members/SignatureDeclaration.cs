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
/// A parse tree for a declaration with a signature.
/// </summary>
using System.Collections.Generic;

namespace Dlrsoft.VBScript.Parser
{
    public abstract class SignatureDeclaration : ModifiedDeclaration
    {
        private readonly Location _KeywordLocation;
        private readonly SimpleName _Name;
        private readonly TypeParameterCollection _TypeParameters;
        private readonly ParameterCollection _Parameters;
        private readonly Location _AsLocation;
        private readonly AttributeBlockCollection _ResultTypeAttributes;
        private readonly TypeName _ResultType;

        /// <summary>
    /// The location of the declaration's keyword.
    /// </summary>
        public Location KeywordLocation
        {
            get
            {
                return _KeywordLocation;
            }
        }

        /// <summary>
    /// The name of the declaration.
    /// </summary>
        public SimpleName Name
        {
            get
            {
                return _Name;
            }
        }

        /// <summary>
    /// The type parameters on the declaration, if any.
    /// </summary>
        public TypeParameterCollection TypeParameters
        {
            get
            {
                return _TypeParameters;
            }
        }

        /// <summary>
    /// The parameters of the declaration.
    /// </summary>
        public ParameterCollection Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        /// <summary>
    /// The location of the 'As', if any.
    /// </summary>
        public Location AsLocation
        {
            get
            {
                return _AsLocation;
            }
        }

        /// <summary>
    /// The result type attributes, if any.
    /// </summary>
        public AttributeBlockCollection ResultTypeAttributes
        {
            get
            {
                return _ResultTypeAttributes;
            }
        }

        /// <summary>
    /// The result type, if any.
    /// </summary>
        public TypeName ResultType
        {
            get
            {
                return _ResultType;
            }
        }

        protected SignatureDeclaration(TreeType type, AttributeBlockCollection attributes, ModifierCollection modifiers, Location keywordLocation, SimpleName name, TypeParameterCollection typeParameters, ParameterCollection parameters, Location asLocation, AttributeBlockCollection resultTypeAttributes, TypeName resultType, Span span, IList<Comment> comments) : base(type, attributes, modifiers, span, comments)
        {
            SetParent(name);
            SetParent(typeParameters);
            SetParent(parameters);
            SetParent(resultType);
            SetParent(resultTypeAttributes);
            _KeywordLocation = keywordLocation;
            _Name = name;
            _TypeParameters = typeParameters;
            _Parameters = parameters;
            _AsLocation = asLocation;
            _ResultTypeAttributes = resultTypeAttributes;
            _ResultType = resultType;
        }

        protected override void GetChildTrees(IList<Tree> childList)
        {
            base.GetChildTrees(childList);
            AddChild(childList, Name);
            AddChild(childList, TypeParameters);
            AddChild(childList, Parameters);
            AddChild(childList, ResultTypeAttributes);
            AddChild(childList, ResultType);
        }
    }
}