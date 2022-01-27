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
/// A parse tree for Me, MyBase or MyClass.
/// </summary>
using System;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class InstanceExpression : Expression
    {
        private InstanceType _InstanceType;

        /// <summary>
    /// The type of the instance expression.
    /// </summary>
        public InstanceType InstanceType
        {
            get
            {
                return _InstanceType;
            }
        }

        /// <summary>
    /// Constructs a new parse tree for My, MyBase or MyClass.
    /// </summary>
    /// <param name="instanceType">The type of the instance expression.</param>
    /// <param name="span">The location of the parse tree.</param>
        public InstanceExpression(InstanceType instanceType, Span span) : base(TreeType.InstanceExpression, span)
        {
            if (instanceType < InstanceType.Me || instanceType > InstanceType.MyBase)
            {
                throw new ArgumentOutOfRangeException("instanceType");
            }

            _InstanceType = instanceType;
        }
    }
}