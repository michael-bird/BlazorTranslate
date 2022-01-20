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
/// An integer literal.
/// </summary>
using System;

namespace Dlrsoft.VBScript.Parser
{
    public sealed class UnsignedIntegerLiteralToken : Token
    {
        private readonly ulong _Literal;
        private readonly TypeCharacter _TypeCharacter;  // The type character after the literal, if any
        private readonly IntegerBase _IntegerBase;      // The base of the literal

        /// <summary>
    /// The value of the literal.
    /// </summary>
        [CLSCompliant(false)]
        public ulong Literal
        {
            get
            {
                return _Literal;
            }
        }

        /// <summary>
    /// The type character of the literal.
    /// </summary>
        public TypeCharacter TypeCharacter
        {
            get
            {
                return _TypeCharacter;
            }
        }

        /// <summary>
    /// The integer base of the literal.
    /// </summary>
        public IntegerBase IntegerBase
        {
            get
            {
                return _IntegerBase;
            }
        }

        /// <summary>
    /// Constructs a new unsigned integer literal.
    /// </summary>
    /// <param name="literal">The literal value.</param>
    /// <param name="integerBase">The integer base of the literal.</param>
    /// <param name="typeCharacter">The type character of the literal.</param>
    /// <param name="span">The location of the literal.</param>
        [CLSCompliant(false)]
        public UnsignedIntegerLiteralToken(ulong literal, IntegerBase integerBase, TypeCharacter typeCharacter, Span span) : base(TokenType.UnsignedIntegerLiteral, span)
        {
            if (integerBase < IntegerBase.Decimal || integerBase > IntegerBase.Hexadecimal)
            {
                throw new ArgumentOutOfRangeException("integerBase");
            }

            if (typeCharacter != TypeCharacter.None && typeCharacter != TypeCharacter.UnsignedIntegerChar && typeCharacter != TypeCharacter.UnsignedLongChar && typeCharacter != TypeCharacter.UnsignedShortChar)


            {
                throw new ArgumentOutOfRangeException("typeCharacter");
            }

            _Literal = literal;
            _IntegerBase = integerBase;
            _TypeCharacter = typeCharacter;
        }
    }
}