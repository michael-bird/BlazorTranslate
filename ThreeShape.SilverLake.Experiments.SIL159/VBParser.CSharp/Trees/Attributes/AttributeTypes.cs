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
/// The type of an attribute usage.
/// </summary>
using System;

namespace Dlrsoft.VBScript.Parser
{
    [Flags()]
    public enum AttributeTypes
    {
        /// <summary>Regular application.</summary>
        Regular = 0x1,

        /// <summary>Applied to the netmodule.</summary>
        Module = 0x2,

        /// <summary>Applied to the assembly.</summary>
        Assembly = 0x4
    }
}