#if DEBUG
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

[assembly: AssemblyConfiguration("DEBUG")]
#else
[assembly: AssemblyConfiguration("RELEASE")]
#endif

[assembly: AssemblyTitle("WinformsSandbox")]
[assembly: AssemblyProduct("Winforms Sandbox")]
[assembly: AssemblyCopyright("Copyright (C) 2024-2025 Simon Mourier. All rights reserved.")]
[assembly: AssemblyCulture("")]
[assembly: AssemblyDescription("Windows Sandbox Programmatically used from .NET Core Windows Forms")]
[assembly: AssemblyCompany("Simon Mourier")]
[assembly: Guid("5e7241c5-7b1a-4dc7-9330-2ac8987a852b")]
[assembly: SupportedOSPlatform("windows10.0.19041.0")]
