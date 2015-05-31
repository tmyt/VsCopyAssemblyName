// Guids.cs
// MUST match guids.h
using System;

namespace CopyAssemblyNameCommand
{

    internal static class GuidList
    {
        public const string guidCopyAssemblyNameCommandPkgString = "c8fd9856-9afe-42fc-a293-d1072c94e770";

        public const string guidCopyAssemblyNameCommandSetString = "4e78a8d1-3732-4dab-879c-c5da77c444ee";

        public static readonly Guid guidCopyAssemblyNameCommandSet = new Guid(guidCopyAssemblyNameCommandSetString);
    };
}
