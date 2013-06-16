// Guids.cs
// MUST match guids.h
using System;

namespace Company.AttachIISPKG
{
    static class GuidList
    {
        public const string guidAttachIISPKGPkgString = "d734a8d0-a619-46b2-bd1b-d847cdef91c7";
        public const string guidAttachIISPKGCmdSetString = "a9911550-fc2b-4711-b653-06d3ebe15d47";

        public static readonly Guid guidAttachIISPKGCmdSet = new Guid(guidAttachIISPKGCmdSetString);
    };
}