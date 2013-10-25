// Guids.cs
// MUST match guids.h
using System;

namespace wwwsysnetpekr.AttachDebug
{
    static class GuidList
    {
        public const string guidAttachDebugPkgString = "21c8e832-0ac8-4e66-9de5-594c89569990";
        public const string guidAttachDebugCmdSetString = "50d58cc5-e5bf-4008-8201-2a095675e7fa";

        public static readonly Guid guidAttachDebugCmdSet = new Guid(guidAttachDebugCmdSetString);
    };
}