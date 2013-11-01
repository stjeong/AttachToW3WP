// Guids.cs
// MUST match guids.h
using System;

namespace wwwsysnetpekr.AttachToW3WP
{
    static class GuidList
    {
        public const string guidAttachToW3WPPkgString = "21c8e832-0ac8-4e66-9de5-594c89569990";
        public const string guidAttachToW3WPCmdSetString = "50d58cc5-e5bf-4008-8201-2a095675e7fa";

        public static readonly Guid guidAttachToW3WPCmdSet = new Guid(guidAttachToW3WPCmdSetString);
    };
}