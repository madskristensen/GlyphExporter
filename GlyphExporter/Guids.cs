// Guids.cs
// MUST match guids.h
using System;

namespace MadsKristensen.GlyphExporter
{
    static class GuidList
    {
        public const string guidGlyphExporterPkgString = "c1e1a7aa-d416-4600-81ce-a0aba0a3b201";
        public const string guidGlyphExporterCmdSetString = "3fc922ee-7ea1-456e-86df-f5a5ad462cdc";

        public static readonly Guid guidGlyphExporterCmdSet = new Guid(guidGlyphExporterCmdSetString);
    }

    static class PkgCmdIDList
    {
        public const uint cmdidMyCommand = 0x100;
    }
}