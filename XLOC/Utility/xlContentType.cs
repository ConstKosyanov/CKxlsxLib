using System;

namespace XLOC.Utility
{
    public enum XlContentType : byte
    {
        Void = 0,
        Boolean = 1,
        Integer = 2,
        Double = 3,
        SharedString = 4,
        String = 5,
        Date = 6,
    }
}