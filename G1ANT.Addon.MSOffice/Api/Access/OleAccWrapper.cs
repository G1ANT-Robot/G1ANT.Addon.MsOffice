using System;
using System.Runtime.InteropServices;

namespace G1ANT.Addon.MSOffice.Api.Access
{
    public static class OleAccWrapper
    {
        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(
            IntPtr hwnd,
            uint dwObjectID,
            ref Guid riid,
            out Microsoft.Office.Interop.Access.Application o
        );

    }
}
