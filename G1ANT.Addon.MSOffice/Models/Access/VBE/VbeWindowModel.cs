using Microsoft.Vbe.Interop;

namespace G1ANT.Addon.MSOffice.Models.Access.VBE
{
    internal class VbeWindowModel
    {
        public string Caption { get; }
        public bool Visible { get; }
        public int PixelWidth { get; }
        public int PixelHeight { get; }
        public vbext_WindowType Type { get; }
        public int PixelLeft { get; }
        public int PixelTop { get; }

        public VbeWindowModel(Window window)
        {
            try
            {
                Caption = window.Caption;
                Visible = window.Visible;
                PixelWidth = window.Width;
                PixelHeight = window.Height;
                Type = window.Type;
                PixelLeft = window.Left;
                PixelTop = window.Top;
            }
            catch { }
        }

        public override string ToString() => $"{Caption} {Type} {PixelLeft}/{PixelTop} {PixelWidth}x{PixelHeight}";
    }
}
