using Microsoft.Office.Interop.Access;
using System.Drawing;
using static G1ANT.Language.RobotWin32;

namespace G1ANT.Addon.MSOffice.Models.Access.Printers
{
    internal class AccessPrinterModel : INameModel
    {
        public Rect Margin { get; }
        public AcPrintColor ColorMode { get; }
        public string Port { get; }
        public AcPrintObjQuality PrintQuality { get; }
        public int RowSpacing { get; }
        public AcPrintItemLayout ItemLayout { get; }
        public int ItemsAcross { get; }
        public Size ItemSize { get; }
        public AcPrintOrientation Orientation { get; }
        public AcPrintPaperBin PaperBin { get; }
        public AcPrintPaperSize PaperSize { get; }
        public int ColumnSpacing { get; }
        public int Copies { get; }
        public bool DataOnly { get; }
        public bool DefaultSize { get; }
        public string Name { get; }
        public string DeviceName { get; }
        public string DriverName { get; }
        public AcPrintDuplex Duplex { get; }

        public AccessPrinterModel(Printer printer)
        {

            Margin = new Rect()
            {
                Bottom = printer.BottomMargin,
                Right = printer.RightMargin,
                Left = printer.LeftMargin,
                Top = printer.TopMargin
            };
            ColorMode = printer.ColorMode;
            ColumnSpacing = printer.ColumnSpacing;
            Copies = printer.Copies;
            DataOnly = printer.DataOnly;
            DefaultSize = printer.DefaultSize;
            Name = printer.DeviceName;
            DeviceName = printer.DeviceName;
            DriverName = printer.DriverName;
            Duplex = printer.Duplex;
            ItemLayout = printer.ItemLayout;
            ItemsAcross = printer.ItemsAcross;
            ItemSize = new Size(printer.ItemSizeWidth, printer.ItemSizeHeight);
            Orientation = printer.Orientation;
            PaperBin = printer.PaperBin;
            PaperSize = printer.PaperSize;
            Port = printer.Port;
            PrintQuality = printer.PrintQuality;
            RowSpacing = printer.RowSpacing;
        }

        public override string ToString() => Name;
    }
}
