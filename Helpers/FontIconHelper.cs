using System.Drawing;
using System.Drawing.Text;

namespace TrayFolder.Helpers
{
    public static class FontIconHelper
    {
        public static Bitmap CreateIcon(string glyph, float size = 6, Color? color = null)
        {
            // Standard icon size for context menu is 16x16
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                // High quality rendering
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                
                // Use Segoe MDL2 Assets
                using (var font = new Font("Segoe MDL2 Assets", size))
                using (var brush = new SolidBrush(color ?? Color.Black))
                using (var format = new StringFormat())
                {
                    format.Alignment = StringAlignment.Center;
                    format.LineAlignment = StringAlignment.Center;
                    
                    g.DrawString(glyph, font, brush, new RectangleF(0, 0, 16, 16), format);
                }
            }
            return bmp;
        }
    }
}
