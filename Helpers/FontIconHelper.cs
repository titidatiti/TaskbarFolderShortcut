using System.Drawing;
using System.Drawing.Text;

namespace TaskbarFolderShortcut.Helpers
{
    public static class FontIconHelper
    {
        public static Bitmap CreateIcon(string glyph, float size = 10, Color? color = null)
        {
            // Standard icon size for context menu is 16x16
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                // High quality rendering
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                
                // Use Segoe MDL2 Assets, fallback to Segoe UI Symbol if needed
                using (var font = new Font("Segoe MDL2 Assets", size))
                using (var brush = new SolidBrush(color ?? Color.Black))
                {
                    // Measure string to center it
                    var textSize = g.MeasureString(glyph, font);
                    
                    // Calculate centered position
                    // Note: MeasureString adds some padding, so we might need to adjust slightly
                    float x = (16 - textSize.Width) / 2;
                    float y = (16 - textSize.Height) / 2;

                    g.DrawString(glyph, font, brush, x, y);
                }
            }
            return bmp;
        }
    }
}
