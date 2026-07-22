using System.Drawing;
using System.Drawing.Text;
using System.Windows.Forms;

namespace OutlookListener
{
    // Shared TKE brand styling for the Rules Dashboard's WinForms UI, following the
    // Productive Application UI Kit (the authoritative section for internal/productive
    // apps, as opposed to the marketing-site component guidance): black is the default
    // action color, purple is reserved for hover/pressed feedback and "current position"
    // signals (e.g. the rules grid's selected-row highlight), not for default button fills.
    // The orange/purple gradient ("TKE Sunrise") is a marketing-site motif only - the Kit's
    // own Titlebar and Dialog reference pages (p75, p72) show a plain solid black bar with
    // white text, no gradient - so headers here are solid black, not gradient-filled.
    internal static class TkeTheme
    {
        public static readonly Color Purple = Color.FromArgb(0x70, 0x00, 0xBD);
        public static readonly Color Purple1 = Color.FromArgb(0x59, 0x00, 0x8D);
        public static readonly Color Purple2 = Color.FromArgb(0xA3, 0x1A, 0xFF);
        public static readonly Color Purple4 = Color.FromArgb(0xED, 0xD1, 0xFF);
        public static readonly Color Black = Color.FromArgb(0x00, 0x00, 0x00);
        public static readonly Color Gray1 = Color.FromArgb(0x26, 0x26, 0x26);
        public static readonly Color Gray5 = Color.FromArgb(0xE9, 0xE9, 0xE9);

        private const string FontBody = "TKE Type Book";
        private const string FontMedium = "TKE Type Medium";
        private const string FontLight = "TKE Type Light";
        private const string FontBold = "TKE Type Bold";

        public static Font Body(float size) { return new Font(FontBody, size, FontStyle.Regular, GraphicsUnit.Point); }
        public static Font Medium(float size) { return new Font(FontMedium, size, FontStyle.Regular, GraphicsUnit.Point); }
        public static Font Light(float size) { return new Font(FontLight, size, FontStyle.Regular, GraphicsUnit.Point); }
        public static Font Bold(float size) { return new Font(FontBold, size, FontStyle.Regular, GraphicsUnit.Point); }

        // Table Header (per the Kit's Tables page, p43): plain black text with a thin
        // underline, no fill - not a colored block.
        public static Font BoldUnderline(float size) { return new Font(FontBold, size, FontStyle.Underline, GraphicsUnit.Point); }

        // Draws a plain black titlebar (per the Kit's Titlebar/Header reference) with the
        // title/subtitle painted directly onto the Graphics surface. Deliberately not child
        // Label controls: a Label's "transparent" background is faked by reading its parent's
        // flat BackColor property, which produces a visible solid-color box wherever a parent
        // is custom-painted (e.g. anything other than a flat single-color fill) - drawing the
        // text directly in the same Paint handler that fills the background avoids that.
        public static void PaintHeader(Graphics g, Rectangle bounds, string title, float titleSize, int titleY, string subtitle, float subtitleSize, int subtitleY)
        {
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            using (var brush = new SolidBrush(Black))
            {
                g.FillRectangle(brush, bounds);
            }
            using (var titleFont = Medium(titleSize))
            {
                g.DrawString(title, titleFont, Brushes.White, 18, titleY);
            }
            if (!string.IsNullOrEmpty(subtitle))
            {
                using (var subtitleFont = Light(subtitleSize))
                {
                    g.DrawString(subtitle, subtitleFont, Brushes.White, 18, subtitleY);
                }
            }
        }

        // Primary = black fill / white text by default; hover = Purple 2; pressed = Purple 1.
        public static Button CreatePrimaryButton(string text, int left, int top, int width, int height)
        {
            return new Button
            {
                Text = text,
                Left = left,
                Top = top,
                Width = width,
                Height = height,
                FlatStyle = FlatStyle.Flat,
                BackColor = Black,
                ForeColor = Color.White,
                Font = Bold(9.5F),
                FlatAppearance =
                {
                    BorderColor = Black,
                    BorderSize = 0,
                    MouseOverBackColor = Purple2,
                    MouseDownBackColor = Purple1
                }
            };
        }

        // Secondary = white fill, black outline/text by default (per the Kit's Buttons page,
        // p12): hover keeps the white fill and turns the outline/text purple (no fill change);
        // only :pressed gets a light-purple (Purple 4) fill.
        public static Button CreateSecondaryButton(string text, int left, int top, int width, int height)
        {
            var btn = new Button
            {
                Text = text,
                Left = left,
                Top = top,
                Width = width,
                Height = height,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Black,
                Font = Bold(9.5F),
                FlatAppearance =
                {
                    BorderColor = Black,
                    BorderSize = 1,
                    MouseOverBackColor = Color.White,
                    MouseDownBackColor = Purple4
                }
            };
            btn.MouseEnter += (s, e) => { btn.ForeColor = Purple; btn.FlatAppearance.BorderColor = Purple; };
            btn.MouseLeave += (s, e) => { btn.ForeColor = Black; btn.FlatAppearance.BorderColor = Black; };
            return btn;
        }

        // Small square secondary-style button (the "..." file-browse buttons).
        public static Button CreateSecondaryButton(string text, int width, int height)
        {
            return CreateSecondaryButton(text, 0, 0, width, height);
        }
    }
}
