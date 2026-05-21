using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;

namespace ProofioAddIn.UI
{
    /// <summary>
    /// Central design tokens + shared helper controls for all Proofio dialogs.
    /// </summary>
    internal static class UITheme
    {
        // ── Colours ───────────────────────────────────────────────────────────
        public static readonly Color Navy         = Color.FromArgb(13,  27,  62);   // title bar
        public static readonly Color NavyLight    = Color.FromArgb(26,  45,  90);
        public static readonly Color Background   = Color.FromArgb(244, 246, 251);
        public static readonly Color Surface      = Color.White;
        public static readonly Color Border       = Color.FromArgb(216, 221, 232);
        public static readonly Color Accent       = Color.FromArgb(41,  82,  227);
        public static readonly Color AccentDark   = Color.FromArgb(26,  59,  191);
        public static readonly Color AccentLight  = Color.FromArgb(237, 240, 252);
        public static readonly Color TextPrimary  = Color.FromArgb(26,  45,  90);
        public static readonly Color TextSecondary= Color.FromArgb(53,  72, 114);
        public static readonly Color TextMuted    = Color.FromArgb(123, 138, 170);

        // ── Fonts ─────────────────────────────────────────────────────────────
        public static readonly Font DefaultFont   = new Font("Segoe UI", 9.5f);
        public static readonly Font TitleFont     = new Font("Segoe UI Semibold", 10.5f, FontStyle.Bold);
        public static readonly Font LabelFont     = new Font("Segoe UI", 8.5f);
        public static readonly Font SmallFont     = new Font("Segoe UI", 8.5f);
        public static readonly Font SmallBoldFont = new Font("Segoe UI Semibold", 8.5f, FontStyle.Bold);

        // ── Favicon / window icon (base64-embedded PNG) ───────────────────────
        internal const string TitleBarIconBase64 = "iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAABCGlDQ1BJQ0MgUHJvZmlsZQAAeJxjYGA8wQAELAYMDLl5JUVB7k4KEZFRCuwPGBiBEAwSk4sLGHADoKpv1yBqL+viUYcLcKakFicD6Q9ArFIEtBxopAiQLZIOYWuA2EkQtg2IXV5SUAJkB4DYRSFBzkB2CpCtkY7ETkJiJxcUgdT3ANk2uTmlyQh3M/Ck5oUGA2kOIJZhKGYIYnBncAL5H6IkfxEDg8VXBgbmCQixpJkMDNtbGRgkbiHEVBYwMPC3MDBsO48QQ4RJQWJRIliIBYiZ0tIYGD4tZ2DgjWRgEL7AwMAVDQsIHG5TALvNnSEfCNMZchhSgSKeDHkMyQx6QJYRgwGDIYMZAKbWPz9HbOBQAAAEL0lEQVR4nJWUXUybVRjH/885b98CLRTqgE0+CqUDBmNzwCRRsxqzqAmJOg1kUeJ0zLkti1Ez48dcuipLvNmNu1AvnEv0gsBi9G4m04VdkMWIMGVZhokZZNTx1RLa0vZ9e87jRRGYLDH+k5OcnPM8v/Oc5/xzCOs1wBLdpMz3os9kiwoHIQ0TBICwURLAUvL7U6fdz4dDDJwGg4iNe4Ju5FI1zF2cJkg7dUILEBgEAgPAyowJtEkXuN7tOxn/CmE6gGaWAGtjw8kAYAFk28vqbMFZvm8A4H51ojRZWvumdrtflm/Hl1U3HUWIjRwwxAJhrOVaGiQk6dDKDf4CYcvqvgCgMzPRcrIMJxY02OU+Io8sJlSY3jEAJoRJAwCiExIh1nR3mUAAwpTNMZgAygGZCURsHJyMsnZNkzQKOWNnke854eiNpQVAnN8z34Gu4Xycq88gTFpanJb2fZ8CoBw4dd4XKUpONRk605S3MN8oEqnPKa/4Q8p/MfoGTHO3yKoUiAegbKWNggMM2Zm+Mf4gRtpWqqR/tXNd1QCcPbFX4C4+LwxbtjgT+halqZnYcYi0s1cq2SotXsZIu70R9A8sJ9f+uzvATIKlU2RBIv5n/zEtEJVW+njiG9f+RH9Rj5FMX3TArPbsW+zydE6WAAC6WKKLJcCEICRA7OqOfyFEyXUQwdDOlLQBw+t5YZeMcnzuctnoOt8MkdJjwukZUArdAA1iEAoA0MYODJHteS5+Rpjuwzq+eAgAC8t2gk0YgmUfS/xatjf9A7P9CwOSlrQFil9jp6eFlUkA8MDepQvM2ZHoj3TO++TiSSHdH/B87FjskvdLAJC2VKwBITIccyRFOSk1zppvC+bbGjQNIa8JW5PDUhIAoNkyzZJPy55YumAYnj6Kxt5fuOT9zBfkPABAyiaRBRvkyhxUGcs/e2Xz7+vbXtERe1y7DaGyuhxgWviJDpfviWvTXfi6nVz6aHbI+wna2DGZgAKYpE56yQYJinB1XrLweGXr3EursPbYx0zmKWlnJwxlPAsQo4vlzNXCo2o+smdmyBNCiAVGyIYfGiAmmzqRSv1Btdtj/SC6qli8JqWc0QxiyvLUdc/Tla2xfUZB8bdWYvrRyFjlMIJsYIiyqx7sYolBUqW77zxkFlSMsrXUi0BgNtjQkBiorY+9Vb110l/ZNBUAQDmLhER1a2q4pt2KeHaOFgNAIMDO3D0GchZC0KjqyNysejj9W24tZzKJDWIBMAUCE5V1O6yFupb02Jb6K5sAoK2NHStBwrdz+Tt/q618zZFtKx8ni/WAtbG25/dPtTQ2ZGIN9Vakru7OIwDg841vrm9IDzVuZ65rmHlq1fz/rVyQr/xmzTZ/6udmP3OjL/F1Y01mrqnOmq+vmn7sf8DuhQKQTZXxMy0+5uaq5ctbvbcqACCItZ//b+PW42kKGz+4AAAAAElFTkSuQmCC";  // 20x20 transparent PNG
        internal const string FaviconBase64 = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAABCGlDQ1BJQ0MgUHJvZmlsZQAAeJxjYGA8wQAELAYMDLl5JUVB7k4KEZFRCuwPGBiBEAwSk4sLGHADoKpv1yBqL+viUYcLcKakFicD6Q9ArFIEtBxopAiQLZIOYWuA2EkQtg2IXV5SUAJkB4DYRSFBzkB2CpCtkY7ETkJiJxcUgdT3ANk2uTmlyQh3M/Ck5oUGA2kOIJZhKGYIYnBncAL5H6IkfxEDg8VXBgbmCQixpJkMDNtbGRgkbiHEVBYwMPC3MDBsO48QQ4RJQWJRIliIBYiZ0tIYGD4tZ2DgjWRgEL7AwMAVDQsIHG5TALvNnSEfCNMZchhSgSKeDHkMyQx6QJYRgwGDIYMZAKbWPz9HbOBQAAAIuklEQVR4nK1Xa2xc1RH+5pxz7778WMd2nMTPGK+D1wEXL4GIEmyr0EAIkIZuVFJoKUgFSqmgLVJ/oDqWopYWQYOEKlG1QlGBKnYLTUioJaRiQwBRYhITCDVOnMQxjuP3Y73e3XvPmf5Y2zHUiSPake6V7r0z8818Z86cuYSLSSMLoE0ot+bP7PdH2bAGswCD5nUIIDDm7nPvQGRIEJOTfDcw+uG3Jp5tGEcjCzSRWQghLgreRCZzqqDcqJzv6JRHMnltlj7Flk8ae/ayfNJYfmksn2Qr/czKJ1l6Lc0eW/uy6uNZta8WRZ/2YQcYjY1fwFQXZQCASZIHXs2CjWNNjT7MEsMQimCYL2hEJCSz0UJe7lg5O11/Vv250H2v1NXX39Ze32bAOwhEfEkBgJUBS4J2U8s+2/vy2f0PxJe0mRX78aE1YPyaJ13HDQRvfmfdq3uoie7kYywBNgDx0gEAgAZgiMaXXZODxjdTqK8H2touYpAvgCFjxikIAohJYVI7OpCzVT46+oLeRT/gZpbYxmbpAFIpkMOAZhClNJoaXOxgQkPDhZcg2izRsk3TQ0MuCRgYAGDFKdcx3px71YPDCWcbPYQoywsX4XwANuAAcAEgsKT6QmFNSjCEcKUQRpEwyqKYMSaQ+6B1/+BTaCGdZiDKEmEQ2tqA9gb3y46EIbABEJ++NOTmqAEx2clPTjvG2wXpLSTWs9uUJdjj40D+z+3vD3xOABNAC+gkoPGXAqgXQJsJnLgn7AbKj7JJTavx4xXxlrUDYJ6v4qWktO4F71DZpiDicTCkZJEUHMh70mQE76J47BwBgH/byU2cVfAwJWba4y/m/nahg8y7jocc32WfgRPTIvZRRbzl2oH/DvpCsried/upr3NW6UF2klr5o6cjbOfuQdJtgfQ/nnH3COnp7peMClrCTTom5ZbAw4ChxRCWEGKACY2QOAsCOoCVr2nr32DHBVgTKSEyb0LK9MSas+4LRAcfQzD/GamrfyWEJWADYAOhCdCAN27LeCMLNOGS6AcANILQRO55Rq5m3rpdgQHBgLJnJve5Ku8XWd+eeo1ZXk8jY08YPfMWhLCI2CXWZewt3A1twMnhWLqX86XRUccKTeRm3nlqM3kyrghMvvbs2f2Ik/KnhAuwy6RG95Udy7nl81uRGXzQxCcfn9i/4o8LfWRv/LSXbaOJBYy/qDb71q7TEweo55LA28nN3HJqs/Gs2CszPWLaqTkE4A3AB9IA3NlWPPaPwncAvDNvGOsg5EcEhjoMSOZISAkpA2xV/BOWcYK39944vq/kLUSbJcJRRhsE2qHnC24WPHtTd4OxV+wR5BEYmnibkhOHAMDSDrsuwBpQORuOFgtf2S4QfE7q1COTb9KJhYlMAIdzbht7GcJXBSMqhNfO1AgUAAAG8wktpAGY+TWua5NoJzf35hPrtGfVK1J4/BSb+tCMHLxj4uDmMQAgo4xwAeMASnqKW8nI01DGsWXxq8tvHPguDAgixYANclKsJ/p3kEz6tR36m5BWBiWMAwBor9e59X3fkL7A1vjgkcZYBw2jHW7w+o+uZLnqAJE3iMmpLjXdcevQwc1j+CFb+AM50BYLnmVAcFZY9+1bP9K1ZargJmZjPJ1IdxrBYLAFEC0HyLAw0ggH5Do028KJWYw/AV92vS83cm1m5PkbEv7rcoW3fD+RN5+npnusWPfGgfcaBhBlibGWNFMOAAmQCygkxg9Y+d/sXFkwzUhM7dc08VN2vYJUYn5yYekVluNIbWf9XbAdktBzX0ikzuzAuOd16cmMmKxtrZJFUFr+Yp6JD6iZE5v634ucRpQlWkgjyhIA2JlShGwIACpv7Nmto8FHfwyQZSbff274o40XbPgFDZOnBSPkas4DgIqbu+3jrZXty6/r2c5c+FchgxsIACeSwypx8pb+d2u6UPemQgt94Xwho3wCABuTVMeONaWApmfOf2aKRDoUAEzkR8TxVkrm1Z7Y4AkUPWJcrCUXLI1dCwDHZ0IadawG22nv8vWn7iVP0YtknBjcvjv631t7ZG43zLseBAFMxMMhQQC7uk8BQGHN0APSDjxqdLKt70N6uKMDzpxN/rr3V1hi1QHWpo/Y9FPKXiGNvRHhZhvtcAACoiwHW+il/Ku7Two4k+cOhT+epf2LJ+tyMEBMemyT0IBxnH9R0RX9NwiRfQBwdkJa9ws2b8DovSSNzdqkjPI/xcIkznyQce2KrzXnK+8dPULZGSZ55u6+D0pems9y4cS7yPQLsADAees6Q8pb1UnC8prkmduFbVStcty+3s7gb+TMzG5JGT9iLfcbx97L7G8VjglYI523Icpy4Mi2IZOaelkwACzbuazi9Sy0wwCNacBoswQWAwdQBwEQK1G401K2F4nY4XNOayutXnuiQDr5B4kFG6JVbEa/F+ejb2vHSCFzzLmeXeNASwpolkDUFNUcWSXkmo+F5Qua1Mhfeg/nbU+zsKATflkibKGDnMJIX5Q8BXsEFHGyd8uZjtK9AIDy8iPLK6smH1q9uveaxet/9vCZ3UZlV565p7yWuTzCXLL23E4AqKtjteghFTlkAUBh9dGa4qsTY6vXM5dcNdq8wN+XjVik3y28FlKZnqTLqoaeCdUwV1zBXB5OB5H+JUoHCQCRCFsAUFx1uLq0Zrq3PMJcduXUZyUlb+ekceZ/UphmM1h6SAVTdJaJisqh3ZeHmddUM1dUjuxeufJ5/3ngtE7pZd31l1XFByprmEPV8cHS6qNV5xP9ysI052BN+eDvwyHmcCXzmvKpjorSk+vntELlZ39SGUrMhKuYK0Px/oqKT66atZ9n6avMWQuCAAFkQqX9P1Ny2ZNSeJSrkw7r6ecM5GplZW8RAjAm1gnujn7aU9s935b/9wDm7FkApEPFXRssKvqdIH8EBBABxmhoM/knjr/1WNfIlql05ufB/29Sh3RhFhU97QsXDu2sLpoZX1scOxku6r3rvNbia/4fnhUz5WbEpf0AAAAASUVORK5CYII=";  // 32x32 transparent PNG

        private static Icon _appIcon;

        /// <summary>
        /// Returns the Proofio icon for use as Form.Icon.
        /// Call once per dialog: <c>this.Icon = UITheme.AppIcon;</c>
        /// </summary>
        public static Icon AppIcon
        {
            get
            {
                if (_appIcon != null) return _appIcon;
                try
                {
                    var bytes = Convert.FromBase64String(FaviconBase64);
                    using (var ms = new MemoryStream(bytes))
                    using (var bmp = new Bitmap(ms))
                    {
                        _appIcon = Icon.FromHandle(bmp.GetHicon());
                    }
                }
                catch { /* fall back to default window icon */ }
                return _appIcon;
            }
        }

        /// <summary>
        /// Returns the favicon as a 32×32 Bitmap for use as a watermark.
        /// The caller is responsible for disposing the returned bitmap.
        /// </summary>
        public static Bitmap CreateWatermarkBitmap(int size = 80)
        {
            try
            {
                var bytes = Convert.FromBase64String(FaviconBase64);
                using (var ms = new MemoryStream(bytes))
                {
                    var bmp = new Bitmap(ms);
                    return new Bitmap(bmp, new Size(size, size));
                }
            }
            catch { return null; }
        }

        /// <summary>Loads a Bitmap from a base64 PNG string. Optionally resizes to <paramref name="size"/>×size.</summary>
        public static Bitmap LoadBitmap(string base64, int size = 0)
        {
            if (string.IsNullOrEmpty(base64)) return null;
            try
            {
                var bytes = Convert.FromBase64String(base64);
                using (var ms = new System.IO.MemoryStream(bytes))
                {
                    var bmp = new Bitmap(ms);
                    if (size > 0 && (bmp.Width != size || bmp.Height != size))
                    {
                        var resized = new Bitmap(bmp, new Size(size, size));
                        bmp.Dispose();
                        return resized;
                    }
                    // Return a copy so the MemoryStream can be closed
                    return new Bitmap(bmp);
                }
            }
            catch { return null; }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    /// <summary>Navy gradient title bar drawn at the top of every dialog.</summary>
    /// <summary>
    /// Navy gradient title bar with the Proofio icon (20x20) and white title text.
    /// </summary>
    internal sealed class ProofioTitleBar : Panel
    {
        private readonly Bitmap _icon;   // 20x20 transparent PNG
        private readonly Bitmap _wm;    // 64x64 watermark (right side)

        public ProofioTitleBar(string title)
        {
            Dock   = DockStyle.Top;
            Height = 46;

            _icon = UITheme.LoadBitmap(UITheme.TitleBarIconBase64);
            _wm   = UITheme.LoadBitmap(UITheme.FaviconBase64, 48);

            // White title label, positioned after icon
            var lbl = new Label
            {
                Text      = title,
                ForeColor = Color.White,
                Font      = UITheme.TitleFont,
                BackColor = Color.Transparent,
                AutoSize  = true
            };
            // Position will be set in Layout
            Layout += (s, e) =>
            {
                int iconW  = (_icon != null) ? _icon.Width + 10 : 0;
                lbl.Location = new Point(12 + iconW, (Height - lbl.PreferredHeight) / 2);
            };
            Controls.Add(lbl);

            Paint += OnPaint;
        }

        private void OnPaint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Navy gradient
            using (var br = new LinearGradientBrush(
                new Rectangle(0, 0, Width, Height),
                UITheme.Navy, UITheme.NavyLight,
                LinearGradientMode.Horizontal))
            {
                g.FillRectangle(br, 0, 0, Width, Height);
            }

            // Icon on the left (20x20, already transparent PNG)
            if (_icon != null)
            {
                int iy = (Height - _icon.Height) / 2;
                g.DrawImage(_icon, new Rectangle(10, iy, _icon.Width, _icon.Height));
            }

            // Subtle watermark on the right
            if (_wm != null)
            {
                var ia = new System.Drawing.Imaging.ImageAttributes();
                float[][] m = {
                    new float[]{1,0,0,0,0},
                    new float[]{0,1,0,0,0},
                    new float[]{0,0,1,0,0},
                    new float[]{0,0,0,0.10f,0},
                    new float[]{0,0,0,0,1}
                };
                ia.SetColorMatrix(new System.Drawing.Imaging.ColorMatrix(m));
                int s = _wm.Width;
                int x = Width - s - 8;
                int y = (Height - s) / 2;
                g.DrawImage(_wm, new Rectangle(x, y, s, s),
                    0, 0, s, s, GraphicsUnit.Pixel, ia);
            }

            // Bottom border line
            g.DrawLine(new Pen(Color.FromArgb(0, 60, 140)), 0, Height - 1, Width, Height - 1);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _icon?.Dispose();
                _wm?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    /// <summary>Primary / ghost button with rounded corners.</summary>
    internal sealed class StyledButton : Button
    {
        private readonly bool _primary;
        private bool _hovered;

        public StyledButton(string text, bool primary)
        {
            _primary = primary;
            Text     = text;
            Size     = new Size(95, 36);
            FlatStyle = FlatStyle.Flat;
            Cursor   = Cursors.Hand;
            Font     = UITheme.DefaultFont;
            ForeColor = primary ? Color.White : UITheme.TextSecondary;
            BackColor = primary ? UITheme.Accent : UITheme.Surface;
            FlatAppearance.BorderSize  = primary ? 0 : 1;
            FlatAppearance.BorderColor = UITheme.Border;
            UseVisualStyleBackColor = false;

            MouseEnter += (s, e) => { _hovered = true;  Invalidate(); };
            MouseLeave += (s, e) => { _hovered = false; Invalidate(); };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = new Rectangle(0, 0, Width - 1, Height - 1);

            Color fill = _primary
                ? (_hovered ? UITheme.AccentDark : UITheme.Accent)
                : (_hovered ? Color.FromArgb(232, 235, 245) : UITheme.Surface);

            using (var brush = new SolidBrush(fill))
            using (var path  = RoundedRect(rect, 6))
            {
                g.FillPath(brush, path);
                if (!_primary) g.DrawPath(new Pen(UITheme.Border), path);
            }

            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(Text, Font, new SolidBrush(ForeColor), new RectangleF(0, 0, Width, Height), sf);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var p = new GraphicsPath();
            p.AddArc(r.X, r.Y, radius * 2, radius * 2, 180, 90);
            p.AddArc(r.Right - radius * 2, r.Y, radius * 2, radius * 2, 270, 90);
            p.AddArc(r.Right - radius * 2, r.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            p.AddArc(r.X, r.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            p.CloseFigure();
            return p;
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    /// <summary>Styled single-line text box.</summary>
    internal sealed class StyledTextBox : TextBox
    {
        public StyledTextBox()
        {
            Height      = 36;
            BorderStyle = BorderStyle.FixedSingle;
            Font        = UITheme.DefaultFont;
            ForeColor   = UITheme.TextPrimary;
            BackColor   = Color.White;
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    /// <summary>Search box with magnifier icon (owner-drawn panel wrapper).</summary>
    internal sealed class SearchBox : Panel
    {
        private readonly TextBox _inner;
        public new event EventHandler TextChanged;
        public new string Text => _inner.Text;

        public SearchBox()
        {
            Height    = 40;
            BackColor = Color.White;

            _inner = new TextBox
            {
                BorderStyle = BorderStyle.None,
                Font        = UITheme.DefaultFont,
                ForeColor   = UITheme.TextPrimary,
                BackColor   = Color.White,
                Left        = 38,
                Top         = (40 - 20) / 2,
                Width       = Width - 50,
                Anchor      = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };
            _inner.TextChanged += (s, e) => TextChanged?.Invoke(this, e);

            Controls.Add(_inner);
            Paint  += OnPaint;
            Resize += (s, e) => _inner.Width = Width - 50;
        }

        private void OnPaint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.DrawRectangle(new Pen(UITheme.Border, 1.5f), 0, 0, Width - 1, Height - 1);

            int cx = 18, cy = Height / 2 - 1;
            g.DrawEllipse(new Pen(UITheme.TextMuted, 1.5f), cx - 6, cy - 6, 10, 10);
            g.DrawLine(new Pen(UITheme.TextMuted, 1.5f), cx + 2, cy + 2, cx + 7, cy + 7);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    /// <summary>
    /// List box that draws each item and paints a subtle watermark in the lower-right corner.
    /// </summary>
    internal sealed class StyledListBox : ListBox
    {
        private readonly Bitmap _watermark;

        public StyledListBox()
        {
            DrawMode      = DrawMode.OwnerDrawFixed;
            ItemHeight    = 42;
            BorderStyle   = BorderStyle.FixedSingle;
            Font          = UITheme.DefaultFont;
            BackColor     = Color.White;
            SelectionMode = SelectionMode.One;
            _watermark    = UITheme.CreateWatermarkBitmap(80);
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= Items.Count) return;

            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            bool selected = (e.State & DrawItemState.Selected) != 0;
            Color bg = selected
                ? UITheme.AccentLight
                : (e.Index % 2 == 0 ? Color.White : Color.FromArgb(249, 250, 252));

            g.FillRectangle(new SolidBrush(bg), e.Bounds);

            // Watermark on last visible item area (bottom-right of whole control)
            if (_watermark != null && e.Index == Items.Count - 1)
            {
                var ia = new System.Drawing.Imaging.ImageAttributes();
                float[][] m = {
                    new float[]{1,0,0,0,0},
                    new float[]{0,1,0,0,0},
                    new float[]{0,0,1,0,0},
                    new float[]{0,0,0,0.06f,0},
                    new float[]{0,0,0,0,1}
                };
                ia.SetColorMatrix(new System.Drawing.Imaging.ColorMatrix(m));
                int s = _watermark.Width;
                int wx = ClientSize.Width - s - 8;
                int wy = e.Bounds.Bottom - s - 4;
                if (wy >= 0)
                    g.DrawImage(_watermark, new Rectangle(wx, wy, s, s),
                        0, 0, s, s, GraphicsUnit.Pixel, ia);
            }

            if (selected)
                g.FillRectangle(new SolidBrush(UITheme.Accent),
                    e.Bounds.X, e.Bounds.Y, 3, e.Bounds.Height);

            g.DrawString(
                Items[e.Index]?.ToString() ?? string.Empty,
                UITheme.DefaultFont,
                new SolidBrush(selected ? UITheme.TextPrimary : UITheme.TextSecondary),
                new RectangleF(e.Bounds.X + 14, e.Bounds.Y + 2, e.Bounds.Width - 18, e.Bounds.Height),
                new StringFormat { LineAlignment = StringAlignment.Center });

            if (!selected)
                g.DrawLine(new Pen(UITheme.Border),
                    e.Bounds.X + 12, e.Bounds.Bottom - 1,
                    e.Bounds.Right - 12, e.Bounds.Bottom - 1);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _watermark?.Dispose();
            base.Dispose(disposing);
        }
    }
}
