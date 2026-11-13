using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace xls2sql
{
    /// <summary>
    /// Logical line index over a <see cref="SqlOutput"/> (global UTF-8 byte offsets).
    /// Built off the UI thread; ~16 bytes per line.
    /// </summary>
    internal sealed class OutputIndex
    {
        public static readonly OutputIndex Empty = new OutputIndex(SqlOutput.Empty);

        public SqlOutput Output { get; }
        public long[] ChunkStarts { get; }
        /// <summary>Start of each line.</summary>
        public long[] LineStarts { get; }
        /// <summary>End of each line, excluding the "\n" / "\r\n" terminator.</summary>
        public long[] LineEnds { get; }

        private OutputIndex(SqlOutput output)
        {
            Output = output;
            byte[][] chunks = output.Chunks;
            ChunkStarts = new long[chunks.Length + 1];
            for (int i = 0; i < chunks.Length; i++)
                ChunkStarts[i + 1] = ChunkStarts[i] + chunks[i].Length;

            var starts = new List<long> { 0 };
            var ends = new List<long>();
            byte prev = 0;
            for (int c = 0; c < chunks.Length; c++)
            {
                byte[] chunk = chunks[c];
                long baseOffset = ChunkStarts[c];
                int pos = 0;
                while (pos < chunk.Length)
                {
                    int nl = Array.IndexOf(chunk, (byte)'\n', pos);
                    if (nl < 0) break;
                    byte before = nl > 0 ? chunk[nl - 1] : prev;
                    long end = baseOffset + nl;
                    ends.Add(before == (byte)'\r' ? end - 1 : end);
                    starts.Add(end + 1);
                    pos = nl + 1;
                }
                if (chunk.Length > 0) prev = chunk[chunk.Length - 1];
            }
            ends.Add(output.Length);
            LineStarts = starts.ToArray();
            LineEnds = ends.ToArray();
        }

        public static OutputIndex Build(SqlOutput output) => new OutputIndex(output);

        public byte ByteAt(long offset)
        {
            int c = ChunkOf(offset);
            return Output.Chunks[c][offset - ChunkStarts[c]];
        }

        /// <summary>Copies bytes [start, end) into <paramref name="buffer"/> (grown if needed).</summary>
        public int CopyBytes(long start, long end, ref byte[] buffer)
        {
            int count = checked((int)(end - start));
            if (buffer.Length < count) buffer = new byte[Math.Max(count, buffer.Length * 2)];
            int written = 0;
            int c = count > 0 ? ChunkOf(start) : 0;
            while (written < count)
            {
                byte[] chunk = Output.Chunks[c];
                int offsetInChunk = (int)(start + written - ChunkStarts[c]);
                int n = Math.Min(chunk.Length - offsetInChunk, count - written);
                Buffer.BlockCopy(chunk, offsetInChunk, buffer, written, n);
                written += n;
                c++;
            }
            return count;
        }

        public string GetText(long start, long end)
        {
            byte[] buffer = new byte[0];
            int n = CopyBytes(start, end, ref buffer);
            return Encoding.UTF8.GetString(buffer, 0, n);
        }

        private int ChunkOf(long offset)
        {
            int i = Array.BinarySearch(ChunkStarts, offset);
            if (i < 0) i = ~i - 1;
            // Skip empty chunks that share a start offset.
            while (i < Output.Chunks.Length - 1 && ChunkStarts[i + 1] <= offset) i++;
            return i;
        }
    }

    /// <summary>
    /// Read-only, virtualized viewer for very large outputs. Lines are hard-wrapped into
    /// fixed-width monospace rows (column count in UTF-8 bytes, like xls2sql-rs) and only the
    /// visible rows are decoded and drawn, so memory and scroll cost don't depend on output size.
    /// Host inside a ScrollViewer with CanContentScroll="True". Supports mouse selection,
    /// Ctrl+C, Ctrl+A and a Copy / Select All context menu.
    /// </summary>
    internal sealed class OutputViewer : FrameworkElement, IScrollInfo
    {
        private const double Padding = 4;

        private readonly Typeface _typeface = new Typeface("Consolas");
        private readonly double _fontSize = 12;
        private readonly double _charWidth;
        private readonly double _rowHeight;

        private OutputIndex _index = OutputIndex.Empty;
        private int _cols;
        private long[] _rowStarts = { 0, 1 };   // first visual row of each line; last = total rows
        private double _firstRow;               // vertical offset, in rows
        private Size _viewport;
        private byte[] _rowBuffer = new byte[256];

        // Selection as global byte offsets; anchor = where the drag started.
        private long _selAnchor = -1, _selCaret = -1;

        public OutputViewer()
        {
            Focusable = true;
            FocusVisualStyle = null;
            Cursor = Cursors.IBeam;
            ClipToBounds = true;

            var probe = MakeText("MMMMMMMMMM", 1.0);
            _charWidth = probe.WidthIncludingTrailingWhitespace / 10;
            _rowHeight = Math.Ceiling(probe.Height);

            var menu = new ContextMenu();
            var copy = new MenuItem { Header = "Copy", InputGestureText = "Ctrl+C" };
            copy.Click += (s, e) => CopySelection();
            var selectAll = new MenuItem { Header = "Select All", InputGestureText = "Ctrl+A" };
            selectAll.Click += (s, e) => SelectAll();
            menu.Items.Add(copy);
            menu.Items.Add(selectAll);
            menu.Opened += (s, e) => copy.IsEnabled = HasSelection;
            ContextMenu = menu;
        }

        // Colors come from the theme (DynamicResource in MainWindow.xaml); defaults match the light theme.
        public static readonly DependencyProperty BackgroundProperty = DependencyProperty.Register(
            nameof(Background), typeof(Brush), typeof(OutputViewer),
            new FrameworkPropertyMetadata(Brushes.White, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty ForegroundProperty = DependencyProperty.Register(
            nameof(Foreground), typeof(Brush), typeof(OutputViewer),
            new FrameworkPropertyMetadata(Brushes.Black, FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty SelectionBrushProperty = DependencyProperty.Register(
            nameof(SelectionBrush), typeof(Brush), typeof(OutputViewer),
            new FrameworkPropertyMetadata(SystemColors.HighlightBrush, FrameworkPropertyMetadataOptions.AffectsRender));

        public Brush Background
        {
            get => (Brush)GetValue(BackgroundProperty);
            set => SetValue(BackgroundProperty, value);
        }

        public Brush Foreground
        {
            get => (Brush)GetValue(ForegroundProperty);
            set => SetValue(ForegroundProperty, value);
        }

        public Brush SelectionBrush
        {
            get => (Brush)GetValue(SelectionBrushProperty);
            set => SetValue(SelectionBrushProperty, value);
        }

        public bool HasSelection => _selAnchor >= 0 && _selAnchor != _selCaret;

        public void SetOutput(OutputIndex index)
        {
            _index = index ?? OutputIndex.Empty;
            _selAnchor = _selCaret = -1;
            _firstRow = 0;
            _cols = 0;
            RebuildRows();
            ScrollOwner?.InvalidateScrollInfo();
            InvalidateVisual();
        }

        public void Clear() => SetOutput(OutputIndex.Empty);

        public void SelectAll()
        {
            _selAnchor = 0;
            _selCaret = _index.Output.Length;
            InvalidateVisual();
        }

        public void CopySelection()
        {
            if (!HasSelection) return;
            long a = Math.Min(_selAnchor, _selCaret), b = Math.Max(_selAnchor, _selCaret);
            try
            {
                Clipboard.SetDataObject(_index.GetText(a, b), true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not copy to clipboard: {ex.Message}", "xls2sql", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ---- layout -------------------------------------------------------------------------

        private long TotalRows => _rowStarts[_rowStarts.Length - 1];

        private int ViewportRows => Math.Max(1, (int)(_viewport.Height / _rowHeight));

        private void RebuildRows()
        {
            int cols = Math.Max(10, (int)((_viewport.Width - 2 * Padding) / _charWidth));
            if (cols == _cols) return;

            // Keep the first visible byte in view across re-wraps.
            long topByte = _cols > 0 ? RowRange((long)_firstRow).Start : 0;

            _cols = cols;
            long[] starts = _index.LineStarts, ends = _index.LineEnds;
            var rowStarts = new long[starts.Length + 1];
            long total = 0;
            for (int i = 0; i < starts.Length; i++)
            {
                rowStarts[i] = total;
                long len = ends[i] - starts[i];
                total += Math.Max(1, (len + cols - 1) / cols);
            }
            rowStarts[starts.Length] = total;
            _rowStarts = rowStarts;
            _firstRow = ClampOffset(RowOfByte(topByte));
        }

        private int LineOfRow(long row)
        {
            int i = Array.BinarySearch(_rowStarts, row);
            if (i < 0) i = ~i - 1;
            return Math.Max(0, Math.Min(i, _index.LineStarts.Length - 1));
        }

        private (long Start, long End) RowRange(long row)
        {
            int line = LineOfRow(row);
            long lineStart = _index.LineStarts[line], lineEnd = _index.LineEnds[line];
            long k = row - _rowStarts[line];
            long a = FloorCharBoundary(Math.Min(lineStart + k * _cols, lineEnd), lineStart);
            long b = FloorCharBoundary(Math.Min(lineStart + (k + 1) * _cols, lineEnd), lineStart);
            return (a, b);
        }

        private long FloorCharBoundary(long i, long lineStart)
        {
            // UTF-8 continuation bytes are 10xxxxxx.
            while (i > lineStart && i < _index.Output.Length && (_index.ByteAt(i) & 0xC0) == 0x80) i--;
            return i;
        }

        private long RowOfByte(long offset)
        {
            int line = Array.BinarySearch(_index.LineStarts, offset);
            if (line < 0) line = ~line - 1;
            line = Math.Max(0, line);
            long rowsInLine = _rowStarts[line + 1] - _rowStarts[line];
            long k = (Math.Min(offset, _index.LineEnds[line]) - _index.LineStarts[line]) / _cols;
            return _rowStarts[line] + Math.Max(0, Math.Min(k, rowsInLine - 1));
        }

        private string RowText(long row, out long start)
        {
            var (a, b) = RowRange(row);
            start = a;
            int n = _index.CopyBytes(a, b, ref _rowBuffer);
            return Encoding.UTF8.GetString(_rowBuffer, 0, n);
        }

        protected override Size MeasureOverride(Size availableSize) =>
            new Size(double.IsInfinity(availableSize.Width) ? 0 : availableSize.Width,
                     double.IsInfinity(availableSize.Height) ? 0 : availableSize.Height);

        protected override Size ArrangeOverride(Size finalSize)
        {
            if (finalSize != _viewport)
            {
                _viewport = finalSize;
                RebuildRows();
                _firstRow = ClampOffset(_firstRow);
                ScrollOwner?.InvalidateScrollInfo();
            }
            return finalSize;
        }

        // ---- rendering ----------------------------------------------------------------------

        private FormattedText MakeText(string text, double pixelsPerDip) =>
            new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                              _typeface, _fontSize, Foreground, pixelsPerDip);

        protected override void OnRender(DrawingContext dc)
        {
            dc.DrawRectangle(Background, null, new Rect(RenderSize));
            if (_index.Output.Length == 0) return;

            double pixelsPerDip = VisualTreeHelper.GetDpi(this).PixelsPerDip;
            long first = (long)_firstRow;
            long last = Math.Min(TotalRows, first + ViewportRows + 1);
            long selA = Math.Min(_selAnchor, _selCaret), selB = Math.Max(_selAnchor, _selCaret);
            bool hasSel = HasSelection;

            for (long row = first; row < last; row++)
            {
                string text = RowText(row, out long rowStart);
                double y = (row - first) * _rowHeight;
                long rowEnd = rowStart + Encoding.UTF8.GetByteCount(text);

                if (hasSel && selA <= rowEnd && selB > rowStart)
                {
                    int cs = CharIndexOfByte(text, selA - rowStart);
                    int ce = CharIndexOfByte(text, selB - rowStart);
                    double x0 = Padding + cs * _charWidth;
                    // Selection running past the row end marks the line break / wrap with half a cell.
                    double x1 = Padding + ce * _charWidth + (selB > rowEnd ? _charWidth / 2 : 0);
                    if (x1 > x0)
                        dc.DrawRectangle(SelectionBrush, null, new Rect(x0, y, x1 - x0, _rowHeight));
                }

                if (text.Length > 0)
                    dc.DrawText(MakeText(Displayable(text), pixelsPerDip), new Point(Padding, y));
            }
        }

        /// <summary>Index of the first char at or after <paramref name="byteOffset"/> UTF-8 bytes into <paramref name="text"/>.</summary>
        private static int CharIndexOfByte(string text, long byteOffset)
        {
            long bytes = 0;
            int i = 0;
            while (i < text.Length && bytes < byteOffset)
            {
                char ch = text[i];
                if (char.IsHighSurrogate(ch) && i + 1 < text.Length)
                {
                    bytes += 4;
                    i += 2;
                }
                else
                {
                    bytes += ch < 0x80 ? 1 : ch < 0x800 ? 2 : 3;
                    i++;
                }
            }
            return i;
        }

        // Keep the monospace grid: tabs and other control chars render as one space.
        private static string Displayable(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] < ' ')
                {
                    var chars = text.ToCharArray();
                    for (int j = i; j < chars.Length; j++)
                        if (chars[j] < ' ') chars[j] = ' ';
                    return new string(chars);
                }
            }
            return text;
        }

        // ---- mouse / keyboard ---------------------------------------------------------------

        private long ByteAtPoint(Point p)
        {
            if (_index.Output.Length == 0) return 0;
            long row = (long)_firstRow + (long)Math.Floor(p.Y / _rowHeight);
            if (row < 0) return 0;
            if (row >= TotalRows) return _index.Output.Length;
            string text = RowText(row, out long rowStart);
            int col = (int)Math.Round((p.X - Padding) / _charWidth);
            col = Math.Max(0, Math.Min(col, text.Length));
            if (col > 0 && col < text.Length && char.IsLowSurrogate(text[col])) col--;
            return rowStart + Encoding.UTF8.GetByteCount(text.Substring(0, col));
        }

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            Focus();
            long pos = ByteAtPoint(e.GetPosition(this));
            if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0 && _selAnchor >= 0)
                _selCaret = pos;
            else
                _selAnchor = _selCaret = pos;
            CaptureMouse();
            InvalidateVisual();
            e.Handled = true;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (!IsMouseCaptured) return;
            Point p = e.GetPosition(this);
            // Auto-scroll while dragging above/below the viewport.
            if (p.Y < 0) SetVerticalOffset(_firstRow - 1);
            else if (p.Y > _viewport.Height) SetVerticalOffset(_firstRow + 1);
            _selCaret = ByteAtPoint(p);
            InvalidateVisual();
        }

        protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            if (IsMouseCaptured) ReleaseMouseCapture();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                if (e.Key == Key.C) { CopySelection(); e.Handled = true; return; }
                if (e.Key == Key.A) { SelectAll(); e.Handled = true; return; }
            }
            base.OnKeyDown(e);
        }

        // ---- IScrollInfo (vertical only; units are rows) ------------------------------------

        public bool CanVerticallyScroll { get; set; }
        public bool CanHorizontallyScroll { get; set; }
        public double ExtentWidth => _viewport.Width;
        public double ExtentHeight => TotalRows;
        public double ViewportWidth => _viewport.Width;
        public double ViewportHeight => ViewportRows;
        public double HorizontalOffset => 0;
        public double VerticalOffset => _firstRow;
        public ScrollViewer ScrollOwner { get; set; }

        private double ClampOffset(double offset) =>
            Math.Max(0, Math.Min(Math.Floor(offset), Math.Max(0, TotalRows - ViewportRows)));

        public void SetVerticalOffset(double offset)
        {
            offset = ClampOffset(offset);
            if (offset == _firstRow) return;
            _firstRow = offset;
            ScrollOwner?.InvalidateScrollInfo();
            InvalidateVisual();
        }

        public void LineUp() => SetVerticalOffset(_firstRow - 1);
        public void LineDown() => SetVerticalOffset(_firstRow + 1);
        public void PageUp() => SetVerticalOffset(_firstRow - ViewportRows);
        public void PageDown() => SetVerticalOffset(_firstRow + ViewportRows);
        public void MouseWheelUp() => SetVerticalOffset(_firstRow - SystemParameters.WheelScrollLines);
        public void MouseWheelDown() => SetVerticalOffset(_firstRow + SystemParameters.WheelScrollLines);
        public void LineLeft() { }
        public void LineRight() { }
        public void PageLeft() { }
        public void PageRight() { }
        public void MouseWheelLeft() { }
        public void MouseWheelRight() { }
        public void SetHorizontalOffset(double offset) { }
        public Rect MakeVisible(Visual visual, Rect rectangle) => rectangle;
    }
}
