/*
 * Form1.cs — Wintabber Dofus  [UI MODERNA]
 *
 * - Interfaz moderna con tabs smooth y colores refinados
 * - Toggles de Autofocus, Autotrade y Autogroup en el menú
 * - Sin menciones técnicas, interfaz amigable para el usuario
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DofusMiniTabber
{
    // ── Custom TabControl con tabs modernos y smooth ───────────────────────────
    public class SmoothTabControl : TabControl
    {
        private int _hoveredIndex = -1;

        public SmoothTabControl()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);
            DrawMode = TabDrawMode.OwnerDrawFixed;
            ItemSize = new Size(0, 36);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int prev = _hoveredIndex;
            _hoveredIndex = -1;
            for (int i = 0; i < TabCount; i++)
                if (GetTabRect(i).Contains(e.Location)) { _hoveredIndex = i; break; }
            if (_hoveredIndex != prev) Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _hoveredIndex = -1;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            g.Clear(Color.FromArgb(0x0D, 0x15, 0x1F));

            for (int i = 0; i < TabCount; i++)
                DrawTab(g, i);

            if (SelectedTab != null)
            {
                var cr = DisplayRectangle;
                using var borderPen = new Pen(Color.FromArgb(0x1E, 0x3A, 0x5F), 1);
                g.DrawRectangle(borderPen, cr.X, cr.Y, cr.Width - 1, cr.Height - 1);
            }
        }

        private void DrawTab(Graphics g, int index)
        {
            var rect  = GetTabRect(index);
            bool sel  = (index == SelectedIndex);
            bool hov  = (index == _hoveredIndex && !sel);
            string txt = TabPages[index].Text;

            Color bg, accent, fg;
            if (sel)
            {
                bg     = Color.FromArgb(0x12, 0x28, 0x4A);
                accent = Color.FromArgb(0x3D, 0x9B, 0xFF);
                fg     = Color.FromArgb(0xF0, 0xF8, 0xFF);
            }
            else if (hov)
            {
                bg     = Color.FromArgb(0x0F, 0x1E, 0x35);
                accent = Color.FromArgb(0x2A, 0x5C, 0x9A);
                fg     = Color.FromArgb(0xB0, 0xC8, 0xE8);
            }
            else
            {
                bg     = Color.FromArgb(0x0D, 0x15, 0x1F);
                accent = Color.FromArgb(0x1A, 0x2E, 0x45);
                fg     = Color.FromArgb(0x70, 0x90, 0xB8);
            }

            using (var bgBrush = new SolidBrush(bg))
                g.FillRectangle(bgBrush, rect);

            int accentH = sel ? 3 : (hov ? 2 : 1);
            using (var accentBrush = new SolidBrush(accent))
                g.FillRectangle(accentBrush, rect.X, rect.Y, rect.Width, accentH);

            if (index < TabCount - 1)
            {
                using var sep = new Pen(Color.FromArgb(0x18, 0x28, 0x40), 1);
                g.DrawLine(sep, rect.Right - 1, rect.Top + 6, rect.Right - 1, rect.Bottom - 6);
            }

            var sf = new StringFormat
            {
                Alignment     = StringAlignment.Center,
                LineAlignment = StringAlignment.Center,
                Trimming      = StringTrimming.EllipsisCharacter,
                FormatFlags   = StringFormatFlags.NoWrap
            };

            using var fgBrush = new SolidBrush(fg);
            var font = sel
                ? new Font("Segoe UI Semibold", 8.5f, FontStyle.Bold)
                : new Font("Segoe UI",           8.5f, FontStyle.Regular);

            var textRect = new RectangleF(rect.X + 4, rect.Y + accentH, rect.Width - 8, rect.Height - accentH);
            g.DrawString(txt, font, fgBrush, textRect, sf);
            font.Dispose();
        }
    }

    // ── Renderer moderno para ToolStrip ──────────────────────────────────────
    public class ModernToolStripRenderer : ToolStripProfessionalRenderer
    {
        public ModernToolStripRenderer() : base(new ModernColorTable()) { }

        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
        {
            e.Graphics.Clear(Color.FromArgb(0x0A, 0x12, 0x1C));
        }

        protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
        {
            if (e.Item is ToolStripButton btn)
            {
                var g    = e.Graphics;
                var rect = new Rectangle(2, 2, e.Item.Width - 4, e.Item.Height - 4);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                Color bg;
                if (btn.Pressed)
                    bg = Color.FromArgb(0x0C, 0x3A, 0x70);
                else if (btn.Selected)
                    bg = Color.FromArgb(0x12, 0x2E, 0x55);
                else if (btn.Tag is "toggle-on")
                    bg = Color.FromArgb(0x0A, 0x2E, 0x18);
                else if (btn.Tag is "toggle-off")
                    bg = Color.FromArgb(0x2E, 0x0A, 0x0A);
                else
                    bg = Color.FromArgb(0x10, 0x1E, 0x32);

                using var path = RoundedRect(rect, 5);
                using (var brush = new SolidBrush(bg))
                    g.FillPath(brush, path);

                Color border = btn.Selected
                    ? Color.FromArgb(0x2A, 0x70, 0xCC)
                    : Color.FromArgb(0x1C, 0x34, 0x55);
                using var pen = new Pen(border, 1f);
                g.DrawPath(pen, path);
            }
            else
            {
                base.OnRenderButtonBackground(e);
            }
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            var g = e.Graphics;
            int cx = e.Item.Width / 2;
            using var pen = new Pen(Color.FromArgb(0x1C, 0x34, 0x55), 1);
            g.DrawLine(pen, cx, 6, cx, e.Item.Height - 6);
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            e.TextColor = e.Item.Enabled
                ? Color.FromArgb(0xCC, 0xDD, 0xEE)
                : Color.FromArgb(0x44, 0x66, 0x88);
            base.OnRenderItemText(e);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(r.X, r.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(r.Right - radius * 2, r.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(r.Right - radius * 2, r.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(r.X, r.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    public class ModernColorTable : ProfessionalColorTable
    {
        public override Color ToolStripGradientBegin   => Color.FromArgb(0x0A, 0x12, 0x1C);
        public override Color ToolStripGradientMiddle  => Color.FromArgb(0x0A, 0x12, 0x1C);
        public override Color ToolStripGradientEnd     => Color.FromArgb(0x0A, 0x12, 0x1C);
        public override Color ToolStripBorder          => Color.FromArgb(0x1C, 0x34, 0x55);
        public override Color SeparatorDark            => Color.FromArgb(0x1C, 0x34, 0x55);
        public override Color SeparatorLight           => Color.FromArgb(0x1C, 0x34, 0x55);
        public override Color MenuItemSelected         => Color.FromArgb(0x12, 0x2E, 0x55);
        public override Color MenuItemBorder           => Color.FromArgb(0x2A, 0x70, 0xCC);
        public override Color MenuBorder               => Color.FromArgb(0x1C, 0x34, 0x55);
        public override Color MenuItemPressedGradientBegin => Color.FromArgb(0x0C, 0x3A, 0x70);
        public override Color MenuItemPressedGradientEnd   => Color.FromArgb(0x0C, 0x3A, 0x70);
        public override Color ImageMarginGradientBegin => Color.FromArgb(0x0A, 0x12, 0x1C);
        public override Color ImageMarginGradientMiddle=> Color.FromArgb(0x0A, 0x12, 0x1C);
        public override Color ImageMarginGradientEnd   => Color.FromArgb(0x0A, 0x12, 0x1C);
    }

    // ════════════════════════════════════════════════════════════════════════
    //  Form1 — Ventana principal
    // ════════════════════════════════════════════════════════════════════════
    public partial class Form1 : Form
    {
        // ── UI ────────────────────────────────────────────────────────────────
        private readonly ToolStrip       _toolbar              = new();
        private readonly ToolStrip       _floatingToolbar      = new();
        private readonly ToolStripButton _captureButton        = new();
        private readonly ToolStripButton _savePositionButton   = new();
        private readonly ToolStripButton _restorePositionButton= new();
        private readonly ToolStripButton _loadPreferredButton  = new();
        private readonly ToolStripButton _manageLayoutsButton  = new();
        private readonly ToolStripButton _hideMenuButton       = new();

        // ── Toggles de funciones ───────────────────────────────────────────
        private readonly ToolStripButton _toggleAutofocusBtn   = new();
        private readonly ToolStripButton _toggleAutotradeBtn   = new();
        private readonly ToolStripButton _toggleAutogroupBtn   = new();

        private readonly ToolStripLabel  _hotkeysLabel         = new();
        private readonly SmoothTabControl _tabs                = new();
        private readonly ContextMenuStrip _tabMenu             = new();
        private readonly System.Windows.Forms.Timer _resizeDebounceTimer = new();
        private readonly System.Windows.Forms.Timer _updateTitleTimer    = new();
        private readonly System.Windows.Forms.Timer _statsTimer          = new();
        private readonly Dictionary<IntPtr, EmbeddedWindowInfo> _embeddedByHwnd = new();
        private readonly NotifyIcon      _trayIcon             = new();
        private readonly ContextMenuStrip _trayMenu            = new();
        private bool    _menuVisible         = true;
        private bool    _isCapturing;
        private TabPage? _previousTab;
        private TabPage? _draggedTab;
        private string? _preferredLayoutName = null;

        // Estados de toggles
        private bool _autofocusEnabled = true;
        private bool _autotradeEnabled = true;
        private bool _autogroupEnabled = true;

        // ── Monitor de red (en-proceso, sin IPC) ─────────────────────────────
        private DofusNetMonitor? _monitor;

        // ── Hotkeys ───────────────────────────────────────────────────────────
        private const int WM_HOTKEY              = 0x0312;
        private const int HOTKEY_ID_PREV         = 1;
        private const int HOTKEY_ID_NEXT         = 2;
        private const int HOTKEY_ID_TOGGLE_MENU  = 3;
        private const int HOTKEY_ID_SAVE_POS     = 4;
        private const int HOTKEY_ID_RESTORE_POS  = 5;
        private const int HOTKEY_ID_MANAGE       = 6;
        private const int HOTKEY_ID_NUM_START    = 10;

        private const uint MOD_ALT      = 0x0001;
        private const uint MOD_CONTROL  = 0x0002;
        private const uint MOD_NOREPEAT = 0x4000;

        // ── Estilos de ventana ────────────────────────────────────────────────
        private const int GWL_STYLE      = -16;
        private const int WS_CAPTION     = 0x00C00000;
        private const int WS_THICKFRAME  = 0x00040000;
        private const int WS_BORDER      = 0x00800000;
        private const int WS_DLGFRAME    = 0x00400000;
        private const int SW_SHOW        = 5;
        private const int SWP_NOZORDER   = 0x0004;
        private const int SWP_NOMOVE     = 0x0002;
        private const int SWP_NOSIZE     = 0x0001;
        private const int SWP_FRAMECHANGED = 0x0020;
        private const int SWP_NOREDRAW   = 0x0008;

        private const uint RDW_FRAME      = 0x0400;
        private const uint RDW_INVALIDATE = 0x0001;
        private const uint RDW_UPDATENOW  = 0x0100;
        private const uint RDW_ALLCHILDREN= 0x0080;

        private const uint   PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
        private const string TARGET_PROCESS_NAME               = "Dofus Retro.exe";

        // ── Win32 ─────────────────────────────────────────────────────────────
        [StructLayout(LayoutKind.Sequential)] public struct RECT  { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc fn, IntPtr lp);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] private static extern IntPtr SetParent(IntPtr child, IntPtr newParent);
        [DllImport("user32.dll")] private static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool r);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr h, int cmd);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetWindowLong(IntPtr h, int idx);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongW")] private static extern int SetWindowLong(IntPtr h, int idx, int val);
        [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr h, int id, uint mod, uint vk);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr h, int id);
        [DllImport("user32.dll")] private static extern bool RedrawWindow(IntPtr h, IntPtr rect, IntPtr rgn, uint flags);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("kernel32.dll")] private static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr h);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern bool QueryFullProcessImageName(IntPtr h, uint flags, StringBuilder name, ref uint size);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr h);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr h);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }
        private const uint FLASHW_ALL = 3;
        private const uint FLASHW_TIMERNOFG = 12;
        private static readonly IntPtr HWND_TOPMOST   = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

        private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lp);

        // ═════════════════════════════════════════════════════════════════════
        public Form1()
        {
            ConfigureForm();
            BuildUi();
            SetupTrayIcon();
            RegisterBaseHotkeys();
            RegisterNumberHotkeys();

            _updateTitleTimer.Interval = 300;
            _updateTitleTimer.Tick += (_, _) => UpdateDynamicTitles();
            _updateTitleTimer.Start();

            _statsTimer.Interval = 15_000;
            _statsTimer.Tick += (_, _) => { };
            _statsTimer.Start();

            TryStartMonitor();
        }

        // ═════════════════════════════════════════════════════════════════════
        private void TryStartMonitor()
        {
            try
            {
                _monitor = new DofusNetMonitor(name =>
                {
                    if (IsHandleCreated)
                        BeginInvoke(() => ActivateTabByCharacterName(name));
                });

                _monitor.FeatureAutofocus = _autofocusEnabled;
                _monitor.FeatureAutotrade = _autotradeEnabled;
                _monitor.FeatureAutogroup = _autogroupEnabled;

                _monitor.Start();
            }
            catch { }
        }

        // ═════════════════════════════════════════════════════════════════════
        private void ConfigureForm()
        {
            Text        = "Wintabber Dofus";
            BackColor   = Color.FromArgb(0x0D, 0x15, 0x1F);
            WindowState = FormWindowState.Maximized;
            KeyPreview  = true;
            FormClosing += OnFormClosingRestoreWindows;
            TopMost     = false;
        }

        private void SetupTrayIcon()
        {
            _trayIcon.Icon    = Icon ?? CreateFallbackIcon();
            _trayIcon.Text    = "Wintabber Dofus";
            _trayIcon.Visible = true;

            var restoreItem = new ToolStripMenuItem("Restaurar",         null, (_, _) => RestoreWindow());
            var captureItem = new ToolStripMenuItem("Capturar ventanas", null, (_, _) => CaptureWindows());
            var sep         = new ToolStripSeparator();
            var exitItem    = new ToolStripMenuItem("Salir",             null, (_, _) => ExitApplication());

            _trayMenu.Items.AddRange(new ToolStripItem[] { restoreItem, captureItem, sep, exitItem });
            _trayMenu.BackColor  = Color.FromArgb(0x0A, 0x12, 0x1C);
            _trayMenu.ForeColor  = Color.FromArgb(0xCC, 0xDD, 0xEE);
            _trayMenu.RenderMode = ToolStripRenderMode.System;

            _trayIcon.ContextMenuStrip = _trayMenu;
            _trayIcon.DoubleClick += (_, _) => RestoreWindow();
        }

        private static Icon CreateFallbackIcon()
        {
            using var bmp  = new Bitmap(16, 16);
            using var g    = Graphics.FromImage(bmp);
            g.Clear(Color.FromArgb(0x0A, 0x12, 0x1C));
            using var font = new Font("Segoe UI", 8f, FontStyle.Bold);
            g.DrawString("W", font, Brushes.White, 1f, 1f);
            return Icon.FromHandle(bmp.GetHicon());
        }

        private void RestoreWindow()
        {
            if (WindowState == FormWindowState.Minimized)
                WindowState = FormWindowState.Maximized;
            Activate();
        }

        private void ExitApplication() => Close();

        // ═════════════════════════════════════════════════════════════════════
        private void BuildUi()
        {
            SuspendLayout();

            _toolbar.Dock    = DockStyle.Top;
            _toolbar.Height  = 1;
            _toolbar.Visible = false;

            _floatingToolbar.Dock        = DockStyle.Top;
            _floatingToolbar.Height      = 46;
            _floatingToolbar.GripStyle   = ToolStripGripStyle.Hidden;
            _floatingToolbar.CanOverflow = true;
            _floatingToolbar.Stretch     = true;
            _floatingToolbar.Padding     = new Padding(6, 0, 6, 0);
            _floatingToolbar.Renderer    = new ModernToolStripRenderer();

            StyleButton(_captureButton,         "⚡  Capturar Tabs",  "Capturar todas las ventanas de Dofus abiertas");
            StyleButton(_savePositionButton,    "💾  Guardar",   "Guardar el orden actual de las pestañas como layout");
            StyleButton(_restorePositionButton, "🔁  Cargar",    "Cargar un layout guardado");
            StyleButton(_loadPreferredButton,   "⭐  Loadout favorito", "Cargar el layout preferido");
            StyleButton(_manageLayoutsButton,   "📋  Loadouts",   "Gestionar layouts guardados");
            StyleButton(_hideMenuButton,        "👁  Menú",      "Mostrar u ocultar esta barra");

            _captureButton.Click          += (_, _) => CaptureWindows();
            _savePositionButton.Click     += (_, _) => SaveCurrentPositions();
            _restorePositionButton.Click  += (_, _) => QuickRestoreLayout();
            _loadPreferredButton.Click    += (_, _) => LoadPreferredLayout();
            _manageLayoutsButton.Click    += (_, _) => OpenLayoutManager();
            _hideMenuButton.Click         += (_, _) => ToggleFloatingMenu();

            SetupToggleButton(_toggleAutofocusBtn, "🎯  Autofocus Turno",
                "Cambiar automáticamente al personaje cuyo turno llegó en el GTS", _autofocusEnabled);
            SetupToggleButton(_toggleAutotradeBtn, "🛒  Autotrade",
                "Cambiar automáticamente al personaje que recibe un intercambio", _autotradeEnabled);
            SetupToggleButton(_toggleAutogroupBtn, "👥  Autogrupo",
                "Cambiar automáticamente al personaje que recibe una invitación de grupo", _autogroupEnabled);

            _toggleAutofocusBtn.Click += (_, _) => ToggleFeature(
                ref _autofocusEnabled, _toggleAutofocusBtn,
                v => { if (_monitor != null) _monitor.FeatureAutofocus = v; });
            _toggleAutotradeBtn.Click += (_, _) => ToggleFeature(
                ref _autotradeEnabled, _toggleAutotradeBtn,
                v => { if (_monitor != null) _monitor.FeatureAutotrade = v; });
            _toggleAutogroupBtn.Click += (_, _) => ToggleFeature(
                ref _autogroupEnabled, _toggleAutogroupBtn,
                v => { if (_monitor != null) _monitor.FeatureAutogroup = v; });

            _hotkeysLabel.Text      = "F1/F2 Ant·Sig   F3 Menú   F4 Guardar   F5 Cargar   F6 Layouts   Ctrl+Alt+1..9";
            _hotkeysLabel.ForeColor = Color.FromArgb(0x33, 0x55, 0x77);
            _hotkeysLabel.Alignment = ToolStripItemAlignment.Right;
            _hotkeysLabel.Font      = new Font("Segoe UI", 7.5f);

            _floatingToolbar.Items.Add(_captureButton);
            _floatingToolbar.Items.Add(new ToolStripSeparator());
            _floatingToolbar.Items.Add(_savePositionButton);
            _floatingToolbar.Items.Add(_restorePositionButton);
            _floatingToolbar.Items.Add(_loadPreferredButton);
            _floatingToolbar.Items.Add(_manageLayoutsButton);
            _floatingToolbar.Items.Add(new ToolStripSeparator());
            _floatingToolbar.Items.Add(_toggleAutofocusBtn);
            _floatingToolbar.Items.Add(_toggleAutotradeBtn);
            _floatingToolbar.Items.Add(_toggleAutogroupBtn);
            _floatingToolbar.Items.Add(new ToolStripSeparator());
            _floatingToolbar.Items.Add(_hideMenuButton);
            _floatingToolbar.Items.Add(_hotkeysLabel);

            _tabs.Dock                 = DockStyle.Fill;
            _tabs.SelectedIndexChanged += (_, _) => OnTabChanged();
            _tabs.MouseDown            += Tabs_MouseDown;
            _tabs.MouseMove            += Tabs_MouseMove;
            _tabs.MouseUp              += Tabs_MouseUp;

            var liberarItem = new ToolStripMenuItem("Liberar", null, (_, _) => ReleaseCurrentTab());
            var cerrarItem  = new ToolStripMenuItem("Cerrar",  null, (_, _) => CloseCurrentTab());
            _tabMenu.Items.Add(liberarItem);
            _tabMenu.Items.Add(cerrarItem);
            _tabMenu.BackColor  = Color.FromArgb(0x0A, 0x12, 0x1C);
            _tabMenu.ForeColor  = Color.FromArgb(0xCC, 0xDD, 0xEE);
            _tabMenu.RenderMode = ToolStripRenderMode.System;

            Controls.Add(_tabs);
            Controls.Add(_floatingToolbar);
            Controls.Add(_toolbar);

            _resizeDebounceTimer.Interval = 60;
            _resizeDebounceTimer.Tick += (_, _) =>
            {
                _resizeDebounceTimer.Stop();
                ResizeActiveEmbeddedWindow();
            };

            UpdatePreferredButton();
            ResumeLayout(true);
        }

        private static void StyleButton(ToolStripButton btn, string text, string tooltip)
        {
            btn.Text         = text;
            btn.ToolTipText  = tooltip;
            btn.DisplayStyle = ToolStripItemDisplayStyle.Text;
            btn.ForeColor    = Color.FromArgb(0xCC, 0xDD, 0xEE);
            btn.Font         = new Font("Segoe UI", 8.5f);
            btn.AutoSize     = true;
            btn.Margin       = new Padding(2, 4, 2, 4);
            btn.Padding      = new Padding(8, 0, 8, 0);
        }

        private static void SetupToggleButton(ToolStripButton btn, string text, string tooltip, bool enabled)
        {
            StyleButton(btn, text, tooltip);
            UpdateToggleVisual(btn, enabled);
        }

        private static void UpdateToggleVisual(ToolStripButton btn, bool enabled)
        {
            btn.Tag       = enabled ? "toggle-on" : "toggle-off";
            btn.ForeColor = enabled
                ? Color.FromArgb(0x7D, 0xEE, 0xA8)
                : Color.FromArgb(0xEE, 0x7D, 0x7D);
        }

        private void ToggleFeature(ref bool state, ToolStripButton btn, Action<bool> apply)
        {
            state = !state;
            apply(state);
            UpdateToggleVisual(btn, state);
            _floatingToolbar.Invalidate();
        }

        private void ToggleFloatingMenu()
        {
            _menuVisible             = !_menuVisible;
            _floatingToolbar.Visible = _menuVisible;
            ScheduleResizeActiveTab();
        }

        private void Tabs_MouseDown(object? sender, MouseEventArgs e)
        {
            for (int i = 0; i < _tabs.TabCount; i++)
                if (_tabs.GetTabRect(i).Contains(e.Location))
                    { _draggedTab = _tabs.TabPages[i]; break; }
        }

        private void Tabs_MouseMove(object? sender, MouseEventArgs e)
        {
            if (_draggedTab == null || e.Button != MouseButtons.Left) return;
            for (int i = 0; i < _tabs.TabCount; i++)
            {
                if (_tabs.GetTabRect(i).Contains(e.Location))
                {
                    var targetTab = _tabs.TabPages[i];
                    if (targetTab == _draggedTab) return;
                    _tabs.TabPages.Remove(_draggedTab);
                    _tabs.TabPages.Insert(i, _draggedTab);
                    _tabs.SelectedTab = _draggedTab;
                    break;
                }
            }
        }

        private void Tabs_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                for (int i = 0; i < _tabs.TabCount; i++)
                    if (_tabs.GetTabRect(i).Contains(e.Location))
                        { _tabs.SelectedIndex = i; _tabMenu.Show(_tabs, e.Location); break; }
            }
            _draggedTab = null;
        }

        private void ReleaseCurrentTab()
        {
            if (_tabs.SelectedTab == null) return;
            var tab  = _tabs.SelectedTab;
            var item = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == tab);
            if (item == null) return;

            SetParent(item.Hwnd, IntPtr.Zero);
            SetWindowLong(item.Hwnd, GWL_STYLE, item.OriginalStyle);
            SetWindowPos(item.Hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
            ShowWindow(item.Hwnd, SW_SHOW);

            _embeddedByHwnd.Remove(item.Hwnd);
            _tabs.TabPages.Remove(tab);
            ReorderTabs();
        }

        private void CloseCurrentTab()
        {
            if (_tabs.SelectedTab == null) return;
            var tab  = _tabs.SelectedTab;
            var item = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == tab);
            if (item != null)
            {
                SetParent(item.Hwnd, IntPtr.Zero);
                SetWindowLong(item.Hwnd, GWL_STYLE, item.OriginalStyle);
                SetWindowPos(item.Hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
                ShowWindow(item.Hwnd, SW_SHOW);
                _embeddedByHwnd.Remove(item.Hwnd);
            }
            _tabs.TabPages.Remove(tab);
            ReorderTabs();
        }

        private void ReorderTabs()
        {
            for (int i = 0; i < _tabs.TabCount; i++)
            {
                var page = _tabs.TabPages[i];
                var info = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == page);
                string title = info != null
                    ? (GetWindowTitle(info.Hwnd).Trim() is { Length: > 0 } t ? t : TARGET_PROCESS_NAME)
                    : page.Text;
                page.Text = $"{i + 1}. {title}";
            }
        }

        private void RegisterBaseHotkeys()
        {
            RegisterHotKey(Handle, HOTKEY_ID_PREV,        MOD_NOREPEAT, (uint)Keys.F1);
            RegisterHotKey(Handle, HOTKEY_ID_NEXT,        MOD_NOREPEAT, (uint)Keys.F2);
            RegisterHotKey(Handle, HOTKEY_ID_TOGGLE_MENU, MOD_NOREPEAT, (uint)Keys.F3);
            RegisterHotKey(Handle, HOTKEY_ID_SAVE_POS,    MOD_NOREPEAT, (uint)Keys.F4);
            RegisterHotKey(Handle, HOTKEY_ID_RESTORE_POS, MOD_NOREPEAT, (uint)Keys.F5);
            RegisterHotKey(Handle, HOTKEY_ID_MANAGE,      MOD_NOREPEAT, (uint)Keys.F6);
        }

        private void RegisterNumberHotkeys()
        {
            for (int i = 1; i <= 9; i++)
                RegisterHotKey(Handle, HOTKEY_ID_NUM_START + i - 1,
                               MOD_CONTROL | MOD_ALT, (uint)(Keys.D0 + i));
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg != 0 && (uint)m.Msg == Program.WM_BRING_TO_FRONT)
            {
                RestoreWindow();
                return;
            }

            if (m.Msg == WM_HOTKEY)
            {
                var id = m.WParam.ToInt32();
                if      (id == HOTKEY_ID_PREV)        PrevTab();
                else if (id == HOTKEY_ID_NEXT)        NextTab();
                else if (id == HOTKEY_ID_TOGGLE_MENU) ToggleFloatingMenu();
                else if (id == HOTKEY_ID_SAVE_POS)    SaveCurrentPositions();
                else if (id == HOTKEY_ID_RESTORE_POS) QuickRestoreLayout();
                else if (id == HOTKEY_ID_MANAGE)      OpenLayoutManager();
                else if (id >= HOTKEY_ID_NUM_START && id <= HOTKEY_ID_NUM_START + 8)
                    JumpToTab(id - HOTKEY_ID_NUM_START);
            }

            base.WndProc(ref m);
        }

        private void JumpToTab(int index)
        {
            if (index >= 0 && index < _tabs.TabCount)
                _tabs.SelectedIndex = index;
        }

        private void NextTab()
        {
            if (_tabs.TabCount == 0) return;
            _tabs.SelectedIndex = (_tabs.SelectedIndex + 1) % _tabs.TabCount;
        }

        private void PrevTab()
        {
            if (_tabs.TabCount == 0) return;
            _tabs.SelectedIndex = (_tabs.SelectedIndex - 1 + _tabs.TabCount) % _tabs.TabCount;
        }

        private void CaptureWindows()
        {
            if (_isCapturing) return;
            _isCapturing = true;
            try
            {
                EnumWindows((hwnd, _) =>
                {
                    if (!IsWindowVisible(hwnd)) return true;
                    if (_embeddedByHwnd.ContainsKey(hwnd)) return true;
                    if (!IsDofusRetroProcess(hwnd)) return true;
                    var title = GetWindowTitle(hwnd);
                    EmbedWindow(hwnd, string.IsNullOrWhiteSpace(title) ? TARGET_PROCESS_NAME : title);
                    return true;
                }, IntPtr.Zero);
            }
            finally { _isCapturing = false; }
        }

        private static bool IsDofusRetroProcess(IntPtr hwnd)
        {
            GetWindowThreadProcessId(hwnd, out uint pid);
            if (pid == 0) return false;
            var hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
            if (hProcess == IntPtr.Zero) return false;
            try
            {
                var sb    = new StringBuilder(1024);
                uint size = (uint)sb.Capacity;
                if (!QueryFullProcessImageName(hProcess, 0, sb, ref size)) return false;
                return System.IO.Path.GetFileName(sb.ToString())
                    .Equals(TARGET_PROCESS_NAME, StringComparison.OrdinalIgnoreCase);
            }
            finally { CloseHandle(hProcess); }
        }

        private void EmbedWindow(IntPtr hwnd, string title)
        {
            var panel     = new Panel { Dock = DockStyle.Fill, BackColor = Color.Black };
            string cleanT = string.IsNullOrWhiteSpace(title) ? TARGET_PROCESS_NAME : title.Trim();
            var tab       = new TabPage($"{_tabs.TabCount + 1}. {cleanT}");
            tab.Controls.Add(panel);
            _tabs.TabPages.Add(tab);

            panel.CreateControl();

            var originalStyle = GetWindowLong(hwnd, GWL_STYLE);
            var stripped      = originalStyle & ~WS_CAPTION & ~WS_THICKFRAME & ~WS_BORDER & ~WS_DLGFRAME;
            SetWindowLong(hwnd, GWL_STYLE, stripped);
            SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
            SetParent(hwnd, panel.Handle);
            ShowWindow(hwnd, SW_SHOW);

            _embeddedByHwnd[hwnd] = new EmbeddedWindowInfo(hwnd, panel, tab, originalStyle);

            panel.Resize += (_, _) =>
            {
                var info = _embeddedByHwnd.Values.FirstOrDefault(v => v.HostPanel == panel);
                if (info != null) ResizeWindowIfNeeded(info);
            };

            BeginInvoke(() =>
            {
                if (_embeddedByHwnd.TryGetValue(hwnd, out var info))
                {
                    info.LastKnownSize = Size.Empty;
                    ResizeWindowIfNeeded(info);
                }
            });

            ScheduleResizeActiveTab();
        }

        private void ResizeWindowIfNeeded(EmbeddedWindowInfo info)
        {
            var size = info.HostPanel.ClientSize;
            if (size.Width <= 0 || size.Height <= 0) return;
            if (size == info.LastKnownSize) return;

            MoveWindow(info.Hwnd, 0, 0, size.Width, size.Height, true);
            SetWindowPos(info.Hwnd, IntPtr.Zero, 0, 0, size.Width, size.Height, SWP_NOZORDER | SWP_FRAMECHANGED);
            RedrawWindow(info.Hwnd, IntPtr.Zero, IntPtr.Zero, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);

            info.LastKnownSize = size;
        }

        private void OnTabChanged()
        {
            var currentTab = _tabs.SelectedTab;
            if (currentTab == null) return;

            if (_previousTab != null && _previousTab != currentTab)
            {
                var prev = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == _previousTab);
                if (prev != null)
                    SetWindowPos(prev.Hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOREDRAW | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER);
            }

            var active = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == currentTab);
            if (active != null)
            {
                active.LastKnownSize = Size.Empty;
                ResizeWindowIfNeeded(active);
            }

            _previousTab = currentTab;
        }

        private void ScheduleResizeActiveTab()
        {
            _resizeDebounceTimer.Stop();
            _resizeDebounceTimer.Start();
        }

        private void ResizeActiveEmbeddedWindow()
        {
            if (_tabs.SelectedTab is null) return;
            var active = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == _tabs.SelectedTab);
            if (active is null) return;
            active.LastKnownSize = Size.Empty;
            ResizeWindowIfNeeded(active);
        }

        private void UpdateDynamicTitles()
        {
            foreach (var info in _embeddedByHwnd.Values)
            {
                string currentTitle = GetWindowTitle(info.Hwnd).Trim();
                if (string.IsNullOrWhiteSpace(currentTitle)) currentTitle = TARGET_PROCESS_NAME;
                int    index    = _tabs.TabPages.IndexOf(info.TabPage) + 1;
                string expected = $"{index}. {currentTitle}";
                if (info.TabPage.Text != expected) info.TabPage.Text = expected;
            }
        }

        private void SaveCurrentPositions()
        {
            try
            {
                using var dlg = new Form
                {
                    Text            = "Guardar Layout",
                    Size            = new Size(420, 250),
                    StartPosition   = FormStartPosition.CenterParent,
                    BackColor       = Color.FromArgb(0x0A, 0x12, 0x1C),
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox     = false,
                    MinimizeBox     = false,
                    Font            = new Font("Segoe UI", 9F)
                };

                var titleLbl  = new Label  { Text = "Guardar configuración actual", ForeColor = Color.FromArgb(0xCC, 0xDD, 0xEE), Location = new Point(20, 15), Size = new Size(350, 30), Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold) };
                var nameLbl   = new Label  { Text = "Nombre del layout:", ForeColor = Color.FromArgb(0x88, 0xAA, 0xCC), Location = new Point(20, 60), Size = new Size(150, 20) };
                var nameTb    = new TextBox{ Location = new Point(20, 82), Size = new Size(360, 30), BackColor = Color.FromArgb(0x0F, 0x1E, 0x32), ForeColor = Color.FromArgb(0xCC, 0xDD, 0xEE), BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) };
                var descLbl   = new Label  { Text = "Descripción (opcional):", ForeColor = Color.FromArgb(0x88, 0xAA, 0xCC), Location = new Point(20, 122), Size = new Size(180, 20) };
                var descTb    = new TextBox{ Location = new Point(20, 144), Size = new Size(360, 30), BackColor = Color.FromArgb(0x0F, 0x1E, 0x32), ForeColor = Color.FromArgb(0xCC, 0xDD, 0xEE), BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) };
                var saveBtn   = MakeDialogButton("Guardar",  new Point(100, 192), Color.FromArgb(0x0A, 0x60, 0x30), DialogResult.OK);
                var cancelBtn = MakeDialogButton("Cancelar", new Point(220, 192), Color.FromArgb(0x60, 0x0A, 0x0A), DialogResult.Cancel);

                dlg.Controls.AddRange(new Control[] { titleLbl, nameLbl, nameTb, descLbl, descTb, saveBtn, cancelBtn });
                dlg.AcceptButton = saveBtn;
                dlg.CancelButton = cancelBtn;

                if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(nameTb.Text))
                    SaveLayoutWithName(nameTb.Text, descTb.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al mostrar diálogo: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static Button MakeDialogButton(string text, Point location, Color back, DialogResult result)
        {
            var b = new Button
            {
                Text             = text,
                Location         = location,
                Size             = new Size(100, 35),
                BackColor        = back,
                ForeColor        = Color.FromArgb(0xCC, 0xDD, 0xEE),
                FlatStyle        = FlatStyle.Flat,
                Font             = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                DialogResult     = result,
                UseVisualStyleBackColor = false
            };
            b.FlatAppearance.BorderSize  = 1;
            b.FlatAppearance.BorderColor = Color.FromArgb(0x1C, 0x34, 0x55);
            return b;
        }

        private void SaveLayoutWithName(string layoutName, string description)
        {
            try
            {
                var positions = GetCurrentTabPositions();
                WindowPositionManager.SaveConfiguration(layoutName, positions, description);
                MessageBox.Show($"Layout '{layoutName}' guardado con {positions.Count} ventanas.",
                    "Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar layout: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<WindowPositionManager.WindowPosition> GetCurrentTabPositions()
        {
            var positions = new List<WindowPositionManager.WindowPosition>();
            for (int i = 0; i < _tabs.TabCount; i++)
            {
                var info = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == _tabs.TabPages[i]);
                if (info != null)
                {
                    string windowName = GetWindowTitle(info.Hwnd).Trim();
                    if (string.IsNullOrWhiteSpace(windowName)) windowName = TARGET_PROCESS_NAME;
                    positions.Add(new WindowPositionManager.WindowPosition { WindowName = windowName, Position = i });
                }
            }
            return positions;
        }

        private void QuickRestoreLayout()
        {
            try
            {
                var layouts = WindowPositionManager.GetConfigurationNames();
                if (layouts.Count == 0)
                {
                    MessageBox.Show("No hay layouts guardados.", "Sin Layouts", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                OpenLayoutManager();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenLayoutManager()
        {
            try
            {
                using var mgr = new LayoutSelectorForm(_preferredLayoutName);
                mgr.ShowDialog(this);

                if (mgr.DialogResult == DialogResult.OK)
                {
                    if (mgr.ShouldLoad) RestoreLayout(mgr.SelectedLayout!);
                    else                SaveLayout(mgr.SelectedLayout!);
                }

                if (mgr.PreferredLayoutChanged)
                {
                    _preferredLayoutName = mgr.NewPreferredLayout;
                    UpdatePreferredButton();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveLayout(string layoutName)
        {
            try
            {
                var positions = GetCurrentTabPositions();
                WindowPositionManager.SaveConfiguration(layoutName, positions);
                MessageBox.Show($"Layout '{layoutName}' guardado con {positions.Count} ventanas.",
                    "Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RestoreLayout(string layoutName)
        {
            try
            {
                var config = WindowPositionManager.LoadConfiguration(layoutName);
                if (config == null)
                {
                    MessageBox.Show($"No se encontró el layout '{layoutName}'.",
                        "Layout No Encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var currentWindows = new Dictionary<string, TabPage>();
                for (int i = 0; i < _tabs.TabCount; i++)
                {
                    var info = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == _tabs.TabPages[i]);
                    if (info != null)
                    {
                        string name = GetWindowTitle(info.Hwnd).Trim();
                        if (string.IsNullOrWhiteSpace(name)) name = TARGET_PROCESS_NAME;
                        currentWindows[name] = _tabs.TabPages[i];
                    }
                }

                var orderedTabs = new List<TabPage>();
                foreach (var pos in config.Positions.OrderBy(p => p.Position))
                    if (currentWindows.Remove(pos.WindowName, out var tp))
                        orderedTabs.Add(tp);
                foreach (var remaining in currentWindows.Values)
                    orderedTabs.Add(remaining);

                _tabs.TabPages.Clear();
                foreach (var t in orderedTabs) _tabs.TabPages.Add(t);

                ReorderTabs();
                MessageBox.Show($"Layout '{layoutName}' restaurado.",
                    "Listo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadPreferredLayout()
        {
            if (string.IsNullOrWhiteSpace(_preferredLayoutName))
            {
                MessageBox.Show(
                    "No hay ningún layout preferido configurado.\n\nUsa '📋 Layouts' y marca uno como preferido.",
                    "Sin Layout Preferido", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            RestoreLayout(_preferredLayoutName);
        }

        public void SetPreferredLayout(string? layoutName)
        {
            _preferredLayoutName = layoutName;
            UpdatePreferredButton();
        }

        private void UpdatePreferredButton()
        {
            if (string.IsNullOrWhiteSpace(_preferredLayoutName))
            {
                _loadPreferredButton.Text        = "⭐  Louadout Favorito";
                _loadPreferredButton.ToolTipText = "Ningún layout preferido configurado";
            }
            else
            {
                _loadPreferredButton.Text        = $"⭐  {_preferredLayoutName}";
                _loadPreferredButton.ToolTipText = $"Cargar layout preferido: {_preferredLayoutName}";
            }
        }

        internal void ActivateTabByCharacterName(string characterName)
        {
            for (int i = 0; i < _tabs.TabCount; i++)
            {
                var info = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == _tabs.TabPages[i]);
                if (info != null)
                {
                    string windowTitle = GetWindowTitle(info.Hwnd).Trim();
                    if (windowTitle.StartsWith(characterName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (_tabs.SelectedIndex != i)
                            _tabs.SelectedIndex = i;
                        ForceWindowToForeground();
                        return;
                    }
                }
            }
        }

        private void ForceWindowToForeground()
        {
            try
            {
                IntPtr hwnd     = Handle;
                IntPtr fgWindow = GetForegroundWindow();

                if (WindowState == FormWindowState.Minimized)
                {
                    WindowState = FormWindowState.Maximized;
                    Show();
                }

                if (fgWindow != hwnd && fgWindow != IntPtr.Zero)
                {
                    uint fgThread = GetWindowThreadProcessId(fgWindow, out _);
                    uint myThread = GetCurrentThreadId();

                    if (fgThread != myThread)
                    {
                        AttachThreadInput(fgThread, myThread, true);
                        try
                        {
                            BringWindowToTop(hwnd);
                            SetForegroundWindow(hwnd);
                        }
                        finally
                        {
                            AttachThreadInput(fgThread, myThread, false);
                        }
                    }
                }

                SetWindowPos(hwnd, HWND_TOPMOST,   0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                Activate();
                Focus();

                var flashInfo = new FLASHWINFO
                {
                    cbSize    = (uint)Marshal.SizeOf<FLASHWINFO>(),
                    hwnd      = hwnd,
                    dwFlags   = FLASHW_ALL | FLASHW_TIMERNOFG,
                    uCount    = 3,
                    dwTimeout = 0
                };
                FlashWindowEx(ref flashInfo);

                var currentTab = _tabs.SelectedTab;
                if (currentTab != null)
                {
                    var active = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == currentTab);
                    if (active != null)
                    {
                        var fd = new FLASHWINFO
                        {
                            cbSize    = (uint)Marshal.SizeOf<FLASHWINFO>(),
                            hwnd      = active.Hwnd,
                            dwFlags   = FLASHW_ALL,
                            uCount    = 2,
                            dwTimeout = 0
                        };
                        FlashWindowEx(ref fd);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ForceWindowToForeground] Error: {ex.Message}");
            }
        }

        private void OnFormClosingRestoreWindows(object? sender, FormClosingEventArgs e)
        {
            _monitor?.Stop();
            _trayIcon.Visible = false;
            _trayIcon.Dispose();

            foreach (var item in _embeddedByHwnd.Values)
            {
                SetParent(item.Hwnd, IntPtr.Zero);
                SetWindowLong(item.Hwnd, GWL_STYLE, item.OriginalStyle);
                SetWindowPos(item.Hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
                ShowWindow(item.Hwnd, SW_SHOW);
            }

            UnregisterHotKey(Handle, HOTKEY_ID_PREV);
            UnregisterHotKey(Handle, HOTKEY_ID_NEXT);
            UnregisterHotKey(Handle, HOTKEY_ID_TOGGLE_MENU);
            UnregisterHotKey(Handle, HOTKEY_ID_SAVE_POS);
            UnregisterHotKey(Handle, HOTKEY_ID_RESTORE_POS);
            UnregisterHotKey(Handle, HOTKEY_ID_MANAGE);
            for (int i = 0; i < 9; i++) UnregisterHotKey(Handle, HOTKEY_ID_NUM_START + i);
        }

        private static string GetWindowTitle(IntPtr hwnd)
        {
            var sb = new StringBuilder(512);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private sealed class EmbeddedWindowInfo(IntPtr hwnd, Panel hostPanel, TabPage tabPage, int originalStyle)
        {
            public IntPtr  Hwnd          { get; } = hwnd;
            public Panel   HostPanel     { get; } = hostPanel;
            public TabPage TabPage       { get; } = tabPage;
            public int     OriginalStyle { get; } = originalStyle;
            public Size    LastKnownSize { get; set; } = Size.Empty;
        }
    }
}
