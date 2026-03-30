/*
 * Form1.cs — Wintabber Dofus  [FUSIONADO con DofusMonitor]
 *
 * Cambios respecto al Wintabber original:
 *  - UDP listener y Named Pipe SERVER eliminados (ya no se necesitan).
 *  - DofusNetMonitor corre en este mismo proceso y llama directamente
 *    a ActivateTabByCharacterName() sin ningún IPC de por medio.
 *  - Se añade en la barra de herramientas un indicador de estado del monitor
 *    y un botón para iniciar/detener la captura de red.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DofusMiniTabber
{
    public partial class Form1 : Form
    {
        // ── UI ────────────────────────────────────────────────────────────────
        private readonly ToolStrip       _toolbar              = new();
        private readonly ToolStrip       _floatingToolbar      = new();
        private readonly ToolStripButton _captureButton        = new("⚡ CAPTURAR VENTANAS");
        private readonly ToolStripButton _savePositionButton   = new("💾 GUARDAR LAYOUT");
        private readonly ToolStripButton _restorePositionButton= new("🔄 CARGAR LAYOUT");
        private readonly ToolStripButton _loadPreferredButton  = new("⭐ CARGAR PREFERIDO");
        private readonly ToolStripButton _manageLayoutsButton  = new("📋 GESTIONAR LAYOUTS");
        private readonly ToolStripButton _hideMenuButton       = new("👁️ OCULTAR MENÚ");
        private readonly ToolStripButton _monitorButton        = new("🔴 MONITOR: OFF");
        private readonly ToolStripLabel  _monitorStatus        = new("Conectado");
        private readonly ToolStripLabel  _hotkeysLabel         = new("[F1/F2] Ant/Sig | [F3] Menú | [F4] Guardar | [F5] Cargar | [F6] Gestionar | [Ctrl+Alt+1..9] Directo");
        private readonly TabControl      _tabs                 = new();
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
        private const uint FLASHW_ALL = 3; // FLASHW_CAPTION | FLASHW_TRAY
        private const uint FLASHW_TIMERNOFG = 12;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
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

            // Timer de estadísticas del monitor - oculto para app friendly
            _statsTimer.Interval = 15_000;
            _statsTimer.Tick += StatsTimer_Tick;
            _statsTimer.Start();
            _monitorStatus.Visible = false;  // Oculto: no mostrar estadísticas de paquetes

            // Intentar iniciar el monitor de red automáticamente
            TryStartMonitor();
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Monitor de red — arranque / parada
        // ═════════════════════════════════════════════════════════════════════
        private void TryStartMonitor()
        {
            try
            {
                _monitor = new DofusNetMonitor(name =>
                {
                    // Llamada directa al UI thread (sin UDP ni pipe)
                    // Usar BeginInvoke para no bloquear el thread de captura de red
                    if (IsHandleCreated)
                        BeginInvoke(() => ActivateTabByCharacterName(name));
                });

                _monitor.OnLog += (tag, msg) =>
                {
                    // Actualizar status label con los mensajes más relevantes
                    if (tag is "NET" or "ERROR" or "WARN" or "FOCUS" or "TRADE" or "GROUP")
                        BeginInvoke(() => _monitorStatus.Text = $"[{tag}] {msg}"[..Math.Min($"[{tag}] {msg}".Length, 60)]);
                };

                _monitor.Start();
                UpdateMonitorButton(running: true);
            }
            catch (Exception ex)
            {
                // Npcap no instalado o sin permisos → el tabber sigue funcionando
                UpdateMonitorButton(running: false);
                _monitorStatus.Text = $"Monitor no disponible: {ex.Message}"[..Math.Min($"Monitor no disponible: {ex.Message}".Length, 60)];
            }
        }

        private void ToggleMonitor()
        {
            if (_monitor?.IsRunning == true)
            {
                _monitor.Stop();
                UpdateMonitorButton(running: false);
                _monitorStatus.Text = "Monitor detenido";
            }
            else
            {
                TryStartMonitor();
            }
        }

        private void UpdateMonitorButton(bool running)
        {
            _monitorButton.Text      = running ? "🟢 MONITOR: ON" : "🔴 MONITOR: OFF";
            _monitorButton.BackColor = running
                ? Color.FromArgb(0x1A, 0x5C, 0x2A)
                : Color.FromArgb(0x5C, 0x1A, 0x1A);
        }

        private void StatsTimer_Tick(object? sender, EventArgs e)
        {
            if (_monitor?.IsRunning == true)
            {
                var (total, dofus, proc, chars) = _monitor.GetStats();
                _monitorStatus.Text = $"pkts:{total}  dofus:{dofus}  proc:{proc}  chars:{chars}";
            }
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Configuración del formulario
        // ═════════════════════════════════════════════════════════════════════
        private void ConfigureForm()
        {
            Text        = "Wintabber Dofus";
            BackColor   = Color.FromArgb(0x0F, 0x19, 0x23);
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

            var restoreItem   = new ToolStripMenuItem("🖥️ Restaurar",        null, (_, _) => RestoreWindow());
            var captureItem   = new ToolStripMenuItem("⚡ Capturar ventanas", null, (_, _) => CaptureWindows());
            var monitorItem   = new ToolStripMenuItem("🔄 Activar auto-cambio",   null, (_, _) => ToggleMonitor());
            var separatorItem = new ToolStripSeparator();
            var exitItem      = new ToolStripMenuItem("❌ Salir",             null, (_, _) => ExitApplication());

            _trayMenu.Items.Add(restoreItem);
            _trayMenu.Items.Add(captureItem);
            _trayMenu.Items.Add(monitorItem);
            _trayMenu.Items.Add(separatorItem);
            _trayMenu.Items.Add(exitItem);
            _trayMenu.BackColor  = Color.FromArgb(0x1E, 0x2A, 0x38);
            _trayMenu.ForeColor  = Color.White;
            _trayMenu.RenderMode = ToolStripRenderMode.System;

            _trayIcon.ContextMenuStrip = _trayMenu;
            _trayIcon.DoubleClick += (_, _) => RestoreWindow();
        }

        private static Icon CreateFallbackIcon()
        {
            using var bmp  = new Bitmap(16, 16);
            using var g    = Graphics.FromImage(bmp);
            g.Clear(Color.FromArgb(0x1E, 0x2A, 0x38));
            using var font = new Font("Arial", 8f, FontStyle.Bold);
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
        //  Construcción de la UI
        // ═════════════════════════════════════════════════════════════════════
        private void BuildUi()
        {
            SuspendLayout();

            _toolbar.Dock    = DockStyle.Top;
            _toolbar.Height  = 1;
            _toolbar.Visible = false;

            _floatingToolbar.Dock        = DockStyle.Top;
            _floatingToolbar.BackColor   = Color.FromArgb(0x1E, 0x2A, 0x38);
            _floatingToolbar.ForeColor   = Color.White;
            _floatingToolbar.GripStyle   = ToolStripGripStyle.Hidden;
            _floatingToolbar.CanOverflow = false;
            _floatingToolbar.Stretch     = true;

            _captureButton.Click          += (_, _) => CaptureWindows();
            _savePositionButton.Click     += (_, _) => SaveCurrentPositions();
            _restorePositionButton.Click  += (_, _) => QuickRestoreLayout();
            _loadPreferredButton.Click    += (_, _) => LoadPreferredLayout();
            _manageLayoutsButton.Click    += (_, _) => OpenLayoutManager();
            _hideMenuButton.Click         += (_, _) => ToggleFloatingMenu();
            _monitorButton.Click          += (_, _) => ToggleMonitor();

            // Botón monitor con color de fondo
            _monitorButton.BackColor = Color.FromArgb(0x5C, 0x1A, 0x1A);
            _monitorButton.ForeColor = Color.White;

            // Label de status del monitor - oculto (a la derecha)
            _monitorStatus.ForeColor  = Color.FromArgb(0x80, 0xCC, 0x80);
            _monitorStatus.Alignment  = ToolStripItemAlignment.Right;
            _monitorStatus.Visible    = false;  // Oculto para app friendly
            _hotkeysLabel.ForeColor   = Color.FromArgb(0x80, 0x80, 0x80);
            _hotkeysLabel.Alignment   = ToolStripItemAlignment.Right;

            _floatingToolbar.Items.Add(_captureButton);
            _floatingToolbar.Items.Add(new ToolStripSeparator());
            _floatingToolbar.Items.Add(_savePositionButton);
            _floatingToolbar.Items.Add(_restorePositionButton);
            _floatingToolbar.Items.Add(_loadPreferredButton);
            _floatingToolbar.Items.Add(_manageLayoutsButton);
            _floatingToolbar.Items.Add(new ToolStripSeparator());
            _floatingToolbar.Items.Add(_hideMenuButton);
            _floatingToolbar.Items.Add(new ToolStripSeparator());
            // _monitorButton REMOVIDO - oculto para app friendly
            _floatingToolbar.Items.Add(_hotkeysLabel);
            // _monitorStatus REMOVIDO - oculto para app friendly

            _tabs.Dock                 = DockStyle.Fill;
            _tabs.SelectedIndexChanged += (_, _) => OnTabChanged();
            _tabs.MouseDown            += Tabs_MouseDown;
            _tabs.MouseMove            += Tabs_MouseMove;
            _tabs.MouseUp              += Tabs_MouseUp;

            var liberarItem = new ToolStripMenuItem("Liberar", null, (_, _) => ReleaseCurrentTab());
            var cerrarItem  = new ToolStripMenuItem("Cerrar",  null, (_, _) => CloseCurrentTab());
            _tabMenu.Items.Add(liberarItem);
            _tabMenu.Items.Add(cerrarItem);

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

        private void ToggleFloatingMenu()
        {
            _menuVisible             = !_menuVisible;
            _floatingToolbar.Visible = _menuVisible;
            _hideMenuButton.Text     = _menuVisible ? "👁️ OCULTAR MENÚ" : "👁️‍🗨️ MOSTRAR MENÚ";
            ScheduleResizeActiveTab();
        }

        // ── Drag & drop tabs ──────────────────────────────────────────────────
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

        // ── Release / close tab ───────────────────────────────────────────────
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

        // ── Hotkeys ───────────────────────────────────────────────────────────
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

        // ── WndProc ───────────────────────────────────────────────────────────
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

        // ── Captura de ventanas ───────────────────────────────────────────────
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

        // ── Tab switching ─────────────────────────────────────────────────────
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

        // ── Títulos dinámicos ─────────────────────────────────────────────────
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

        // ── Gestión de layouts ────────────────────────────────────────────────
        private void SaveCurrentPositions()
        {
            try
            {
                using var dlg = new Form
                {
                    Text            = "Guardar Layout",
                    Size            = new Size(420, 250),
                    StartPosition   = FormStartPosition.CenterParent,
                    BackColor       = Color.FromArgb(0x0F, 0x19, 0x23),
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox     = false,
                    MinimizeBox     = false,
                    Font            = new Font("Segoe UI", 9F)
                };

                var titleLbl  = new Label  { Text = " Guardar Configuración Actual", ForeColor = Color.White, Location = new Point(20, 15),  Size = new Size(350, 30), Font = new Font("Segoe UI", 12F, FontStyle.Bold) };
                var nameLbl   = new Label  { Text = "Nombre del Layout:",             ForeColor = Color.White, Location = new Point(20, 60),  Size = new Size(120, 20) };
                var nameTb    = new TextBox{ Location = new Point(20, 85),  Size = new Size(360, 30), BackColor = Color.FromArgb(0x1E, 0x2A, 0x38), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) };
                var descLbl   = new Label  { Text = "Descripción (opcional):",        ForeColor = Color.White, Location = new Point(20, 125), Size = new Size(150, 20) };
                var descTb    = new TextBox{ Location = new Point(20, 150), Size = new Size(360, 30), BackColor = Color.FromArgb(0x1E, 0x2A, 0x38), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) };
                var saveBtn   = MakeDialogButton(" GUARDAR",   new Point(100, 195), Color.FromArgb(0x28, 0xA7, 0x45), DialogResult.OK);
                var cancelBtn = MakeDialogButton(" CANCELAR",  new Point(220, 195), Color.FromArgb(0xDC, 0x35, 0x45), DialogResult.Cancel);

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
                ForeColor        = Color.White,
                FlatStyle        = FlatStyle.Flat,
                Font             = new Font("Segoe UI", 10F, FontStyle.Bold),
                DialogResult     = result,
                UseVisualStyleBackColor = false
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private void SaveLayoutWithName(string layoutName, string description)
        {
            try
            {
                var positions = GetCurrentTabPositions();
                WindowPositionManager.SaveConfiguration(layoutName, positions, description);
                MessageBox.Show($"Layout '{layoutName}' guardado con {positions.Count} ventanas.",
                    "Guardado Exitoso", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    "Guardado Exitoso", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                MessageBox.Show($"Layout '{layoutName}' restaurado exitosamente.",
                    "Restauración Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    "No hay ningún layout preferido configurado.\n\nUsa '📋 GESTIONAR LAYOUTS' y marca un layout como preferido.",
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
                _loadPreferredButton.Text        = "⭐ CARGAR PREFERIDO";
                _loadPreferredButton.ToolTipText = "Ningún layout preferido configurado";
            }
            else
            {
                _loadPreferredButton.Text        = $"⭐ {_preferredLayoutName}";
                _loadPreferredButton.ToolTipText = $"Cargar layout preferido: {_preferredLayoutName}";
            }
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Activar pestaña por nombre de personaje
        //  Llamado directamente por DofusNetMonitor (sin UDP ni Pipe)
        // ═════════════════════════════════════════════════════════════════════
        internal void ActivateTabByCharacterName(string characterName)
        {
            var startTime = DateTime.Now;
            for (int i = 0; i < _tabs.TabCount; i++)
            {
                var info = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == _tabs.TabPages[i]);
                if (info != null)
                {
                    string windowTitle = GetWindowTitle(info.Hwnd).Trim();
                    if (windowTitle.StartsWith(characterName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (_tabs.SelectedIndex != i)
                        {
                            _tabs.SelectedIndex = i;
                            Debug.WriteLine($"[NOTIFY] Activada pestaña {i} para {characterName}");
                        }
                        // Forzar la ventana al frente inmediatamente
                        ForceWindowToForeground();
                        var totalMs = (DateTime.Now - startTime).TotalMilliseconds;
                        Debug.WriteLine($"[NOTIFY] Total time: {totalMs:F1}ms");
                        return;
                    }
                }
            }
            Debug.WriteLine($"[NOTIFY] No se encontró pestaña para {characterName}");
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Forzar ventana al frente - útil para GTS/Trade/Grupo
        //  Usa AttachThreadInput para saltar las restricciones de Windows
        // ═════════════════════════════════════════════════════════════════════
        private void ForceWindowToForeground()
        {
            try
            {
                IntPtr hwnd = Handle;
                IntPtr fgWindow = GetForegroundWindow();

                // Restaurar si está minimizada
                if (WindowState == FormWindowState.Minimized)
                {
                    WindowState = FormWindowState.Maximized;
                    Show();
                }

                // Técnica 1: AttachThreadInput - enganchar nuestro thread al foreground thread
                // Esto nos permite llamar SetForegroundWindow sin restricciones
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

                // Técnica 2: Hacerla TOPMOST temporalmente y luego quitarlo
                // Esto fuerza a Windows a traerla al frente inmediatamente
                SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);

                Activate();
                Focus();

                // Flash visual para notificar al usuario (también ayuda a forzar atención)
                var flashInfo = new FLASHWINFO
                {
                    cbSize = (uint)Marshal.SizeOf<FLASHWINFO>(),
                    hwnd = hwnd,
                    dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG,
                    uCount = 3,
                    dwTimeout = 0
                };
                FlashWindowEx(ref flashInfo);

                // Forzar también la ventana de Dofus embebida activa al frente
                var currentTab = _tabs.SelectedTab;
                if (currentTab != null)
                {
                    var active = _embeddedByHwnd.Values.FirstOrDefault(v => v.TabPage == currentTab);
                    if (active != null)
                    {
                        // Flash también en la ventana de Dofus
                        var flashDofus = new FLASHWINFO
                        {
                            cbSize = (uint)Marshal.SizeOf<FLASHWINFO>(),
                            hwnd = active.Hwnd,
                            dwFlags = FLASHW_ALL,
                            uCount = 2,
                            dwTimeout = 0
                        };
                        FlashWindowEx(ref flashDofus);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ForceWindowToForeground] Error: {ex.Message}");
            }
        }

        // ── Cierre ────────────────────────────────────────────────────────────
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

        // ── Helpers ───────────────────────────────────────────────────────────
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
