/*
 * DofusNetMonitor — Lógica de captura de red de DofusMonitor
 * Integrada en el proceso de Wintabber Dofus (sin proceso externo ni IPC).
 *
 * Canal de notificación: llamada directa a Form1.ActivateTabByCharacterName()
 * UDP y Named Pipe ya NO son necesarios — la comunicación es interna.
 */

using System.Collections.Concurrent;
using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using PacketDotNet;
using SharpPcap;

namespace DofusMiniTabber
{
    /// <summary>
    /// Monitor de red para Dofus Retro.
    /// Detecta Autofocus (GTS), Autotrade (ERK) y Autogroup (PIK)
    /// y notifica directamente a la UI sin ningún IPC.
    /// </summary>
    internal sealed class DofusNetMonitor
    {
        // ── Callback hacia la UI ──────────────────────────────────────────────
        private readonly Action<string> _notifyUi;

        // ── Servidores conocidos ──────────────────────────────────────────────
        private readonly Dictionary<string, string> _servers = new()
        {
            ["Allisteria"]  = "34.253.140.241",
            ["Allisteria2"] = "18.200.38.104",
            ["Fallaster"]   = "34.255.49.243",
            ["Fallaster2"]  = "54.228.180.96",
        };
        private readonly HashSet<string> _remoteIPs;
        private const int    RemotePort = 443;

        // ── Estado ────────────────────────────────────────────────────────────
        private readonly ConcurrentDictionary<string, CharInfo>        _allDetected   = new();
        private readonly ConcurrentDictionary<string, int>             _idToPort      = new();
        private readonly ConcurrentDictionary<string, string>          _idToName      = new();
        private readonly List<string>                                   _slotOrder     = [];
        private readonly object                                         _slotLock      = new();
        private readonly CancellationTokenSource                        _cts           = new();

        // ── Buffer TCP por flujo para reassembly ──────────────────────────────
        private readonly ConcurrentDictionary<string, StringBuilder>   _streamBuffers = new();
        private readonly HashSet<string>                                _tlsWarnedFlows = [];

        // ── Regexes ───────────────────────────────────────────────────────────
        private static readonly Regex RxHost = new(@"dofusretro-[\w\-]+\.ankama-games\.com", RegexOptions.Compiled);
        private static readonly Regex RxAsk  = new(@"ASK\|(\d+)\|([^|]+)\|",  RegexOptions.Compiled);
        private static readonly Regex RxGts  = new(@"GTS(\d+)\|",              RegexOptions.Compiled);
        private static readonly Regex RxErk  = new(@"ERK(\d+)\|(\d+)\|",       RegexOptions.Compiled);
        private static readonly Regex RxPik  = new(@"PIK([^|]+)\|([^|\x00\r\n]+)", RegexOptions.Compiled);

        // ── Features ──────────────────────────────────────────────────────────
        public volatile bool FeatureAutofocus = true;
        public volatile bool FeatureAutotrade = true;
        public volatile bool FeatureAutogroup = true;

        // ── Contadores ────────────────────────────────────────────────────────
        private long _totalPkts;
        private long _dofusPkts;
        private long _processedPkts;

        // ── Log ───────────────────────────────────────────────────────────────
        private static readonly object LogLock = new();

        // ── Estado público ────────────────────────────────────────────────────
        public bool IsRunning { get; private set; }
        public string StatusMessage { get; private set; } = "Detenido";

        // ── Evento de log para la UI ──────────────────────────────────────────
        public event Action<string, string>? OnLog; // (tag, mensaje)

        // ═════════════════════════════════════════════════════════════════════
        public DofusNetMonitor(Action<string> notifyUi)
        {
            _notifyUi  = notifyUi;
            _remoteIPs = [.. _servers.Values];
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Arranque / parada
        // ═════════════════════════════════════════════════════════════════════
        public void Start()
        {
            if (IsRunning) return;
            IsRunning = true;
            StatusMessage = "Iniciando...";

            new Thread(RunCapture) { IsBackground = true, Name = "DofusNetMonitor" }.Start();
            new Thread(MonitorDisconnections) { IsBackground = true, Name = "DisconnMonitor" }.Start();
        }

        public void Stop()
        {
            _cts.Cancel();
            IsRunning = false;
            StatusMessage = "Detenido";
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Captura — abre TODAS las interfaces reales
        // ═════════════════════════════════════════════════════════════════════
        private void RunCapture()
        {
            try
            {
                var devs   = CaptureDeviceList.Instance;
                var opened = new List<ICaptureDevice>();

                foreach (var dev in devs)
                {
                    string desc = dev.Description;
                    if (desc.Contains("Loopback",     StringComparison.OrdinalIgnoreCase) ||
                        desc.Contains("WAN Miniport", StringComparison.OrdinalIgnoreCase))
                        continue;

                    try
                    {
                        dev.Open(new DeviceConfiguration { Mode = DeviceModes.Promiscuous, ReadTimeout = 1 });
                        dev.Filter = $"tcp and port {RemotePort}";
                        dev.OnPacketArrival += OnPacketArrival;
                        dev.StartCapture();
                        opened.Add(dev);
                        Log("NET", $"Capturando: {desc}");
                    }
                    catch (Exception ex)
                    {
                        Log("WARN", $"No se pudo abrir '{desc}': {ex.Message}");
                    }
                }

                if (opened.Count == 0)
                {
                    StatusMessage = "Sin interfaces — ejecuta como Administrador y verifica Npcap";
                    Log("ERROR", StatusMessage);
                    IsRunning = false;
                    return;
                }

                StatusMessage = $"Escuchando en {opened.Count} interfaz(es)";
                Log("NET", StatusMessage);

                _cts.Token.WaitHandle.WaitOne();

                foreach (var dev in opened)
                    try { dev.StopCapture(); dev.Close(); } catch { }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error de captura: {ex.Message}";
                Log("ERROR", StatusMessage);
                IsRunning = false;
            }
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Llegada de paquete - CON LOGGING DE TIMING
        // ═════════════════════════════════════════════════════════════════════
        private void OnPacketArrival(object _, PacketCapture e)
        {
            if (_cts.IsCancellationRequested) return;
            Interlocked.Increment(ref _totalPkts);

            var arrivalTime = DateTime.Now;

            try
            {
                var pkt = Packet.ParsePacket(e.GetPacket().LinkLayerType, e.GetPacket().Data);
                var ip  = pkt.Extract<IPv4Packet>();
                var tcp = pkt.Extract<TcpPacket>();
                if (ip == null || tcp == null) return;

                string src = ip.SourceAddress.ToString();
                string dst = ip.DestinationAddress.ToString();

                bool fromServer = _remoteIPs.Contains(src);
                bool toServer   = _remoteIPs.Contains(dst);
                if (!fromServer && !toServer) return;

                Interlocked.Increment(ref _dofusPkts);

                int payload = tcp.PayloadData?.Length ?? 0;
                if (payload == 0) return;

                byte first = tcp.PayloadData![0];
                if (first >= 0x14 && first <= 0x17)
                {
                    string flowKey = $"{src}:{tcp.SourcePort}->{dst}:{tcp.DestinationPort}";
                    if (_tlsWarnedFlows.Add(flowKey))
                        Log("TLS!", $"Flujo {flowKey} usa TLS — cifrado, se ignora.");
                    return;
                }

                int    localPort = fromServer ? tcp.DestinationPort : tcp.SourcePort;
                string server    = GetServerName(fromServer ? src : dst);
                string flowId    = $"{src}:{tcp.SourcePort}->{dst}:{tcp.DestinationPort}";

                // Log de timing para debugging
                string payloadPreview = Encoding.Latin1.GetString(tcp.PayloadData!, 0, Math.Min(50, payload));
                if (payloadPreview.Contains("GTS") || payloadPreview.Contains("ERK") || payloadPreview.Contains("PIK"))
                    Log("TIMING", $"Packet arrived at {arrivalTime:HH:mm:ss.fff}, payload: {payloadPreview[..Math.Min(30, payloadPreview.Length)]}");

                FeedStream(flowId, tcp.PayloadData, localPort, server, arrivalTime);
            }
            catch { }
        }

        // ═════════════════════════════════════════════════════════════════════
        //  TCP Stream Reassembly
        // ═════════════════════════════════════════════════════════════════════
        private void FeedStream(string flowId, byte[] data, int localPort, string server, DateTime arrivalTime)
        {
            string chunk = Encoding.Latin1.GetString(data);
            var buf = _streamBuffers.GetOrAdd(flowId, _ => new StringBuilder());

            lock (buf)
            {
                buf.Append(chunk);
                string full = buf.ToString();

                ProcessUrgentPatterns(full, localPort, server, arrivalTime);

                int splitPos;
                while ((splitPos = IndexOfTerminator(full)) >= 0)
                {
                    string msg = full[..splitPos].TrimEnd('\r', '\n', '\x00');
                    full = full[(splitPos + 1)..];
                    if (msg.Length > 0)
                        ProcessMessage(msg, localPort, server);
                }

                if (full.Length > 8192)
                {
                    Log("WARN", $"Buffer {flowId} muy grande ({full.Length}b) — descartando");
                    full = "";
                }

                buf.Clear();
                buf.Append(full);
            }

            Interlocked.Increment(ref _processedPkts);
        }

        private void ProcessUrgentPatterns(string buffer, int localPort, string server, DateTime arrivalTime)
        {
            var now = DateTime.Now;
            var elapsedMs = (now - arrivalTime).TotalMilliseconds;

            if (FeatureAutofocus)
            {
                foreach (Match m in RxGts.Matches(buffer))
                {
                    string cid = m.Groups[1].Value;
                    if (_idToPort.TryGetValue(cid, out int p) && p == localPort &&
                        _idToName.TryGetValue(cid, out string? name))
                    {
                        Log("FOCUS", $"¡TURNO! {name} (detect:{elapsedMs:F1}ms ago)");
                        _notifyUi(name);
                    }
                }
            }

            if (FeatureAutotrade)
            {
                foreach (Match m in RxErk.Matches(buffer))
                {
                    string recId = m.Groups[2].Value;
                    if (_idToPort.ContainsKey(recId) && _idToName.TryGetValue(recId, out string? recN))
                    {
                        Log("TRADE", $"¡TRADE! {recN} (detect:{elapsedMs:F1}ms ago)");
                        _notifyUi(recN);
                    }
                }
            }

            if (FeatureAutogroup)
            {
                foreach (Match m in RxPik.Matches(buffer))
                {
                    string recN = m.Groups[2].Value.TrimEnd('\0', '\r', '\n', ' ');
                    if (!string.IsNullOrEmpty(recN))
                    {
                        Log("GROUP", $"¡INVITACIÓN! -> {recN} (detect:{elapsedMs:F1}ms ago)");
                        _notifyUi(recN);
                    }
                }
            }
        }

        private static int IndexOfTerminator(string s)
        {
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\x00') return i;
                if (c == '.' && i + 1 < s.Length)
                {
                    char next = s[i + 1];
                    if (next == '\n' || next == '\r' || next == '\x00') return i + 1;
                }
            }
            return -1;
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Procesamiento de mensaje completo
        // ═════════════════════════════════════════════════════════════════════
        private void ProcessMessage(string text, int localPort, string server)
        {
            if (IsLikelyEncrypted(text)) return;

            var hm = RxHost.Match(text);
            if (hm.Success) TryResolveNewServer(hm.Value);

            foreach (Match m in RxAsk.Matches(text))
            {
                string cid  = m.Groups[1].Value;
                string name = m.Groups[2].Value;
                Log("MATCH", $"ASK id={cid} name={name} puerto={localPort} server={server}");
                RegisterChar(cid, localPort, name, server);
            }

            if (FeatureAutofocus)
            {
                foreach (Match m in RxGts.Matches(text))
                {
                    string cid = m.Groups[1].Value;
                    if (!_idToPort.TryGetValue(cid, out int p) || p != localPort) continue;
                    if (!_idToName.TryGetValue(cid, out string? name)) continue;
                    Log("FOCUS", $"Turno: {name}  slot={GetSlot(name)}");
                    _notifyUi(name);
                }
            }

            if (FeatureAutotrade)
            {
                foreach (Match m in RxErk.Matches(text))
                {
                    string emiId = m.Groups[1].Value;
                    string recId = m.Groups[2].Value;
                    string emiN  = _idToName.GetValueOrDefault(emiId, $"<{emiId}>");
                    string recN  = _idToName.GetValueOrDefault(recId, $"<{recId}>");
                    if (!_idToPort.ContainsKey(recId)) continue;
                    Log("TRADE", $"{emiN} -> {recN}  slot={GetSlot(recN)}");
                    _notifyUi(recN);
                }
            }

            if (FeatureAutogroup)
            {
                foreach (Match m in RxPik.Matches(text))
                {
                    string emiN = m.Groups[1].Value;
                    string recN = m.Groups[2].Value.TrimEnd('\0', '\r', '\n', ' ');
                    Log("GROUP", $"Invitacion: {emiN} -> {recN}  slot={GetSlot(recN)}");
                    _notifyUi(recN);
                }
            }
        }

        private static bool IsLikelyEncrypted(string s)
        {
            if (s.Length < 20) return false;
            if (s.Contains('|') || s.Contains(';')) return false;
            int b64 = s.Count(c => char.IsLetterOrDigit(c) || c == '+' || c == '/' || c == '=');
            return (double)b64 / s.Length > 0.92;
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Gestión de personajes
        // ═════════════════════════════════════════════════════════════════════
        private void RegisterChar(string charId, int port, string name, string server)
        {
            bool isNew = !_allDetected.ContainsKey(charId) || _allDetected[charId].Port != port;
            _allDetected[charId] = new CharInfo(name, port, server);
            _idToPort[charId]    = port;
            _idToName[charId]    = name;
            if (!isNew) return;
            lock (_slotLock) if (!_slotOrder.Contains(name)) _slotOrder.Add(name);
            Log("CHAR", $"[{server}] {name}  id={charId}  puerto={port}  slot={GetSlot(name)}");
        }

        private int GetSlot(string name)
        {
            lock (_slotLock) { int i = _slotOrder.IndexOf(name); return i >= 0 ? i + 1 : 0; }
        }

        private string GetServerName(string ip)
        {
            foreach (var kv in _servers) if (kv.Value == ip) return kv.Key;
            return ip;
        }

        private void TryResolveNewServer(string host)
        {
            try
            {
                foreach (var addr in System.Net.Dns.GetHostAddresses(host))
                {
                    string ip = addr.ToString();
                    if (_remoteIPs.Add(ip))
                    {
                        string s = host.Contains('-') ? host.Split('-')[1].Split('.')[0] : host;
                        _servers[s] = ip;
                        Log("NET", $"Nuevo servidor: {host} -> {ip}");
                    }
                }
            }
            catch { }
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Monitor desconexiones
        // ═════════════════════════════════════════════════════════════════════
        private void MonitorDisconnections()
        {
            while (!_cts.IsCancellationRequested)
            {
                try
                {
                    var active = IPGlobalProperties.GetIPGlobalProperties()
                        .GetActiveTcpConnections()
                        .Where(c => c.State is TcpState.Established or TcpState.SynSent)
                        .Select(c => c.LocalEndPoint.Port).ToHashSet();

                    foreach (var (cid, info) in _allDetected.Where(kv => !active.Contains(kv.Value.Port)).ToList())
                    {
                        _allDetected.TryRemove(cid, out _);
                        _idToPort.TryRemove(cid, out _);
                        _idToName.TryRemove(cid, out _);
                        lock (_slotLock) _slotOrder.Remove(info.Name);
                        Log("DISC", $"Desconectado: {info.Name}  puerto={info.Port}");
                    }
                }
                catch { }

                _cts.Token.WaitHandle.WaitOne(5_000);
            }
        }

        // ═════════════════════════════════════════════════════════════════════
        //  Estadísticas públicas
        // ═════════════════════════════════════════════════════════════════════
        public (long total, long dofus, long procesados, int chars) GetStats() =>
            (Interlocked.Read(ref _totalPkts),
             Interlocked.Read(ref _dofusPkts),
             Interlocked.Read(ref _processedPkts),
             _allDetected.Count);

        // ═════════════════════════════════════════════════════════════════════
        //  Log
        // ═════════════════════════════════════════════════════════════════════
        private void Log(string tag, string msg)
        {
            OnLog?.Invoke(tag, msg);
            Debug.WriteLine($"[{tag}] {msg}");
        }
    }

    // ── Record compartido ─────────────────────────────────────────────────────
    internal record CharInfo(string Name, int Port, string Server);
}
