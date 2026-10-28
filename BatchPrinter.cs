using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace BatchPrinter
{
    public class MainForm : Form
    {
        static readonly string[] ShellExts = new string[]
        {
            ".pdf", ".doc", ".docx", ".dot", ".dotx", ".rtf", ".odt",
            ".xls", ".xlsx", ".xlsm", ".ods", ".csv",
            ".ppt", ".pptx", ".odp",
            ".txt", ".log", ".xml", ".htm", ".html", ".xps"
        };
        static readonly string[] ImageExts = new string[]
        {
            ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".emf", ".wmf"
        };
        static readonly string ProfileDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BatchPrinter");
        static readonly string TempRoot = Path.Combine(Path.GetTempPath(), "BatchPrinter");

        // Pool stampato: lo spooler Windows rasterizza in modo asincrono,
        // quindi immagini/documenti/immagini renderizzate NON vanno rilasciate
        // subito, altrimenti crash dopo l'invio alla stampante di rete.
        static readonly List<Image> PrintPool = new List<Image>();
        static readonly List<Stream> StreamPool = new List<Stream>();
        static readonly List<PrintDocument> DocPool = new List<PrintDocument>();

        ListView lvFiles;
        List<string> Files = new List<string>();
        ListBox lstLog;
        ComboBox cboPrinter;
        TextBox numCopies;
        CheckBox chkDuplex;
        CheckBox chkCollate;
        CheckBox chkPreview;
        bool previewOn;
        TextBox txtRange;
        string defaultPrinter = "";
        SplitContainer split;
        Button btnPrint;
        Button btnTheme;
        Label lblLed;
        Label lblTitle;
        Label lblSub;
        Panel pnlHeader;
        Panel pnlSide;
        Panel pnlQueue;
        Panel pnlLog;
        Label lblState;
        Label lblCount;
        Label lblCredit;
        Panel pnlProg;
        double progressFrac = 0.0;

        // --------------------------------------------------- tema Technical Minimal
        class Theme
        {
            public string Key, Name, Bg, Surface, Text, Muted, Hairline, Accent;
        }

        static readonly Theme[] Themes = new Theme[]
        {
            new Theme { Key="light",      Name="LIGHT",      Bg="#f3f2ef", Surface="#fbfaf8", Text="#1c1c1a", Muted="#87847c", Hairline="#dcdad4", Accent="#e5481a" },
            new Theme { Key="mango",      Name="MANGO",      Bg="#f7f3ea", Surface="#fcfaf3", Text="#1c1c1a", Muted="#8f8878", Hairline="#e2dccd", Accent="#d2700a" },
            new Theme { Key="limoncello", Name="LIMONCELLO", Bg="#fffef9", Surface="#ffffff", Text="#1c1c1a", Muted="#97926f", Hairline="#ece9d8", Accent="#b08d00" },
            new Theme { Key="aperol",     Name="APEROL",     Bg="#fff7f2", Surface="#fffdfb", Text="#1c1c1a", Muted="#9a8a80", Hairline="#f0e0d6", Accent="#e24a18" },
            new Theme { Key="spritz",     Name="APEROL SPRITZ", Bg="#fffdf9", Surface="#ffffff", Text="#1c1c1a", Muted="#9c908a", Hairline="#f0e8e0", Accent="#f2622b" },
            new Theme { Key="dark",       Name="DARK",       Bg="#0e0e0e", Surface="#161616", Text="#e8e8e6", Muted="#8f8f8a", Hairline="#262626", Accent="#ff5a1f" },
            new Theme { Key="cactus",     Name="CACTUS",     Bg="#0a0e0b", Surface="#121713", Text="#e6ede8", Muted="#7f8f84", Hairline="#1f2a22", Accent="#35d06a" },
            new Theme { Key="navy",       Name="NAVY MSC",   Bg="#041e42", Surface="#0a2a52", Text="#eaf2ff", Muted="#7c93b8", Hairline="#16345c", Accent="#ffce00" }
        };

        Theme cur;

        static Color Html(string s) { return System.Drawing.ColorTranslator.FromHtml(s); }

        void ApplyTheme(Theme t)
        {
            cur = t;
            Color bg = Html(t.Bg), surface = Html(t.Surface), text = Html(t.Text),
                  muted = Html(t.Muted), hair = Html(t.Hairline), accent = Html(t.Accent);

            BackColor = bg;
            ForeColor = text;

            // header
            pnlHeader.BackColor = bg;
            pnlHeader.Invalidate(); // hairline bottom
            lblLed.BackColor = accent;
            lblTitle.ForeColor = text;
            lblSub.ForeColor = muted;
            lblStatus.ForeColor = muted;
            btnTheme.BackColor = surface;
            btnTheme.ForeColor = text;
            btnTheme.FlatAppearance.BorderColor = hair;
            SetDarkTitle(IsDark(Html(t.Bg)));
            ThemeScrollbars();
            if (lblStatus != null) PlaceStatus();

            // pannelli
            pnlSide.BackColor = bg;
            pnlQueue.BackColor = bg;
            pnlLog.BackColor = bg;
            pnlSide.Invalidate();
            pnlQueue.Invalidate();
            pnlLog.Invalidate();

            // controlli impostazioni
            foreach (Control c in new Control[] { cboPrinter, numCopies, txtRange })
            {
                c.BackColor = surface;
                c.ForeColor = text;
            }
            cboPrinter.FlatStyle = FlatStyle.Flat;
            chkDuplex.BackColor = bg;
            chkDuplex.ForeColor = text;
            chkCollate.BackColor = bg;
            chkCollate.ForeColor = text;
            chkPreview.BackColor = bg;
            chkPreview.ForeColor = text;

            // lista documenti e log
            lvFiles.BackColor = surface;
            lvFiles.Invalidate();
            lstLog.BackColor = surface;
            lstLog.ForeColor = muted;
            if (lblCredit != null) lblCredit.ForeColor = muted;

            // pulsanti tessere (tutti i Button dentro pnlSide e pnlQueue)
            foreach (Control c in pnlSide.Controls)
            {
                Button b = c as Button;
                if (b == null) continue;
                if (b == btnPrint)
                {
                    b.BackColor = accent;
                    // su accent chiari (giallo MSC) il bianco non si legge
                    b.ForeColor = OnAccent(accent);
                    b.FlatAppearance.BorderColor = accent;
                }
                else
                {
                    b.BackColor = surface;
                    b.ForeColor = text;
                    b.FlatAppearance.BorderColor = hair;
                }
            }
            foreach (Control c in pnlQueue.Controls)
            {
                FlowLayoutPanel flp = c as FlowLayoutPanel;
                if (flp == null) continue;
                flp.BackColor = bg;
                foreach (Control c2 in flp.Controls)
                {
                    Button b = c2 as Button;
                    if (b == null) continue;
                    b.BackColor = surface;
                    b.ForeColor = text;
                    b.FlatAppearance.BorderColor = hair;
                }
            }
            if (lblState != null) SetState(lblState.Text, lblState.ForeColor == Html(cur.Accent), progressFrac);
            btnTheme.Text = "THEME: " + t.Name;
            btnTheme.Location = new Point(pnlHeader.Width - btnTheme.Width - 12, 9);
        }

        // --------------------------------------------- barra titolo dark (DWM)

        [DllImport("dwmapi.dll")]
        static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int val, int size);

        // testo leggibile sopra un colore accent (scuro su accent chiari)
        static Color OnAccent(Color acc)
        {
            bool light = (acc.R * 299 + acc.G * 587 + acc.B * 114) / 1000 > 140;
            return light ? Color.FromArgb(20, 24, 32) : Color.White;
        }

        static bool IsDark(Color c)
        {
            return (0.299 * c.R + 0.587 * c.G + 0.114 * c.B) < 128;
        }

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);

        // scrollbar scure della lista nei temi scuri
        void ThemeScrollbars()
        {
            try
            {
                if (lvFiles == null || !lvFiles.IsHandleCreated) return;
                SetWindowTheme(lvFiles.Handle, IsDark(Html(cur.Bg)) ? "DarkMode_Explorer" : "Explorer", null);
                if (lstLog != null && lstLog.IsHandleCreated)
                    SetWindowTheme(lstLog.Handle, IsDark(Html(cur.Bg)) ? "DarkMode_Explorer" : "Explorer", null);
            }
            catch { }
        }

        static int ColorRef(Color c) { return c.R | (c.G << 8) | (c.B << 16); }

        void SetDarkTitle(bool dark)
        {
            try
            {
                if (!IsHandleCreated) return;
                int v = dark ? 1 : 0;
                // 20 = DWMWA_USE_IMMERSIVE_DARK_MODE (Win10 2004+/Win11);
                // su alcune build precedenti l'attributo valido e' 19
                if (DwmSetWindowAttribute(Handle, 20, ref v, 4) != 0)
                    DwmSetWindowAttribute(Handle, 19, ref v, 4);
                // Win11 (build 22000+): colori della barra del titolo dal tema;
                // sulle build precedenti le chiamate falliscono e restano ignorate
                int cap = ColorRef(Html(cur.Bg));
                int txt = ColorRef(Html(cur.Text));
                int brd = ColorRef(Html(cur.Hairline));
                DwmSetWindowAttribute(Handle, 35, ref cap, 4); // DWMWA_CAPTION_COLOR
                DwmSetWindowAttribute(Handle, 36, ref txt, 4); // DWMWA_TEXT_COLOR
                DwmSetWindowAttribute(Handle, 34, ref brd, 4); // DWMWA_BORDER_COLOR
            }
            catch { }
        }

        void CycleTheme()
        {
            int i = 0;
            for (int k = 0; k < Themes.Length; k++)
                if (Themes[k].Key == cur.Key) { i = k; break; }
            ApplyTheme(Themes[(i + 1) % Themes.Length]);
        }

        Label SectionLabel(string text)
        {
            Label l = new Label();
            l.Text = text;
            l.AutoSize = true;
            l.Font = new Font("Consolas", 11f, FontStyle.Bold);
            l.ForeColor = Html(cur.Muted);
            l.BackColor = Color.Transparent;
            return l;
        }

        Button MakeButton(string text, EventHandler onClick, bool accent)
        {
            Button b = new Button();
            b.Text = text;
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 1;
            b.Height = 40;
            b.AutoSize = true;
            b.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            b.Padding = new Padding(10, 6, 10, 6);
            b.Font = new Font("Consolas", 12f);
            b.Click += onClick;
            if (accent) b.Font = new Font("Consolas", 13f, FontStyle.Bold);
            return b;
        }

        void HairlinePaint(object s, PaintEventArgs e)
        {
            Control c = (Control)s;
            using (Pen p = new Pen(Html(cur.Hairline)))
                e.Graphics.DrawRectangle(p, 0, 0, c.Width - 1, c.Height - 1);
        }

        // istanza singola: i file inviati da "Invia a" / tasto dx finiscono
        // nella coda dell'istanza già aperta
        const int WM_COPYDATA = 0x004A;

        [StructLayout(LayoutKind.Sequential)]
        struct COPYDATASTRUCT
        {
            public IntPtr dwData;
            public int cbData;
            public IntPtr lpData;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern IntPtr FindWindow(string cls, string title);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessage")]
        static extern IntPtr SendMessageCD(IntPtr hwnd, uint msg, IntPtr w, ref COPYDATASTRUCT cd);

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_COPYDATA)
            {
                COPYDATASTRUCT cd = (COPYDATASTRUCT)m.GetLParam(typeof(COPYDATASTRUCT));
                if (cd.lpData != IntPtr.Zero)
                {
                    string paths = Marshal.PtrToStringUni(cd.lpData);
                    BeginInvoke((MethodInvoker)delegate
                    {
                        if (paths != null && paths.Length > 0) AddPaths(paths.Split('\n'));
                        if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
                        Activate();
                    });
                }
                m.Result = (IntPtr)1;
                return;
            }
            base.WndProc(ref m);
        }

        const string AppTitle = "BATCHPRINTER // stampa in blocco";
        string[] startFiles;

        public MainForm(string[] args)
        {
            startFiles = args ?? new string[0];
            // Rimuove cartelle di rendering lasciate da esecuzioni precedenti
            try
            {
                if (Directory.Exists(TempRoot))
                    foreach (string d in Directory.GetDirectories(TempRoot))
                        try { Directory.Delete(d, true); } catch { }
            }
            catch { }

            cur = Themes[0];
            try { Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            Text = AppTitle;
            Font = new Font("Segoe UI", 12f);
            MinimumSize = new Size(1180, 630);
            StartPosition = FormStartPosition.CenterScreen;
            BuildUi();
            Load += delegate
            {
                try { split.SplitterDistance = (int)(split.Width * 0.55); } catch { }
                SetDarkTitle(IsDark(Html(cur.Bg)));
                ThemeScrollbars();
                UpdateStatusBar();
            };
            FormClosing += delegate { Cleanup(); };
            LoadAppSettings();
            LoadDefaultProfile();
            if (defaultPrinter.Length > 0 && cboPrinter.Items.Contains(defaultPrinter))
                cboPrinter.SelectedItem = defaultPrinter;
            if (startFiles.Length > 0) AddPaths(startFiles);
        }

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                BootTrace("avvio");
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += delegate(object s, System.Threading.ThreadExceptionEventArgs e)
                {
                    ReportUnhandled(e.Exception);
                };
                AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e)
                {
                    ReportUnhandled(e.ExceptionObject as Exception);
                };

                // istanza singola: il mutex resta trattenuto per tutta la vita
                // dell'app (l'abbandono alla chiusura e' normale); chi arriva
                // dopo inoltra i file alla finestra esistente
                bool createdNew;
                using (Mutex mx = new Mutex(true, "Local\\BatchPrinterSingleInstance", out createdNew))
                {
                    if (!createdNew)
                    {
                        // la prima istanza potrebbe non avere ancora la finestra:
                        // riprova per ~2 secondi
                        IntPtr other = IntPtr.Zero;
                        for (int i = 0; i < 20 && other == IntPtr.Zero; i++)
                        {
                            other = FindWindow(null, AppTitle);
                            if (other == IntPtr.Zero) System.Threading.Thread.Sleep(100);
                        }
                        if (other != IntPtr.Zero)
                        {
                            // anche senza file: il messaggio vuoto porta la finestra in primo piano
                            {
                                string joined = args == null ? "" : string.Join("\n", args);
                                IntPtr buf = Marshal.StringToHGlobalUni(joined);
                                COPYDATASTRUCT cd = new COPYDATASTRUCT();
                                cd.dwData = IntPtr.Zero;
                                cd.cbData = Encoding.Unicode.GetByteCount(joined) + 2;
                                cd.lpData = buf;
                                SendMessageCD(other, WM_COPYDATA, IntPtr.Zero, ref cd);
                                Marshal.FreeHGlobal(buf);
                                BootTrace(joined.Length > 0 ? "file inoltrati a istanza esistente" : "istanza gia' attiva: attivazione finestra");
                            }
                            return;
                        }
                        // finestra non trovata (chiusura in corso): prosegui
                    }

                    BootTrace("handler registrati, avvio UI");
                    Application.Run(new MainForm(args ?? new string[0]));
                    BootTrace("UI chiusa normalmente");
                }
            }
            catch (Exception ex)
            {
                ReportUnhandled(ex);
                throw;
            }
        }

        static void ReportUnhandled(Exception ex)
        {
            try
            {
                if (!Directory.Exists(ProfileDir)) Directory.CreateDirectory(ProfileDir);
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("=== " + DateTime.Now + " ===");
                if (ex != null)
                {
                    sb.AppendLine(ex.GetType().FullName + ": " + ex.Message);
                    sb.AppendLine(ex.StackTrace);
                    Exception inner = ex.InnerException;
                    while (inner != null)
                    {
                        sb.AppendLine("inner: " + inner.GetType().FullName + ": " + inner.Message);
                        inner = inner.InnerException;
                    }
                }
                File.AppendAllText(Path.Combine(ProfileDir, "errorlog.txt"), sb.ToString());
                MessageBox.Show(null,
                    "Errore inatteso (il dettaglio \u00e8 in " + Path.Combine(ProfileDir, "errorlog.txt") + "):\n\n" +
                    (ex != null ? ex.Message : "sconosciuto"),
                    "BatchPrinter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }

        static void BootTrace(string msg)
        {
            try
            {
                if (!Directory.Exists(ProfileDir)) Directory.CreateDirectory(ProfileDir);
                File.AppendAllText(Path.Combine(ProfileDir, "errorlog.txt"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  [boot] " + msg + Environment.NewLine);
            }
            catch { }
        }

        void Cleanup()
        {
            try
            {
                foreach (PrintDocument d in DocPool) d.Dispose();
                foreach (Image i in PrintPool) i.Dispose();
                foreach (Stream s in StreamPool) s.Dispose();
                DocPool.Clear();
                PrintPool.Clear();
                StreamPool.Clear();
            }
            catch { }
        }

        // ---------------------------------------------------------------- UI

        void BuildUi()
        {
            ClientSize = new Size(1240, 660);

            // ORDINE DI DOCK (le righe del screenshot: fill per primo, poi
            // header, poi impostazioni, poi barra pulsanti)
            split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.Orientation = Orientation.Horizontal;
            split.FixedPanel = FixedPanel.None;
            split.BorderStyle = BorderStyle.None;
            Controls.Add(split);

            pnlHeader = new Panel();
            pnlHeader.Height = 52;
            pnlHeader.Paint += delegate(object s, PaintEventArgs e)
            {
                using (Pen p = new Pen(Html(cur.Hairline)))
                    e.Graphics.DrawLine(p, 0, pnlHeader.Height - 1, pnlHeader.Width, pnlHeader.Height - 1);
            };

            pnlSide = new Panel();
            pnlSide.Width = 300;
            pnlSide.Paint += HairlinePaint;

            Panel pnlBottom = new Panel();
            pnlBottom.Height = 30;
            pnlBottom.Paint += delegate(object s, PaintEventArgs e)
            {
                using (Pen p = new Pen(Html(cur.Hairline)))
                    e.Graphics.DrawLine(p, 0, 0, pnlBottom.Width, 0);
            };

            // ordine di dock: colonna tessere (sinistra), poi header (in alto,
            // sopra la colonna), poi barra di stato (in basso)
            pnlHeader.Dock = DockStyle.Top;
            pnlSide.Dock = DockStyle.Left;
            pnlBottom.Dock = DockStyle.Bottom;
            Controls.Add(pnlSide);
            Controls.Add(pnlHeader);
            Controls.Add(pnlBottom);

            // --- header
            lblLed = new Label();
            lblLed.Bounds = new Rectangle(0, 0, 4, 52);
            lblLed.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
            pnlHeader.Controls.Add(lblLed);

            lblTitle = new Label();
            lblTitle.Text = "BATCHPRINTER";
            lblTitle.Font = new Font("Consolas", 14f, FontStyle.Bold);
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(16, 9);
            pnlHeader.Controls.Add(lblTitle);

            lblSub = new Label();
            lblSub.Text = "// STAMPA IN BLOCCO v2";
            lblSub.Font = new Font("Consolas", 11f);
            lblSub.AutoSize = true;
            // segue la larghezza reale del titolo (font grande)
            lblTitle.SizeChanged += delegate { lblSub.Location = new Point(lblTitle.Right + 12, 14); };
            lblSub.Location = new Point(178, 13);
            pnlHeader.Controls.Add(lblSub);

            // status live: coda, stampante, copie, opzioni
            lblStatus = new Label();
            lblStatus.Font = new Font("Consolas", 10.5f);
            lblStatus.AutoSize = true;
            lblStatus.ForeColor = Color.Gray;
            lblStatus.Location = new Point(400, 20);
            lblStatus.TextChanged += delegate { PlaceStatus(); };
            pnlHeader.Controls.Add(lblStatus);

            btnTheme = MakeButton("THEME: LIGHT", delegate { CycleTheme(); }, false);
            btnTheme.Font = new Font("Consolas", 10.5f);
            pnlHeader.Controls.Add(btnTheme);
            pnlHeader.Resize += delegate { PlaceThemeButton(); PlaceStatus(); };

            BuildSidePanel();
            BuildBottomBar(pnlBottom);

            // --- lista documenti (pannello destro alto)
            pnlQueue = new Panel();
            pnlQueue.Dock = DockStyle.Fill;
            pnlQueue.Paint += HairlinePaint;
            split.Panel1.Controls.Add(pnlQueue);

            // Fill PRIMO, poi strip: il layout dock elabora in ordine inverso
            lvFiles = new ListView();
            lvFiles.Dock = DockStyle.Fill;
            lvFiles.View = View.Details;
            lvFiles.OwnerDraw = true;
            lvFiles.FullRowSelect = true;
            lvFiles.MultiSelect = true;
            lvFiles.HideSelection = false;
            lvFiles.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            lvFiles.BorderStyle = BorderStyle.None;
            lvFiles.Font = new Font("Consolas", 12.5f);
            lvFiles.Columns.Add("#", 44);
            lvFiles.Columns.Add("NOME", 400);
            lvFiles.Columns.Add("TIPO", 110);
            lvFiles.Columns.Add("DIMENSIONE", 100);
            lvFiles.Columns.Add("STATO", 130);
            lvFiles.DrawColumnHeader += Lv_DrawColumnHeader;
            lvFiles.DrawItem += Lv_DrawItem;
            lvFiles.DrawSubItem += Lv_DrawSubItem;
            lvFiles.AllowDrop = true;
            lvFiles.DragEnter += delegate(object s, DragEventArgs e)
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
            };
            lvFiles.DragDrop += delegate(object s, DragEventArgs e)
            {
                string[] dropped = (string[])e.Data.GetData(DataFormats.FileDrop);
                AddPaths(dropped);
            };
            pnlQueue.Controls.Add(lvFiles);
            // colonne con larghezza gestita dal codice: NOME occupa lo spazio
            // rimasto (si usa Width, immune dalle scrollbar, piu' una guardia
            // anti-ricorsione) e il trascinamento manuale dell'intestazione e'
            // bloccato, cosi' non compare mai la barra orizzontale
            bool sizingCols = false;
            lvFiles.ColumnWidthChanging += delegate(object s, ColumnWidthChangingEventArgs e)
            {
                if (sizingCols) return;
                e.Cancel = true;
                e.NewWidth = lvFiles.Columns[e.ColumnIndex].Width;
            };
            lvFiles.Resize += delegate
            {
                if (sizingCols) return;
                sizingCols = true;
                int others = 44 + 110 + 100 + 130;
                int w = lvFiles.Width - others - SystemInformation.VerticalScrollBarWidth - 4;
                lvFiles.Columns[1].Width = Math.Max(w, 200);
                sizingCols = false;
            };

            FlowLayoutPanel strip = new FlowLayoutPanel();
            strip.Dock = DockStyle.Top;
            strip.Height = 46;
            strip.Padding = new Padding(8, 6, 8, 0);
            strip.WrapContents = false;
            pnlQueue.Controls.Add(strip);

            Label sec1 = SectionLabel("01 // LISTA DOCUMENTI");
            sec1.AutoSize = false;
            sec1.Width = 210;
            sec1.Height = 24;
            strip.Controls.Add(sec1);
            strip.Controls.Add(MakeButton("RIMUOVI", delegate { RemoveSelected(); }, false));
            strip.Controls.Add(MakeButton("SVUOTA", delegate { Files.Clear(); RefreshList(); }, false));
            strip.Controls.Add(MakeButton("SU", delegate { MoveItem(-1); }, false));
            strip.Controls.Add(MakeButton("GI\u00d9", delegate { MoveItem(1); }, false));
            strip.Controls.Add(MakeButton("SALVA PROFILO", delegate { SaveProfile(); }, false));
            strip.Controls.Add(MakeButton("CARICA PROFILO", delegate { LoadProfile(false); }, false));

            // --- log (stesso schema)
            pnlLog = new Panel();
            pnlLog.Dock = DockStyle.Fill;
            pnlLog.Paint += HairlinePaint;
            split.Panel2.Controls.Add(pnlLog);

            lstLog = new ListBox();
            lstLog.Dock = DockStyle.Fill;
            lstLog.Font = new Font("Consolas", 11.5f);
            lstLog.HorizontalScrollbar = true;
            lstLog.BorderStyle = BorderStyle.None;
            pnlLog.Controls.Add(lstLog);

            Label sec2 = SectionLabel("02 // LOG");
            sec2.Dock = DockStyle.Top;
            sec2.Height = 24;
            sec2.Padding = new Padding(8, 5, 0, 0);
            pnlLog.Controls.Add(sec2);

            Load += delegate
            {
                try { split.SplitterDistance = (int)(split.Width * 0.55); } catch { }
                // i controlli interni di NumericUpDown non esistono piu' (ora
                // e' una TextBox): niente da ricollore qui
                PlaceThemeButton();
            };

            ApplyTheme(cur);
            PlaceThemeButton();
            UpdateCount();
        }

        void PlaceThemeButton()
        {
            if (btnTheme == null || pnlHeader == null) return;
            btnTheme.Location = new Point(pnlHeader.Width - btnTheme.Width - 12, 9);
        }

        // ------------------------------------------------ colonna sinistra

        void BuildSidePanel()
        {
            Label sec0 = SectionLabel("00 // DOCUMENTI");
            sec0.Location = new Point(14, 8);
            pnlSide.Controls.Add(sec0);

            Button tFiles = MakeButton("AGGIUNGI FILE", delegate { AddFiles(); }, false);
            tFiles.Font = new Font("Consolas", 14f, FontStyle.Bold);
            tFiles.SetBounds(14, 34, 272, 54);
            pnlSide.Controls.Add(tFiles);

            Button tFolder = MakeButton("AGGIUNGI CARTELLA", delegate { AddFolder(); }, false);
            tFolder.Font = new Font("Consolas", 14f, FontStyle.Bold);
            tFolder.SetBounds(14, 92, 272, 54);
            pnlSide.Controls.Add(tFolder);

            Label secP = SectionLabel("02 // STAMPANTE");
            secP.Location = new Point(14, 154);
            pnlSide.Controls.Add(secP);

            cboPrinter = new ComboBox();
            cboPrinter.DropDownStyle = ComboBoxStyle.DropDownList;
            cboPrinter.FlatStyle = FlatStyle.Flat;
            cboPrinter.SetBounds(14, 180, 272, 26);
            foreach (string p in PrinterSettings.InstalledPrinters) cboPrinter.Items.Add(p);
            if (cboPrinter.Items.Count > 0) cboPrinter.SelectedIndex = 0;
            cboPrinter.SelectedIndexChanged += delegate { UpdateStatusBar(); };
            pnlSide.Controls.Add(cboPrinter);

            Label secO = SectionLabel("03 // OPZIONI");
            secO.Location = new Point(14, 212);
            pnlSide.Controls.Add(secO);

            Label lblCop = new Label();
            lblCop.Text = "COPIE";
            lblCop.Font = new Font("Consolas", 10.5f);
            lblCop.ForeColor = Html(cur.Muted);
            lblCop.SetBounds(14, 238, 70, 20);
            pnlSide.Controls.Add(lblCop);

            numCopies = new TextBox();
            numCopies.BorderStyle = BorderStyle.FixedSingle;
            numCopies.TextAlign = HorizontalAlignment.Center;
            numCopies.BackColor = Html(cur.Surface);
            numCopies.ForeColor = Html(cur.Text);
            numCopies.Text = "1";
            numCopies.SetBounds(92, 234, 80, 26);
            numCopies.TextChanged += delegate { UpdateStatusBar(); };
            pnlSide.Controls.Add(numCopies);

            // accanto alle copie: la colonna e' alta quanto la finestra e non c'e' una riga libera
            chkPreview = new CheckBox();
            chkPreview.Text = "Anteprima";
            chkPreview.Font = new Font("Consolas", 11f);
            chkPreview.SetBounds(182, 234, 108, 26);
            chkPreview.CheckedChanged += delegate { UpdateStatusBar(); };
            pnlSide.Controls.Add(chkPreview);

            chkDuplex = new CheckBox();
            chkDuplex.Text = "Fronte/retro";
            chkDuplex.Font = new Font("Consolas", 12f);
            chkDuplex.SetBounds(14, 264, 170, 26);
            chkDuplex.CheckedChanged += delegate { UpdateStatusBar(); };
            pnlSide.Controls.Add(chkDuplex);

            chkCollate = new CheckBox();
            chkCollate.Text = "Fascicola";
            chkCollate.Font = new Font("Consolas", 12f);
            chkCollate.SetBounds(14, 292, 170, 26);
            chkCollate.CheckedChanged += delegate { UpdateStatusBar(); };
            pnlSide.Controls.Add(chkCollate);

            Label lblRng = new Label();
            lblRng.Text = "PAGINE (1-3, 1-3,5; vuoto = tutte)";
            lblRng.Font = new Font("Consolas", 10.5f);
            lblRng.ForeColor = Html(cur.Muted);
            lblRng.SetBounds(14, 324, 272, 20);
            pnlSide.Controls.Add(lblRng);

            txtRange = new TextBox();
            txtRange.BorderStyle = BorderStyle.FixedSingle;
            txtRange.SetBounds(14, 346, 272, 26);
            txtRange.TextChanged += delegate { UpdateStatusBar(); };
            pnlSide.Controls.Add(txtRange);

            Button tSet = MakeButton("IMPOSTAZIONI", delegate { ShowSettings(); }, false);
            tSet.SetBounds(14, 380, 130, 36);
            pnlSide.Controls.Add(tSet);

            Button tProf = MakeButton("PROFILI", delegate { LoadProfile(false); }, false);
            tProf.SetBounds(156, 380, 130, 36);
            pnlSide.Controls.Add(tProf);

            btnPrint = MakeButton("\u25b6 STAMPA CODA", delegate { PrintQueue(); }, true);
            btnPrint.Font = new Font("Consolas", 15f, FontStyle.Bold);
            btnPrint.SetBounds(14, 422, 272, 60);
            pnlSide.Controls.Add(btnPrint);

            // box stato (stile Print Conductor: sotto i tessere)
            Panel pnlState = new Panel();
            pnlState.SetBounds(14, 488, 272, 116);
            pnlState.Paint += delegate(object s, PaintEventArgs e)
            {
                using (Pen p = new Pen(Html(cur.Hairline)))
                    e.Graphics.DrawRectangle(p, 0, 0, pnlState.Width - 1, pnlState.Height - 1);
            };
            pnlSide.Controls.Add(pnlState);

            lblState = new Label();
            lblState.Text = "PRONTO";
            lblState.Font = new Font("Consolas", 14f, FontStyle.Bold);
            lblState.SetBounds(12, 10, 248, 26);
            pnlState.Controls.Add(lblState);

            lblCount = new Label();
            lblCount.Text = "0 DOCUMENTI IN LISTA";
            lblCount.Font = new Font("Consolas", 10.5f);
            lblCount.SetBounds(12, 44, 248, 22);
            pnlState.Controls.Add(lblCount);

            pnlProg = new Panel();
            pnlProg.SetBounds(12, 76, 248, 10);
            pnlProg.Paint += delegate(object s, PaintEventArgs e)
            {
                using (Pen p = new Pen(Html(cur.Hairline)))
                    e.Graphics.DrawRectangle(p, 0, 0, pnlProg.Width - 1, pnlProg.Height - 1);
                if (progressFrac > 0)
                {
                    using (SolidBrush b = new SolidBrush(Html(cur.Accent)))
                        e.Graphics.FillRectangle(b, 1, 1,
                            Math.Max(0, (int)((pnlProg.Width - 2) * progressFrac)), pnlProg.Height - 2);
                }
            };
            pnlState.Controls.Add(pnlProg);
        }

        void BuildBottomBar(Panel pnlBottom)
        {
            Label lblHint = new Label();
            lblHint.Text = "trascina file o cartelle nella lista \u00b7 tasto destro sul file \u2192 Invia a \u2192 BatchPrinter";
            lblHint.Font = new Font("Consolas", 10.5f);
            lblHint.AutoSize = true;
            lblHint.Location = new Point(10, 6);
            pnlBottom.Controls.Add(lblHint);

            // credito autore
            lblCredit = new Label();
            lblCredit.Text = "FATTO DA MDEANGELIS";
            lblCredit.Font = new Font("Consolas", 10.5f, FontStyle.Bold);
            lblCredit.AutoSize = true;
            lblCredit.Location = new Point(pnlBottom.Width - 220, 6);
            pnlBottom.Controls.Add(lblCredit);
            pnlBottom.Resize += delegate
            {
                lblCredit.Location = new Point(pnlBottom.Width - lblCredit.Width - 12, 6);
            };
        }

        void SetState(string state, bool accent, double frac)
        {
            if (lblState == null) return;
            lblState.Text = state;
            lblState.ForeColor = accent ? Html(cur.Accent) : Html(cur.Text);
            progressFrac = frac;
            if (pnlProg != null) pnlProg.Invalidate();
        }

        void PlaceStatus()
        {
            if (lblStatus == null || pnlHeader == null) return;
            int right = (btnTheme != null) ? btnTheme.Left - 16 : pnlHeader.Width - 12;
            lblStatus.Location = new Point(right - lblStatus.Width, 20);
        }

        void Lv_DrawColumnHeader(object s, DrawListViewColumnHeaderEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(Html(cur.Surface)))
                e.Graphics.FillRectangle(b, e.Bounds);
            using (Pen p = new Pen(Html(cur.Hairline)))
            {
                e.Graphics.DrawLine(p, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
                e.Graphics.DrawLine(p, e.Bounds.Right - 1, e.Bounds.Top, e.Bounds.Right - 1, e.Bounds.Bottom - 1);
            }
            TextFormatFlags sf = TextFormatFlags.VerticalCenter | TextFormatFlags.Left;
            TextRenderer.DrawText(e.Graphics, e.Header.Text, new Font("Consolas", 11f, FontStyle.Bold),
                new Rectangle(e.Bounds.Left + 8, e.Bounds.Top, e.Bounds.Width - 8, e.Bounds.Height),
                Html(cur.Muted), sf);
        }

        Font rowFontBold;
        Font RowFontBold
        {
            get { return rowFontBold ?? (rowFontBold = new Font(lvFiles.Font, FontStyle.Bold)); }
        }

        void Lv_DrawItem(object s, DrawListViewItemEventArgs e)
        {
            // il fondo di riga lo disegna DrawSubItem colonna per colonna
            e.DrawDefault = false;
        }

        void Lv_DrawSubItem(object s, DrawListViewSubItemEventArgs e)
        {
            bool sel = e.Item.Selected;
            Color bgRow = (e.Item.Index % 2 == 0) ? Html(cur.Surface) : Html(cur.Bg);
            // per la colonna 0 Win32 passa i limiti dell'intera riga: se li si
            // riempie al passaggio del mouse si cancellano le altre colonne
            Rectangle cell = e.SubItem.Bounds;
            if (e.ColumnIndex == 0) cell.Width = lvFiles.Columns[0].Width;
            using (SolidBrush b = new SolidBrush(sel ? Html(cur.Accent) : bgRow))
                e.Graphics.FillRectangle(b, cell);
            Color textC = sel ? Color.White : Html(cur.Text);
            Color mutedC = sel ? Color.FromArgb(220, 230, 240) : Html(cur.Muted);
            string path = Files[e.Item.Index];
            string text = e.SubItem.Text;
            bool bold = e.ColumnIndex == 1;
            if (e.ColumnIndex == 0 || e.ColumnIndex == 2 || e.ColumnIndex == 3 || e.ColumnIndex == 4)
                text = "  " + text;
            TextRenderer.DrawText(e.Graphics, text,
                bold ? RowFontBold : lvFiles.Font,
                cell, e.ColumnIndex == 1 ? textC : mutedC,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            using (Pen p = new Pen(Html(cur.Hairline)))
                e.Graphics.DrawLine(p, cell.Left, cell.Top, cell.Right, cell.Top);
        }

        string TypeLabel(string file)
        {
            string ext = Path.GetExtension(file).ToLowerInvariant();
            switch (ext)
            {
                case ".pdf": return "PDF";
                case ".doc": case ".docx": case ".dot": case ".dotx": case ".odt": return "Word";
                case ".xls": case ".xlsx": case ".xlsm": case ".ods": case ".csv": return "Excel";
                case ".ppt": case ".pptx": case ".odp": return "PowerPoint";
                case ".jpg": case ".jpeg": return "JPEG";
                case ".png": return "PNG";
                case ".gif": return "GIF";
                case ".bmp": return "BMP";
                case ".tif": case ".tiff": return "TIFF";
                case ".txt": case ".log": case ".rtf": return "Testo";
                case ".htm": case ".html": case ".mht": return "HTML";
                case ".xps": case ".oxps": return "XPS";
                case ".eml": case ".msg": return "E-mail";
                default: return ext.Trim('.').ToUpperInvariant();
            }
        }

        string SizeLabel(string file)
        {
            try
            {
                double len = new FileInfo(file).Length;
                if (len < 1024) return len + " B";
                if (len < 1024 * 1024) return (len / 1024).ToString("0") + " KB";
                return (len / 1024 / 1024).ToString("0.0") + " MB";
            }
            catch { return "?"; }
        }

        void RefreshList()
        {
            if (lvFiles == null) return;
            lvFiles.BeginUpdate();
            lvFiles.Items.Clear();
            for (int i = 0; i < Files.Count; i++)
            {
                string f = Files[i];
                ListViewItem it = new ListViewItem((i + 1).ToString());
                it.SubItems.Add(Path.GetFileName(f));
                it.SubItems.Add(TypeLabel(f));
                it.SubItems.Add(SizeLabel(f));
                it.SubItems.Add("");
                lvFiles.Items.Add(it);
            }
            lvFiles.EndUpdate();
            UpdateCount();
        }

        void RemoveSelected()
        {
            if (lvFiles == null) return;
            List<int> idx = new List<int>();
            foreach (int i in lvFiles.SelectedIndices) idx.Add(i);
            idx.Sort((a, b) => b.CompareTo(a));
            foreach (int i in idx) Files.RemoveAt(i);
            RefreshList();
        }

        void UpdateCount()
        {
            if (lblCount != null)
                lblCount.Text = Files.Count == 1 ? "1 DOCUMENTO IN LISTA" :
                    Files.Count + " DOCUMENTI IN LISTA";
            UpdateStatusBar();
        }

        // ------------------------------------------------- status bar (header)

        Label lblStatus;

        void UpdateStatusBar()
        {
            if (lblStatus == null) return;
            int files = Files.Count;
            string printer = cboPrinter != null && cboPrinter.Text.Length > 0 ? cboPrinter.Text : "nessuna";
            string copies = numCopies != null ? Copies().ToString() : "1";
            string extra = "";
            if (chkDuplex != null && chkDuplex.Checked) extra += " · fronte/retro";
            if (chkCollate != null && chkCollate.Checked) extra += " · fascicola";
            if (chkPreview != null && chkPreview.Checked) extra += " · anteprima";
            string range = txtRange != null ? txtRange.Text.Trim() : "";
            if (range.Length > 0) extra += " · pag. " + range;
            lblStatus.Text = "CODA " + files + " · " + printer + " · x" + copies + extra;
            PlaceStatus();
        }

        // ------------------------------------------------------------ coda

        int Copies()
        {
            int n;
            if (numCopies != null && int.TryParse(numCopies.Text.Trim(), out n) && n > 0)
                return Math.Min(n, 999);
            return 1;
        }

        void MoveItem(int delta)
        {
            if (lvFiles.SelectedIndices.Count == 0) return;
            int i = lvFiles.SelectedIndices[0];
            int j = i + delta;
            if (j < 0 || j >= Files.Count) return;
            string o = Files[i];
            Files[i] = Files[j];
            Files[j] = o;
            RefreshList();
            lvFiles.Items[j].Selected = true;
            lvFiles.EnsureVisible(j);
        }

        bool IsPrintable(string file)
        {
            string ext = Path.GetExtension(file).ToLowerInvariant();
            return ShellExts.Contains(ext) || ImageExts.Contains(ext);
        }

        void AddFiles()
        {
            using (OpenFileDialog d = new OpenFileDialog())
            {
                d.Title = "Seleziona file da stampare";
                d.Filter = "Tutti i file stampabili (*.pdf;*.doc*;*.rtf;*.xls*;*.ppt*;*.txt;*.csv;*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff)|" +
                    "*.pdf;*.doc;*.docx;*.dot;*.dotx;*.rtf;*.xls;*.xlsx;*.xlsm;*.ppt;*.pptx;*.txt;*.csv;*.xml;*.htm;*.html;*.xps;*.odt;*.ods;*.odp;*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.log|" +
                    "Tutti i file (*.*)|*.*";
                d.Multiselect = true;
                if (d.ShowDialog(this) == DialogResult.OK) AddPaths(d.FileNames);
            }
        }

        void AddFolder()
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "Cartella da scansionare (incluse le sottocartelle)";
                d.ShowNewFolderButton = false;
                if (d.ShowDialog(this) != DialogResult.OK) return;
                List<string> found = new List<string>();
                ScanFolder(d.SelectedPath, found);
                AddPaths(found.ToArray());
                Log("Cartella scansionata: " + found.Count + " file stampabili", true);
            }
        }

        // ricorsione manuale: una sottocartella senza permessi non deve
        // far fallire l'intera scansione
        void ScanFolder(string dir, List<string> found)
        {
            try
            {
                foreach (string f in Directory.GetFiles(dir))
                    if (IsPrintable(f)) found.Add(f);
                foreach (string sub in Directory.GetDirectories(dir))
                    ScanFolder(sub, found);
            }
            catch (UnauthorizedAccessException) { Log("Cartella non accessibile: " + dir, false); }
            catch (IOException) { Log("Cartella non leggibile: " + dir, false); }
        }

        void AddPaths(string[] paths)
        {
            foreach (string p in paths)
            {
                if (Directory.Exists(p))
                {
                    try { AddPaths(Directory.GetFiles(p)); }
                    catch (Exception ex) { Log("Cartella non leggibile: " + p + " - " + ex.Message, false); }
                    continue;
                }
                if (File.Exists(p) && IsPrintable(p) && !Files.Contains(p))
                    Files.Add(p);
            }
            RefreshList();
        }

        HashSet<int> ParseRange(string text)
        {
            HashSet<int> set = new HashSet<int>();
            if (text == null) return set;
            text = text.Trim();
            if (text.Length == 0) return set;
            foreach (string part in text.Split(','))
            {
                string s = part.Trim();
                if (s.Length == 0) continue;
                int dash = s.IndexOf('-');
                if (dash > 0)
                {
                    int a, b;
                    if (int.TryParse(s.Substring(0, dash).Trim(), out a) &&
                        int.TryParse(s.Substring(dash + 1).Trim(), out b))
                        for (int n = a; n <= b; n++) set.Add(n);
                }
                else
                {
                    int n;
                    if (int.TryParse(s, out n)) set.Add(n);
                }
            }
            return set;
        }

        // -------------------------------------------------------- stampa

        void PrintQueue()
        {
            if (Files.Count == 0)
            {
                MessageBox.Show(this, "La coda \u00e8 vuota.", "BatchPrinter");
                return;
            }
            if (cboPrinter.Text.Length == 0)
            {
                MessageBox.Show(this, "Nessuna stampante selezionata.", "BatchPrinter");
                return;
            }
            string[] files = Files.ToArray();
            if (asTaskOp == null) LoadAsTaskMethods();
            btnPrint.Enabled = false;
            previewOn = chkPreview.Checked;
            DateTime t0 = DateTime.Now;
            Log("Coda avviata: " + files.Length + " file", true);
            if (previewOn)
                foreach (string f in files)
                {
                    string x = Path.GetExtension(f).ToLowerInvariant();
                    if (x != ".pdf" && !ImageExts.Contains(x))
                    {
                        Log("Anteprima disponibile solo per PDF e immagini: gli altri file partono direttamente", false);
                        break;
                    }
                }
            SetState("STAMPA 0 DI " + files.Length, true, 0.0);
            Thread th = new Thread((ThreadStart)delegate
            {
                int done = 0, failed = 0;
                bool cancelled = false;
                try
                {
                    foreach (string f in files)
                    {
                        string st = PrintOne(f);
                        done++;
                        if (st == StErr) failed++;
                        int d = done;
                        BeginInvoke((MethodInvoker)delegate
                        {
                            SetState("STAMPA " + d + " DI " + files.Length, true,
                                (double)d / files.Length);
                            if (lvFiles != null && d - 1 < lvFiles.Items.Count)
                                lvFiles.Items[d - 1].SubItems[4].Text = st;
                            lvFiles.Invalidate();
                        });
                    }
                }
                catch (OperationCanceledException)
                {
                    cancelled = true;
                    Log("Coda annullata dall'utente", true);
                }
                catch (Exception ex)
                {
                    ReportUnhandled(ex);
                }
                TimeSpan dt = DateTime.Now - t0;
                Log("=== Coda " + (cancelled ? "annullata" : "completata") + ": " + done + "/" +
                    files.Length + " file, tempo " + dt.ToString("mm\\:ss") + " ===", true);
                BeginInvoke((MethodInvoker)delegate
                {
                    SetState((cancelled ? "ANNULLATO" : "COMPLETATO") + " \u00b7 " +
                        (done - failed) + "/" + files.Length, true, cancelled ? progressFrac : 1.0);
                    btnPrint.Enabled = true;
                    MessageBox.Show(this,
                        (cancelled ? "Coda annullata.\n\n" : "Lavoro completato.\n\n") +
                        done + " documenti su " + files.Length + " elaborati" +
                        (failed > 0 ? ", " + failed + " con errore (vedi il log)" : "") +
                        ".\nTempo impiegato: " + dt.ToString("mm\\:ss"),
                        "BatchPrinter", MessageBoxButtons.OK,
                        (failed > 0 || cancelled) ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                });
            });
            th.SetApartmentState(ApartmentState.MTA);
            th.IsBackground = true;
            th.Start();
        }

        // stati mostrati nella colonna STATO
        const string StOk = "OK";
        const string StErr = "ERRORE";
        const string StSkip = "SALTATO";

        string PrintOne(string file)
        {
            try
            {
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".pdf") return PrintPdfSilent(file);
                if (ImageExts.Contains(ext)) return PrintImage(file);
                return ShellPrint(file) ? StOk : StErr;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Log(file + "  ERRORE: " + ex.Message, true);
                return StErr;
            }
        }

        // Rendering PDF nativo in-process tramite Windows.Data.Pdf (WinRT):
        // nessun processo esterno, nessuna app associata richiesta.
        static MethodInfo asTaskOp;
        static MethodInfo asTaskAction;

        static void LoadAsTaskMethods()
        {
            Type ext = Type.GetType(
                "System.WindowsRuntimeSystemExtensions, System.Runtime.WindowsRuntime, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");
            foreach (MethodInfo m in ext.GetMethods())
            {
                if (m.Name == "AsTask" && m.GetParameters().Length == 1 &&
                    m.GetParameters()[0].ParameterType.Name == "IAsyncOperation`1")
                    asTaskOp = m;
                if (m.Name == "AsTask" && m.GetParameters().Length == 1 &&
                    m.GetParameters()[0].ParameterType.Name == "IAsyncAction")
                    asTaskAction = m;
            }
            if (asTaskOp == null || asTaskAction == null)
                throw new Exception("proiezioni WinRT non disponibili");
        }

        static object AwaitOp(object op, Type resultType)
        {
            MethodInfo mi = asTaskOp.MakeGenericMethod(resultType);
            Task t = (Task)mi.Invoke(null, new object[] { op });
            t.Wait();
            return t.GetType().GetProperty("Result").GetValue(t, null);
        }

        static void AwaitAct(object action)
        {
            Task t = (Task)asTaskAction.Invoke(null, new object[] { action });
            t.Wait();
        }

        string PrintPdfSilent(string file)
        {
            BootTrace("PDF: inizio rendering in-process " + file);
            string outDir = Path.Combine(TempRoot, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(outDir);
            List<byte[]> pageData;
            try { pageData = RenderPdfPages(file, outDir); }
            finally { try { Directory.Delete(outDir, true); } catch { } }
            BootTrace("PDF: rendering completato, " + pageData.Count + " pagine");
            string st = PrintSheets(file, pageData.Count, delegate(int i) { return PoolImage(pageData[i]); });
            if (st == StOk)
                Log(file + "  stampata (" + pageData.Count + " pagine, rendering nativo)", true);
            return st;
        }

        List<byte[]> RenderPdfPages(string file, string outDir)
        {
            StorageFolder folder = (StorageFolder)AwaitOp(
                StorageFolder.GetFolderFromPathAsync(outDir), typeof(StorageFolder));
            StorageFile sf = (StorageFile)AwaitOp(
                StorageFile.GetFileFromPathAsync(file), typeof(StorageFile));
            PdfDocument pdf = (PdfDocument)AwaitOp(
                PdfDocument.LoadFromFileAsync(sf), typeof(PdfDocument));

            // warm-up: evita il bug del primo rendering vuoto di Windows.Data.Pdf
            StorageFile warm = (StorageFile)AwaitOp(
                folder.CreateFileAsync("warm.png", CreationCollisionOption.ReplaceExisting), typeof(StorageFile));
            IRandomAccessStream wst = (IRandomAccessStream)AwaitOp(
                warm.OpenAsync(FileAccessMode.ReadWrite), typeof(IRandomAccessStream));
            PdfPage warmPage = pdf.GetPage(0);
            AwaitAct(warmPage.RenderToStreamAsync(wst));
            wst.Dispose();
            warmPage.Dispose();

            List<byte[]> pageData = new List<byte[]>();
            for (uint i = 0; i < pdf.PageCount; i++)
            {
                PdfPage page = pdf.GetPage(i);
                StorageFile outf = (StorageFile)AwaitOp(
                    folder.CreateFileAsync(string.Format("page-{0:D4}.png", i),
                        CreationCollisionOption.ReplaceExisting), typeof(StorageFile));
                IRandomAccessStream st = (IRandomAccessStream)AwaitOp(
                    outf.OpenAsync(FileAccessMode.ReadWrite), typeof(IRandomAccessStream));
                PdfPageRenderOptions opts = new PdfPageRenderOptions();
                opts.DestinationWidth = 2480u;
                opts.BackgroundColor = Windows.UI.Colors.White;
                AwaitAct(page.RenderToStreamAsync(st, opts));
                st.Dispose();
                page.Dispose();
                pageData.Add(File.ReadAllBytes(Path.Combine(outDir, string.Format("page-{0:D4}.png", i))));
            }
            return pageData;
        }

        Image PoolImage(byte[] data)
        {
            MemoryStream ms = new MemoryStream(data);
            StreamPool.Add(ms);
            Image img = Image.FromStream(ms);
            PrintPool.Add(img);
            return img;
        }

        // disegna l'immagine al massimo ingrandimento nell'area stampabile
        // del foglio (non nei margini di 2,5 cm di MarginBounds), centrata e
        // senza deformarla
        static void DrawFit(PrintPageEventArgs e, Image img)
        {
            RectangleF pa = e.PageSettings.PrintableArea;
            Rectangle page = (pa.Width > 0 && pa.Height > 0) ? Rectangle.Round(pa) : e.PageBounds;
            float scale = Math.Min(
                page.Width / (float)Math.Max(img.Width, 1),
                page.Height / (float)Math.Max(img.Height, 1));
            int w = (int)(img.Width * scale);
            int h = (int)(img.Height * scale);
            e.Graphics.DrawImage(img, new Rectangle(
                page.X + (page.Width - w) / 2,
                page.Y + (page.Height - h) / 2, w, h));
        }

        // PrintDocument comune ad anteprima e stampa: stampa solo le pagine
        // di 'sel' (gia' filtrate per range), ognuna presa da getPage
        PrintDocument BuildDoc(string file, List<int> sel, Func<int, Image> getPage)
        {
            PrintDocument doc = new PrintDocument();
            DocPool.Add(doc);
            doc.PrinterSettings.PrinterName = cboPrinter.Text;
            doc.PrinterSettings.Copies = (short)Copies();
            doc.PrinterSettings.Collate = chkCollate.Checked;
            if (chkDuplex.Checked && doc.PrinterSettings.CanDuplex)
                doc.PrinterSettings.Duplex = Duplex.Vertical;
            doc.DocumentName = Path.GetFileName(file);
            int idx = 0;
            doc.BeginPrint += delegate { idx = 0; };
            doc.PrintPage += delegate(object s, PrintPageEventArgs e)
            {
                BootTrace("stampa: pagina " + (sel[idx] + 1));
                DrawFit(e, getPage(sel[idx]));
                idx++;
                e.HasMorePages = idx < sel.Count;
            };
            return doc;
        }

        // Stampa (con anteprima opzionale) le pagine di un documento PDF o
        // immagine. Ritorna lo stato per la colonna STATO; con "ANNULLA CODA"
        // nell'anteprima lancia OperationCanceledException.
        string PrintSheets(string file, int total, Func<int, Image> getPage)
        {
            HashSet<int> range = ParseRange(txtRange.Text);
            // pagine da stampare: quelle fuori range non devono produrre fogli bianchi
            List<int> sel = new List<int>();
            for (int p = 0; p < total; p++)
                if (range.Count == 0 || range.Contains(p + 1)) sel.Add(p);
            if (sel.Count == 0)
                throw new Exception("nessuna pagina nel range indicato (" + total + " pagine)");

            PrintDocument doc = BuildDoc(file, sel, getPage);
            if (previewOn)
            {
                DialogResult r = DialogResult.Yes;
                string name = Path.GetFileName(file);
                int pages = sel.Count;
                Invoke((MethodInvoker)delegate { r = ShowPreview(doc, name, pages); });
                if (r == DialogResult.Cancel) throw new OperationCanceledException();
                if (r == DialogResult.OK)
                {
                    // niente piu' anteprime per i documenti rimasti in coda
                    previewOn = false;
                    Log("Anteprima disattivata: stampa di tutti i documenti rimasti", false);
                }
                if (r == DialogResult.No)
                {
                    Log(file + "  saltata dall'utente", true);
                    return StSkip;
                }
            }

            string notes = "";
            if (chkDuplex.Checked && !doc.PrinterSettings.CanDuplex)
                notes += " (fronte/retro non supportato dalla stampante)";
            if (sel.Count < total) notes += " pagine fuori range ignorate";
            BootTrace("stampa: PrintDocument pronto, " + sel.Count + "/" + total + " pagine, stampante " + cboPrinter.Text);
            doc.Print();
            BootTrace("stampa: doc.Print() completato");
            if (notes.Length > 0) Log("  note:" + notes, false);
            return StOk;
        }

        Button ThemedButton(string text, EventHandler onClick, bool accent)
        {
            Button b = MakeButton(text, onClick, accent);
            Color ac = Html(cur.Accent);
            b.BackColor = accent ? ac : Html(cur.Surface);
            b.ForeColor = accent ? OnAccent(ac) : Html(cur.Text);
            b.FlatAppearance.BorderColor = accent ? ac : Html(cur.Hairline);
            return b;
        }

        // Anteprima di stampa a tema: mostra le pagine come usciranno (stesso
        // PrintDocument della stampa). Yes = stampa, OK = stampa tutto senza altre anteprime, No = salta, Cancel = annulla coda.
        DialogResult ShowPreview(PrintDocument doc, string name, int pages)
        {
            using (Form d = new Form())
            {
                d.Text = "BATCHPRINTER // ANTEPRIMA";
                d.Font = Font;
                d.StartPosition = FormStartPosition.CenterParent;
                d.ClientSize = new Size(880, 760);
                d.MinimumSize = new Size(700, 480);
                d.ShowInTaskbar = false;
                d.BackColor = Html(cur.Bg);
                d.ForeColor = Html(cur.Text);

                PrintPreviewControl pc = new PrintPreviewControl();
                pc.Dock = DockStyle.Fill;
                pc.BackColor = Html(cur.Surface);
                pc.UseAntiAlias = true;
                pc.AutoZoom = true;
                pc.Rows = 1;
                pc.Columns = 1;
                pc.Document = doc;
                d.Controls.Add(pc);

                Label lbl = new Label();
                lbl.Dock = DockStyle.Top;
                lbl.Height = 40;
                lbl.Padding = new Padding(12, 10, 0, 0);
                lbl.Font = new Font("Consolas", 12f, FontStyle.Bold);
                lbl.ForeColor = Html(cur.Text);
                lbl.BackColor = Html(cur.Bg);
                d.Controls.Add(lbl);

                FlowLayoutPanel bar = new FlowLayoutPanel();
                bar.Dock = DockStyle.Bottom;
                bar.Height = 58;
                bar.Padding = new Padding(10, 8, 0, 0);
                bar.BackColor = Html(cur.Bg);
                d.Controls.Add(bar);

                int cp = 0;
                Action show = delegate
                {
                    pc.StartPage = cp;
                    lbl.Text = "ANTEPRIMA · " + name + " · pagina " + (cp + 1) + " di " + pages;
                };
                bar.Controls.Add(ThemedButton("◀", delegate { if (cp > 0) { cp--; show(); } }, false));
                bar.Controls.Add(ThemedButton("▶", delegate { if (cp < pages - 1) { cp++; show(); } }, false));
                bar.Controls.Add(ThemedButton("STAMPA", delegate { d.DialogResult = DialogResult.Yes; }, true));
                bar.Controls.Add(ThemedButton("OK, STAMPA TUTTO", delegate { d.DialogResult = DialogResult.OK; }, true));
                bar.Controls.Add(ThemedButton("SALTA", delegate { d.DialogResult = DialogResult.No; }, false));
                bar.Controls.Add(ThemedButton("ANNULLA CODA", delegate { d.DialogResult = DialogResult.Cancel; }, false));
                show();
                return d.ShowDialog(this);
            }
        }

        string PrintImage(string file)
        {
            byte[] data = File.ReadAllBytes(file);
            Image img = PoolImage(data);
            FrameDimension dim = new FrameDimension(img.FrameDimensionsList[0]);
            int frames = img.GetFrameCount(dim);
            string st = PrintSheets(file, frames, delegate(int i)
            {
                img.SelectActiveFrame(dim, i);
                return img;
            });
            if (st == StOk) Log(file + "  stampata, " + Copies() + " copia/e", true);
            return st;
        }

        bool ShellPrint(string file)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = file;
            psi.UseShellExecute = true;
            psi.WorkingDirectory = Path.GetDirectoryName(file);
            try
            {
                psi.Verb = "printto";
                psi.Arguments = "\"" + cboPrinter.Text + "\"";
                Process.Start(psi);
                Log(file + "  inviata alla stampante tramite app associata", true);
                return true;
            }
            catch (System.ComponentModel.Win32Exception)
            {
                try
                {
                    psi.Verb = "print";
                    psi.Arguments = "";
                    Process.Start(psi);
                    Log(file + "  inviata con verbo 'print' (stampante predefinita dell'app)", true);
                    return true;
                }
                catch (Exception ex)
                {
                    Log(file + "  ERRORE: nessun verbo di stampa valido - " + ex.Message, true);
                    return false;
                }
            }
        }

        // ---------------------------------------------- pannello impostazioni

        string SettingsPath { get { return Path.Combine(ProfileDir, "impostazioni.txt"); } }

        void LoadAppSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return;
                foreach (string line in File.ReadAllLines(SettingsPath))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string k = line.Substring(0, eq);
                    string v = line.Substring(eq + 1);
                    if (k == "theme")
                        foreach (Theme t in Themes)
                            if (t.Key == v) { ApplyTheme(t); break; }
                    if (k == "default_printer") defaultPrinter = v;
                }
            }
            catch (Exception ex)
            {
                Log("impostazioni non lette: " + ex.Message, false);
            }
        }

        void ShowSettings()
        {
            Form d = new Form();
            d.Text = "BATCHPRINTER // IMPOSTAZIONI";
            d.Font = Font;
            d.FormBorderStyle = FormBorderStyle.FixedDialog;
            d.MaximizeBox = false;
            d.MinimizeBox = false;
            d.StartPosition = FormStartPosition.CenterParent;
            d.ClientSize = new Size(470, 466);
            d.BackColor = Html(cur.Bg);
            d.ForeColor = Html(cur.Text);
            d.Paint += delegate(object s, PaintEventArgs e)
            {
                using (Pen p = new Pen(Html(cur.Hairline)))
                    e.Graphics.DrawRectangle(p, 0, 0, d.Width - 1, d.Height - 1);
            };

            Label sec1 = SectionLabel("01 // TEMA");
            sec1.Location = new Point(14, 14);
            d.Controls.Add(sec1);

            string selTheme = cur.Key;
            List<Button> pills = new List<Button>();
            FlowLayoutPanel pillsPanel = new FlowLayoutPanel();
            pillsPanel.Location = new Point(14, 40);
            pillsPanel.Size = new Size(440, 150);
            pillsPanel.WrapContents = true;
            d.Controls.Add(pillsPanel);
            foreach (Theme t in Themes)
            {
                Theme tc = t;
                Button pb = MakeButton(t.Name, null, false);
                pb.Click += delegate
                {
                    selTheme = tc.Key;
                    foreach (Button other in pills)
                    {
                        Theme ot = (Theme)other.Tag;
                        other.BackColor = Html(ot.Surface);
                        other.ForeColor = Html(ot.Text);
                        other.FlatAppearance.BorderColor = Html(ot.Hairline);
                    }
                    pb.BackColor = Html(tc.Accent);
                    pb.ForeColor = OnAccent(Html(tc.Accent));
                    pb.FlatAppearance.BorderColor = Html(tc.Accent);
                };
                pb.Tag = t;
                pills.Add(pb);
                pillsPanel.Controls.Add(pb);
                if (t.Key == selTheme)
                {
                    pb.BackColor = Html(t.Accent);
                    pb.ForeColor = OnAccent(Html(t.Accent));
                    pb.FlatAppearance.BorderColor = Html(t.Accent);
                }
                else
                {
                    pb.BackColor = Html(t.Surface);
                    pb.ForeColor = Html(t.Text);
                    pb.FlatAppearance.BorderColor = Html(t.Hairline);
                }
            }

            Label sec2 = SectionLabel("02 // STAMPANTE PREDEFINITA");
            sec2.Location = new Point(14, 206);
            d.Controls.Add(sec2);

            ComboBox cmb = new ComboBox();
            cmb.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb.Location = new Point(14, 234);
            cmb.Width = 440;
            cmb.BackColor = Html(cur.Surface);
            cmb.ForeColor = Html(cur.Text);
            cmb.FlatStyle = FlatStyle.Flat;
            foreach (string p in PrinterSettings.InstalledPrinters) cmb.Items.Add(p);
            string selPrinter = defaultPrinter.Length > 0 ? defaultPrinter : cboPrinter.Text;
            if (cmb.Items.Contains(selPrinter)) cmb.SelectedItem = selPrinter;
            d.Controls.Add(cmb);

            Label sec3 = SectionLabel("03 // STAMPANTE ATTIVA");
            sec3.Location = new Point(14, 276);
            d.Controls.Add(sec3);

            Label lblActive = new Label();
            lblActive.Text = cboPrinter.Text.Length > 0 ? cboPrinter.Text : "(nessuna)";
            lblActive.Font = new Font("Consolas", 12f, FontStyle.Bold);
            lblActive.ForeColor = Html(cur.Accent);
            lblActive.Location = new Point(14, 298);
            lblActive.AutoSize = true;
            d.Controls.Add(lblActive);

            Label note = new Label();
            note.Text = "La stampante predefinita viene usata all'avvio; la stampante\nattiva si cambia anche dal pannello principale.";
            note.Font = new Font("Consolas", 11f);
            note.ForeColor = Html(cur.Muted);
            note.Location = new Point(14, 328);
            note.AutoSize = true;
            d.Controls.Add(note);

            Button btnSave = MakeButton("SALVA", delegate
            {
                defaultPrinter = cmb.Text;
                List<string> lines = new List<string>();
                lines.Add("theme=" + selTheme);
                lines.Add("default_printer=" + defaultPrinter);
                Directory.CreateDirectory(ProfileDir);
                File.WriteAllLines(SettingsPath, lines.ToArray());
                foreach (Theme t in Themes)
                    if (t.Key == selTheme) { ApplyTheme(t); break; }
                if (cboPrinter.Items.Contains(defaultPrinter))
                    cboPrinter.SelectedItem = defaultPrinter;
                Log("Impostazioni salvate: tema " + selTheme + ", predefinita " + defaultPrinter, true);
                d.Close();
            }, true);
            btnSave.Location = new Point(14, 414);
            d.Controls.Add(btnSave);

            Button btnCancel = MakeButton("ANNULLA", delegate { d.Close(); }, false);
            btnCancel.Location = new Point(14 + btnSave.GetPreferredSize(Size.Empty).Width + 16, 414);
            d.Controls.Add(btnCancel);

            d.ShowDialog(this);
        }

        // ------------------------------------------------------ log/profili

        void Log(string line, bool alsoFile)
        {
            if (InvokeRequired)
            {
                BeginInvoke((MethodInvoker)delegate { Log(line, alsoFile); });
                return;
            }
            lstLog.Items.Add(DateTime.Now.ToString("HH:mm:ss") + "  " + line);
            lstLog.TopIndex = lstLog.Items.Count - 1;
            if (alsoFile)
            {
                try
                {
                    Directory.CreateDirectory(ProfileDir);
                    File.AppendAllText(Path.Combine(ProfileDir, "printlog.txt"),
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + line + Environment.NewLine);
                }
                catch { }
            }
        }

        Dictionary<string, string> SnapshotSettings()
        {
            Dictionary<string, string> s = new Dictionary<string, string>();
            s["printer"] = cboPrinter.Text;
            s["copies"] = Copies().ToString();
            s["duplex"] = chkDuplex.Checked ? "1" : "0";
            s["collate"] = chkCollate.Checked ? "1" : "0";
            s["preview"] = chkPreview.Checked ? "1" : "0";
            s["range"] = txtRange.Text;
            s["theme"] = cur.Key;
            List<string> files = new List<string>();
            foreach (string o in Files) files.Add(o);
            s["files"] = string.Join("|", files.ToArray());
            return s;
        }

        void ApplySettings(Dictionary<string, string> s)
        {
            if (s.ContainsKey("printer") && cboPrinter.Items.Contains(s["printer"]))
                cboPrinter.SelectedItem = s["printer"];
            short c;
            if (s.ContainsKey("copies") && short.TryParse(s["copies"], out c))
                numCopies.Text = Math.Max(1, Math.Min(999, (int)c)).ToString();
            chkDuplex.Checked = s.ContainsKey("duplex") && s["duplex"] == "1";
            chkCollate.Checked = s.ContainsKey("collate") && s["collate"] == "1";
            chkPreview.Checked = s.ContainsKey("preview") && s["preview"] == "1";
            txtRange.Text = s.ContainsKey("range") ? s["range"] : "";
            Files.Clear();
            if (s.ContainsKey("files"))
                foreach (string f in s["files"].Split('|'))
                    if (f.Length > 0 && File.Exists(f)) Files.Add(f);
            RefreshList();
        }

        void SaveProfile()
        {
            using (SaveFileDialog d = new SaveFileDialog())
            {
                d.Title = "Salva profilo";
                d.InitialDirectory = ProfileDir;
                d.Filter = "Profili (*.bpr)|*.bpr";
                d.FileName = "profilo.bpr";
                if (d.ShowDialog(this) != DialogResult.OK) return;
                List<string> lines = new List<string>();
                foreach (KeyValuePair<string, string> kv in SnapshotSettings())
                    lines.Add(kv.Key + "=" + kv.Value);
                File.WriteAllLines(d.FileName, lines.ToArray());
                Log("Profilo salvato: " + d.FileName, false);
            }
        }

        void LoadProfile(bool silent)
        {
            string path = Path.Combine(ProfileDir, "profilo.bpr");
            if (!File.Exists(path))
            {
                if (!silent) MessageBox.Show(this, "Nessun profilo in " + ProfileDir, "BatchPrinter");
                return;
            }
            Dictionary<string, string> s = new Dictionary<string, string>();
            foreach (string line in File.ReadAllLines(path))
            {
                int eq = line.IndexOf('=');
                if (eq > 0)
                    s[line.Substring(0, eq)] = line.Substring(eq + 1);
            }
            ApplySettings(s);
            if (!silent) Log("Profilo caricato", true);
        }

        void LoadDefaultProfile()
        {
            LoadProfile(true);
        }
    }
}
