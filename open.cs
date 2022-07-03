using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

[assembly: AssemblyVersion("3.0.0.0")]
[assembly: AssemblyFileVersion ("3.0.0.0")]
[assembly: AssemblyInformationalVersion("3.0")]
[assembly: AssemblyTitle("")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyProduct("OPEN")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyCopyright("Copyright (c) 2022 m-owada.")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

class Program
{
    [STAThread]
    static void Main()
    {
        bool createdNew;
        var mutex = new Mutex(true, @"Global\OPEN_COPYRIGHT_2022_M-OWADA", out createdNew);
        try
        {
            if(!createdNew)
            {
                MessageBox.Show("複数起動はできません。");
                return;
            }
            Application.EnableVisualStyles();
            Application.ThreadException += (object sender, ThreadExceptionEventArgs e) =>
            {
                throw new Exception(e.Exception.Message);
            };
            var mainForm = new MainForm();
            Application.Run();
        }
        catch(Exception e)
        {
            MessageBox.Show(e.Message, e.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
            Application.Exit();
        }
        finally
        {
            mutex.ReleaseMutex();
            mutex.Close();
        }
    }
}

class MainForm : Form
{
    private ComboBox comboBox = new ComboBox();
    private Button button = new Button();
    private ListBox listBox = new ListBox();
    private NotifyIcon notifyIcon = new NotifyIcon();
    private HotKey hotKey;
    private Dictionary<string, List<string>> lookaheadFiles = new Dictionary<string, List<string>>();
    private readonly string windowsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
    
    public MainForm()
    {
        // フォーム
        this.Text = "OPEN";
        this.Size = new Size(220, 140);
        this.MinimumSize  = this.Size;
        this.Location = new Point(10, 10);
        this.StartPosition = FormStartPosition.Manual;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.FormClosing += OnFormClosing;
        this.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);
        
        // コンボボックス
        comboBox.KeyDown += KeyDownComboBox;
        comboBox.KeyPress += KeyPressComboBox;
        comboBox.Text = string.Empty;
        comboBox.Location = new Point(10, 10);
        comboBox.Size = new Size(140, 20);
        comboBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        SetComboBoxItems();
        this.Controls.Add(comboBox);
        
        // ボタン
        button.Click += ClickButton;
        button.EnabledChanged += EnabledChangedButton;
        button.Location = new Point(155, 10);
        button.Size = new Size(40, 20);
        button.Text = "設定";
        button.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        this.Controls.Add(button);
        
        // リストボックス
        listBox.DoubleClick += DoubleClickListBox;
        listBox.KeyDown += KeyDownListBox;
        listBox.KeyPress += KeyPressListBox;
        listBox.Leave += LeaveListBox;
        listBox.MouseDown += MouseDownListBox;
        listBox.Location = new Point(10, 40);
        listBox.Size = new Size(185, 52);
        listBox.MultiColumn = false;
        listBox.SelectionMode = SelectionMode.MultiExtended;
        listBox.IntegralHeight = false;
        listBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(listBox);
        
        // 設定ファイル
        var config = new Config();
        this.Location = new Point(config.X, config.Y);
        this.Size = new Size(config.Width, config.Height);
        this.TopMost = config.TopMost;
        
        // タスクトレイ
        if(config.TaskTray)
        {
            // メニュー
            var menu = new ContextMenuStrip();
            menu.Items.Add(SetMenuItem("&起動", ClickMenuOpen));
            menu.Items.Add(SetMenuItem("&設定", ClickMenuConfig));
            menu.Items.Add(SetMenuItem("&終了", ClickMenuClose));
            
            // タスクトレイ
            notifyIcon.Icon = this.Icon;
            notifyIcon.Visible = true;
            notifyIcon.Text = this.Text;
            notifyIcon.ContextMenuStrip = menu;
            notifyIcon.DoubleClick += DoubleClickIcon;
        }
        else
        {
            this.Show();
        }
        
        // ホットキー
        SetHotKey(config.HotKey);
        
        // 終了イベント
        Application.ApplicationExit += ApplicationExit;
        
        // ファイル先読み
        if(config.Lookahead)
        {
            SetLookaheadFiles();
        }
    }
    
    private void OnFormClosing(object sender, FormClosingEventArgs e)
    {
        if(this.WindowState == FormWindowState.Normal)
        {
            var config = new Config();
            config.X = this.Location.X;
            config.Y = this.Location.Y;
            config.Width = this.Width;
            config.Height = this.Height;
            config.Save();
        }
        if(notifyIcon.Visible)
        {
            if(e.CloseReason != CloseReason.ApplicationExitCall)
            {
                e.Cancel = true;
                this.Hide();
            }
        }
        else
        {
            Application.Exit();
        }
    }
    
    private void KeyDownComboBox(object sender, KeyEventArgs e)
    {
        if(e.KeyCode == Keys.Enter)
        {
            EnterComboBox();
        }
    }
    
    private async void EnterComboBox()
    {
        comboBox.Text = comboBox.Text.Replace(@"/", "").Replace(@"\", "");
        var config = new Config();
        if(config.AutoTrim) comboBox.Text = comboBox.Text.Trim();
        if(!String.IsNullOrWhiteSpace(comboBox.Text))
        {
            await GetFileList();
        }
    }
    
    private async Task GetFileList()
    {
        comboBox.Enabled = false;
        button.Enabled = false;
        listBox.DataSource = null;
        
        await Task.Run(() =>
        {
            var items = new Dictionary<string, List<string>>();
            if(lookaheadFiles.Count > 0)
            {
                items = GetLookaheadFiles(comboBox.Text);
            }
            else
            {
                items = GetEnumerateFiles(comboBox.Text);
            }
            if(items.Count > 0)
            {
                listBox.BeginUpdate();
                listBox.DisplayMember = "Key";
                listBox.ValueMember = "Value";
                listBox.DataSource = new BindingSource(items.OrderBy(i => i.Key), null);
                listBox.EndUpdate();
                SaveHistory(comboBox.Text);
                SetComboBoxItems();
                comboBox.Text = string.Empty;
            }
        });
        
        comboBox.Enabled = true;
        button.Enabled = true;
        if(listBox.Items.Count == 0)
        {
            comboBox.SelectionStart = 0;
            comboBox.SelectionLength = comboBox.Text.Length;
            comboBox.Select();
        }
        else if(listBox.Items.Count == 1)
        {
            OpenFiles();
            comboBox.Select();
        }
        else
        {
            listBox.Select();
        }
    }
    
    private Dictionary<string, List<string>> GetEnumerateFiles(string keyword)
    {
        var items = new Dictionary<string, List<string>>();
        var config = new Config();
        var paths = new List<string>();
        foreach(var path in config.Path.Where(p => !p.Disable && Directory.Exists(p.Value)).ToList())
        {
            paths.Add(path.Value);
        }
        var extensions = new List<string>();
        foreach(var editor in config.Editor.Where(e => !e.Disable && !String.IsNullOrWhiteSpace(e.Attribute)).ToList())
        {
            extensions.Add(editor.Attribute);
        }
        if(config.Parallel)
        {
            items = ParallelSearch(paths, extensions, keyword);
        }
        else
        {
            items = SerialSearch(paths, extensions, keyword);
        }
        return items;
    }
    
    private Dictionary<string, List<string>> ParallelSearch(List<string> paths, List<string> extensions, string keyword)
    {
        var items = new Dictionary<string, List<string>>();
        var items_all = new Dictionary<int, Dictionary<string, List<string>>>();
        var lockobj = new object();
        Parallel.For
        (
            0,
            paths.Count,
            () => new Dictionary<int, Dictionary<string, List<string>>>(),
            (i, loopState, items_wk) =>
            {
                items_wk[i] = new Dictionary<string, List<string>>();
                foreach(var file in EnumerateFilesAllDirectories(paths[i], keyword + "*", extensions))
                {
                    var name = Path.GetFileName(file);
                    if(!items_wk[i].ContainsKey(name))
                    {
                        items_wk[i][name] = new List<string>();
                    }
                    items_wk[i][name].Add(file);
                }
                return items_wk;
            },
            (items_wk) =>
            {
                lock(lockobj)
                {
                    foreach(var item_wk in items_wk)
                    {
                        if(!items_all.ContainsKey(item_wk.Key))
                        {
                            items_all[item_wk.Key] = item_wk.Value;
                        }
                    }
                }
            }
        );
        for(var i = 0; i < items_all.Count; i++)
        {
            foreach(var item_all in items_all[i])
            {
                if(!items.ContainsKey(item_all.Key))
                {
                    items[item_all.Key] = new List<string>();
                }
                items[item_all.Key].AddRange(item_all.Value);
            }
        }
        return items;
    }
    
    private Dictionary<string, List<string>> SerialSearch(List<string> paths, List<string> extensions, string keyword)
    {
        var items = new Dictionary<string, List<string>>();
        foreach(var path in paths)
        {
            foreach(var file in EnumerateFilesAllDirectories(path, keyword + "*", extensions))
            {
                var name = Path.GetFileName(file);
                if(!items.ContainsKey(name))
                {
                    items[name] = new List<string>();
                }
                items[name].Add(file);
            }
        }
        return items;
    }
    
    private IEnumerable<string> EnumerateFilesAllDirectories(string path, string searchPattern, List<string> extensions)
    {
        var files = Enumerable.Empty<string>();
        if(path.StartsWith(windowsFolder, true, null))
        {
            return files;
        }
        try
        {
            files = Directory.EnumerateFiles(path, searchPattern);
            if(extensions.Count > 0)
            {
                files = files.Where(file => extensions.Any(extension => file.EndsWith("." + extension, true, null)));
            }
        }
        catch(UnauthorizedAccessException)
        {
        }
        catch(PathTooLongException)
        {
        }
        try
        {
            foreach(var p in Directory.EnumerateDirectories(path))
            {
                files = files.Union(EnumerateFilesAllDirectories(p, searchPattern, extensions));
            }
        }
        catch(UnauthorizedAccessException)
        {
        }
        catch(PathTooLongException)
        {
        }
        return files;
    }
    
    private void SaveHistory(string word)
    {
        var config = new Config();
        foreach(var history in config.History.ToList())
        {
            if(word == history.Value)
            {
                config.History.Remove(history);
            }
        }
        for(var i = config.History.Count(); i >= config.Count; i--)
        {
            if(config.History.Count() > 0)
            {
                config.History.RemoveAt(0);
            }
        }
        config.AddHistory(word);
        config.Save();
    }
    
    private void KeyPressComboBox(object sender, KeyPressEventArgs e)
    {
        if(e.KeyChar == (char)Keys.Enter)
        {
            e.Handled = true;
        }
    }
    
    private void SetComboBoxItems()
    {
        var items = new List<string>();
        var config = new Config();
        foreach(var history in config.History)
        {
            items.Add(history.Value);
        }
        items.Add(string.Empty);
        items.Reverse();
        comboBox.DataSource = null;
        comboBox.Items.Add(new object());
        comboBox.Items.Clear();
        if(items.Count > 0)
        {
            comboBox.DataSource = items.ToArray();
        }
    }
    
    private void ClickButton(object sender, EventArgs e)
    {
        ShowSubForm();
    }
    
    private void ShowSubForm()
    {
        hotKey.Dispose();
        button.Enabled = false;
        var topmost = this.TopMost;
        this.TopMost = false;
        var subForm = new SubForm();
        subForm.ShowDialog();
        this.TopMost = topmost;
        button.Enabled = true;
        comboBox.Select();
        
        var config = new Config();
        this.TopMost = config.TopMost;
        if(subForm.IsChanged)
        {
            comboBox.Text = string.Empty;
            listBox.DataSource = null;
            lookaheadFiles = new Dictionary<string, List<string>>();
            GC.Collect();
            if(config.Lookahead)
            {
                SetLookaheadFiles();
            }
        }
        SetHotKey(config.HotKey);
    }
    
    private void EnabledChangedButton(object sender, EventArgs e)
    {
        if(notifyIcon.Visible)
        {
            notifyIcon.ContextMenuStrip.Items[1].Enabled = button.Enabled;
        }
    }
    
    private void DoubleClickListBox(object sender, EventArgs e)
    {
        OpenFiles();
    }
    
    private void KeyDownListBox(object sender, KeyEventArgs e)
    {
        if(e.KeyCode == Keys.Enter)
        {
            OpenFiles();
        }
        else if(e.KeyData == (Keys.Control | Keys.C))
        {
            CopyFiles();
        }
    }
    
    private void KeyPressListBox(object sender, KeyPressEventArgs e)
    {
        if(e.KeyChar == (char)Keys.Enter)
        {
            e.Handled = true;
        }
        else if((Control.ModifierKeys & Keys.Control) == Keys.Control)
        {
            e.Handled = true;
        }
    }
    
    private void LeaveListBox(object sender, EventArgs e)
    {
        listBox.Update();
    }
    
    private void MouseDownListBox(object sender, MouseEventArgs e)
    {
        if(e.Button == MouseButtons.Right)
        {
            listBox.ContextMenuStrip = new ContextMenuStrip();
            int index = listBox.IndexFromPoint(e.Location);
            if(index >= 0)
            {
                listBox.SelectedIndex = index;
                for(int i = 0; i < listBox.Items.Count; i++)
                {
                    if(index != i && listBox.GetSelected(i)) listBox.SetSelected(i, false);
                }
                listBox.ContextMenuStrip = SetListBoxMenu();
            }
        }
    }
    
    private ContextMenuStrip SetListBoxMenu()
    {
        var menu = new ContextMenuStrip();
        var item = listBox.SelectedItem;
        var files = ((KeyValuePair<string, List<string>>)item).Value.Distinct();
        if(files.Count() <= 50)
        {
            foreach(var path in files)
            {
                var menuItem = new ToolStripMenuItem();
                menuItem.Text = path;
                menuItem.Enabled = true;
                menuItem.Click += MenuClickListBox;
                menu.Items.Add(menuItem);
            }
        }
        else
        {
            var menuItem = new ToolStripMenuItem();
            menuItem.Text = "ファイル数が50個を超えています。";
            menuItem.Enabled = false;
            menu.Items.Add(menuItem);
        }
        return menu;
    }
    
    private void MenuClickListBox(object sender, EventArgs e)
    {
        var menuItem = (ToolStripMenuItem)sender;
        if(OpenFile(menuItem.Text))
        {
            WindowAutoMinimized();
        }
    }
    
    private void OpenFiles()
    {
        var error = false;
        foreach(var item in listBox.SelectedItems)
        {
            var file = (KeyValuePair<string, List<string>>)item;
            if(!OpenFile(file.Value.First()))
            {
                error = true;
            }
        }
        if(!error)
        {
            WindowAutoMinimized();
        }
    }
    
    private bool OpenFile(string path)
    {
        var result = false;
        listBox.Enabled = false;
        if(File.Exists(path))
        {
            var found = false;
            var config = new Config();
            foreach(var editor in config.Editor.Where(e => !e.Disable).ToList())
            {
                if(path.EndsWith("." + editor.Attribute, true, null))
                {
                    if(File.Exists(editor.Value))
                    {
                        Process.Start(editor.Value, path);
                    }
                    else
                    {
                        Process.Start(path);
                    }
                    found = true;
                }
            }
            if(!found)
            {
                Process.Start(path);
            }
            result = true;
        }
        else
        {
            MessageBox.Show("ファイル \"" + path + "\" が見つかりません。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        listBox.Enabled = true;
        listBox.Select();
        return result;
    }
    
    private void WindowAutoMinimized()
    {
        var config = new Config();
        if(config.Minimize)
        {
            this.WindowState = FormWindowState.Minimized;
        }
    }
    
    private void CopyFiles()
    {
        var files = new StringCollection();
        foreach(var item in listBox.SelectedItems)
        {
            var file = (KeyValuePair<string, List<string>>)item;
            files.Add(file.Value.First());
        }
        if(files.Count > 0)
        {
            Clipboard.SetFileDropList(files);
        }
    }
    
    private ToolStripMenuItem SetMenuItem(string text, EventHandler handler)
    {
        var menuItem = new ToolStripMenuItem();
        menuItem.Text = text;
        menuItem.Click += handler;
        return menuItem;
    }
    
    private void ClickMenuOpen(object sender, EventArgs e)
    {
        ShowForm();
    }
    
    private void ClickMenuConfig(object sender, EventArgs e)
    {
        ShowForm();
        ShowSubForm();
    }
    
    private void ClickMenuClose(object sender, EventArgs e)
    {
        if(MessageBox.Show("終了しますか？", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
        {
            notifyIcon.Dispose();
            Application.Exit();
        }
    }
    
    private void DoubleClickIcon(object sender, EventArgs e)
    {
        ShowForm();
    }
    
    private void ShowForm()
    {
        this.Show();
        if(this.WindowState == FormWindowState.Minimized)
        {
            this.WindowState = FormWindowState.Normal;
        }
        this.Activate();
        comboBox.Select();
    }
    
    private void SetHotKey(Keys keyData)
    {
        var modKeyCode = (keyData & Keys.Modifiers);
        var hotKeyCode = (keyData & Keys.KeyCode);
        hotKey = new HotKey(ConvertModKey(modKeyCode), hotKeyCode);
        hotKey.HotKeyPush += HotKeyPush;
    }
    
    private HotKey.ModKey ConvertModKey(Keys modKeyCode)
    {
        var key = HotKey.ModKey.None;
        if((modKeyCode & Keys.Shift) == Keys.Shift)
        {
            key |= HotKey.ModKey.Shift;
        }
        if((modKeyCode & Keys.Control) == Keys.Control)
        {
            key |= HotKey.ModKey.Control;
        }
        if((modKeyCode & Keys.Alt) == Keys.Alt)
        {
            key |= HotKey.ModKey.Alt;
        }
        return key;
    }
    
    private void HotKeyPush(object sender, EventArgs e)
    {
        ShowForm();
    }
    
    private void ApplicationExit(object sender, EventArgs e)
    {
        hotKey.Dispose();
        Application.ApplicationExit -= ApplicationExit;
    }
    
    private async void SetLookaheadFiles()
    {
        comboBox.Enabled = false;
        button.Enabled = false;
        await Task.Run(() =>
        {
            lookaheadFiles = GetEnumerateFiles(string.Empty);
        });
        comboBox.Enabled = true;
        button.Enabled = true;
        comboBox.Select();
    }
    
    private Dictionary<string, List<string>> GetLookaheadFiles(string keyword)
    {
        var items = new Dictionary<string, List<string>>();
        foreach(var file in lookaheadFiles)
        {
            if(IsWildcardMatch(file.Key, keyword + "*"))
            {
                items.Add(file.Key, file.Value);
            }
        }
        return items;
    }
    
    private bool IsWildcardMatch(string target, string keyword)
    {
        var regexPattern = Regex.Escape(keyword);
        regexPattern = regexPattern.Replace(@"\*", ".*");
        regexPattern = regexPattern.Replace(@"\?", ".*");
        regexPattern = "^" + regexPattern;
        return new Regex(regexPattern, RegexOptions.IgnoreCase).IsMatch(target);
    }
}

class SubForm : Form
{
    public bool IsChanged { get; private set; }
    
    private DataGridView dataGridView1 = new DataGridView();
    private BindingList<Columns1> dataSource1 = new BindingList<Columns1>();
    private TextBox textBoxPath = new TextBox();
    
    private DataGridView dataGridView2 = new DataGridView();
    private BindingList<Columns2> dataSource2 = new BindingList<Columns2>();
    private TextBox textBoxExtension = new TextBox();
    private TextBox textBoxEditor = new TextBox();
    
    private CheckBox checkBoxTopMost = new CheckBox();
    private CheckBox checkBoxAutoTrim = new CheckBox();
    private CheckBox checkBoxParallel = new CheckBox();
    private CheckBox checkBoxMinimize = new CheckBox();
    private CheckBox checkBoxLookahead = new CheckBox();
    private CheckBox checkBoxTaskTray = new CheckBox();
    private TextBox textBoxHotKey = new TextBox();
    
    private Keys hotKeyData = Keys.None;
    
    public SubForm()
    {
        // フォーム
        this.Text = "設定";
        this.Size = new Size(800, 700);
        this.MaximizeBox = false;
        this.MinimizeBox = true;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.Load += FormLoad;
        
        // ツールチップ
        var toolTip = new ToolTip();
        
        // グループボックス1
        var groupBox1 = new GroupBox();
        groupBox1.Location = new Point(10, 10);
        groupBox1.Size = new Size(775, 210);
        groupBox1.FlatStyle = FlatStyle.Standard;
        groupBox1.Text = "【検索フォルダ】";
        this.Controls.Add(groupBox1);
        
        // ラベル
        groupBox1.Controls.Add(CreateLabel(10, 20, "検索対象のフォルダを指定します。先頭に指定したフォルダから順番に検索されます。Windowsフォルダは指定しても検索されません。"));
        
        // データグリッドビュー1
        dataGridView1.Location = new Point(10, 40);
        dataGridView1.Size = new Size(755, 130);
        dataGridView1.DataSource = dataSource1;
        dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        dataGridView1.AllowUserToAddRows = false;
        dataGridView1.MultiSelect = false;
        dataGridView1.ReadOnly = true;
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dataGridView1.CellPainting += CellPaintingDataGridView1;
        dataGridView1.CellDoubleClick += CellDoubleClickDataGridView1;
        groupBox1.Controls.Add(dataGridView1);
        
        // ラベル
        groupBox1.Controls.Add(CreateLabel(10, 180, "フォルダ"));
        
        // テキストボックス（Path）
        textBoxPath.Location = new Point(55, 180);
        textBoxPath.Size = new Size(430, 20);
        textBoxPath.Text = string.Empty;
        groupBox1.Controls.Add(textBoxPath);
        
        // 参照ボタン1
        var buttonRef1 = new Button();
        buttonRef1.Location = new Point(495, 180);
        buttonRef1.Size = new Size(40, 20);
        buttonRef1.Text = "参照";
        buttonRef1.Click += ClickButtonRef1;
        toolTip.SetToolTip(buttonRef1, "検索フォルダの選択ダイアログを表示します。");
        groupBox1.Controls.Add(buttonRef1);
        
        // 追加ボタン1
        var buttonAdd1 = new Button();
        buttonAdd1.Location = new Point(540, 180);
        buttonAdd1.Size = new Size(40, 20);
        buttonAdd1.Text = "追加";
        buttonAdd1.Click += ClickButtonAdd1;
        toolTip.SetToolTip(buttonAdd1, "入力した情報を一覧に追加します。");
        groupBox1.Controls.Add(buttonAdd1);
        
        // 更新ボタン1
        var buttonMod1 = new Button();
        buttonMod1.Location = new Point(585, 180);
        buttonMod1.Size = new Size(40, 20);
        buttonMod1.Text = "更新";
        buttonMod1.Click += ClickButtonMod1;
        toolTip.SetToolTip(buttonMod1, "一覧の選択行を入力した情報で更新します。");
        groupBox1.Controls.Add(buttonMod1);
        
        // 削除ボタン1
        var buttonDel1 = new Button();
        buttonDel1.Location = new Point(630, 180);
        buttonDel1.Size = new Size(40, 20);
        buttonDel1.Text = "削除";
        buttonDel1.Click += ClickButtonDel1;
        toolTip.SetToolTip(buttonDel1, "一覧の選択行を削除します。");
        groupBox1.Controls.Add(buttonDel1);
        
        // 無効ボタン1
        var buttonDis1 = new Button();
        buttonDis1.Location = new Point(675, 180);
        buttonDis1.Size = new Size(40, 20);
        buttonDis1.Text = "無効";
        buttonDis1.Click += ClickButtonDis1;
        toolTip.SetToolTip(buttonDis1, "一覧の選択行を無効にします。既に無効の場合は有効にします。");
        groupBox1.Controls.Add(buttonDis1);
        
        // ↑ボタン1
        var buttonUp1 = new Button();
        buttonUp1.Location = new Point(720, 180);
        buttonUp1.Size = new Size(20, 20);
        buttonUp1.Text = "↑";
        buttonUp1.Click += ClickButtonUp1;
        toolTip.SetToolTip(buttonUp1, "一覧の選択行を上に移動します。");
        groupBox1.Controls.Add(buttonUp1);
        
        // ↓ボタン1
        var buttonDown1 = new Button();
        buttonDown1.Location = new Point(745, 180);
        buttonDown1.Size = new Size(20, 20);
        buttonDown1.Text = "↓";
        buttonDown1.Click += ClickButtonDown1;
        toolTip.SetToolTip(buttonDown1, "一覧の選択行を下に移動します。");
        groupBox1.Controls.Add(buttonDown1);
        
        // グループボックス2
        var groupBox2 = new GroupBox();
        groupBox2.Location = new Point(10, 230);
        groupBox2.Size = new Size(775, 210);
        groupBox2.FlatStyle = FlatStyle.Standard;
        groupBox2.Text = "【拡張子／エディタ指定】";
        this.Controls.Add(groupBox2);
        
        // ラベル
        groupBox2.Controls.Add(CreateLabel(10, 20, "検索対象のファイルの拡張子及び使用するエディタを指定します。エディタが指定されていない場合は既定のアプリケーションを使用します。"));
        
        // データグリッドビュー2
        dataGridView2.Location = new Point(10, 40);
        dataGridView2.Size = new Size(755, 130);
        dataGridView2.DataSource = dataSource2;
        dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        dataGridView2.AllowUserToAddRows = false;
        dataGridView2.MultiSelect = false;
        dataGridView2.ReadOnly = true;
        dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dataGridView2.CellPainting += CellPaintingDataGridView2;
        dataGridView2.CellDoubleClick += CellDoubleClickDataGridView2;
        groupBox2.Controls.Add(dataGridView2);
        
        // ラベル
        groupBox2.Controls.Add(CreateLabel(10, 180, "拡張子"));
        
        // テキストボックス（Extension）
        textBoxExtension.Location = new Point(55, 180);
        textBoxExtension.Size = new Size(60, 20);
        textBoxExtension.Text = string.Empty;
        groupBox2.Controls.Add(textBoxExtension);
        
        // ラベル
        groupBox2.Controls.Add(CreateLabel(125, 180, "エディタ"));
        
        // テキストボックス（Editor）
        textBoxEditor.Location = new Point(170, 180);
        textBoxEditor.Size = new Size(315, 20);
        textBoxEditor.Text = string.Empty;
        groupBox2.Controls.Add(textBoxEditor);
        
        // 参照ボタン2
        var buttonRef2 = new Button();
        buttonRef2.Location = new Point(495, 180);
        buttonRef2.Size = new Size(40, 20);
        buttonRef2.Text = "参照";
        buttonRef2.Click += ClickButtonRef2;
        toolTip.SetToolTip(buttonRef2, "使用するエディタの選択ダイアログを表示します。");
        groupBox2.Controls.Add(buttonRef2);
        
        // 追加ボタン2
        var buttonAdd2 = new Button();
        buttonAdd2.Location = new Point(540, 180);
        buttonAdd2.Size = new Size(40, 20);
        buttonAdd2.Text = "追加";
        buttonAdd2.Click += ClickButtonAdd2;
        toolTip.SetToolTip(buttonAdd2, "入力した情報を一覧に追加します。");
        groupBox2.Controls.Add(buttonAdd2);
        
        // 更新ボタン2
        var buttonMod2 = new Button();
        buttonMod2.Location = new Point(585, 180);
        buttonMod2.Size = new Size(40, 20);
        buttonMod2.Text = "更新";
        buttonMod2.Click += ClickButtonMod2;
        toolTip.SetToolTip(buttonMod2, "一覧の選択行を入力した情報で更新します。");
        groupBox2.Controls.Add(buttonMod2);
        
        // 削除ボタン2
        var buttonDel2 = new Button();
        buttonDel2.Location = new Point(630, 180);
        buttonDel2.Size = new Size(40, 20);
        buttonDel2.Text = "削除";
        buttonDel2.Click += ClickButtonDel2;
        toolTip.SetToolTip(buttonDel2, "一覧の選択行を削除します。");
        groupBox2.Controls.Add(buttonDel2);
        
        // 無効ボタン2
        var buttonDis2 = new Button();
        buttonDis2.Location = new Point(675, 180);
        buttonDis2.Size = new Size(40, 20);
        buttonDis2.Text = "無効";
        buttonDis2.Click += ClickButtonDis2;
        toolTip.SetToolTip(buttonDis2, "一覧の選択行を無効にします。既に無効の場合は有効にします。");
        groupBox2.Controls.Add(buttonDis2);
        
        // ↑ボタン2
        var buttonUp2 = new Button();
        buttonUp2.Location = new Point(720, 180);
        buttonUp2.Size = new Size(20, 20);
        buttonUp2.Text = "↑";
        buttonUp2.Click += ClickButtonUp2;
        toolTip.SetToolTip(buttonUp2, "一覧の選択行を上に移動します。");
        groupBox2.Controls.Add(buttonUp2);
        
        // ↓ボタン2
        var buttonDown2 = new Button();
        buttonDown2.Location = new Point(745, 180);
        buttonDown2.Size = new Size(20, 20);
        buttonDown2.Text = "↓";
        buttonDown2.Click += ClickButtonDown2;
        toolTip.SetToolTip(buttonDown2, "一覧の選択行を下に移動します。");
        groupBox2.Controls.Add(buttonDown2);
        
        // グループボックス3
        var groupBox3 = new GroupBox();
        groupBox3.Location = new Point(10, 450);
        groupBox3.Size = new Size(775, 170);
        groupBox3.FlatStyle = FlatStyle.Standard;
        groupBox3.Text = "【その他】";
        this.Controls.Add(groupBox3);
        
        // チェックボックス（Topmost）
        checkBoxTopMost.Text = "フォームを常に手前に表示する";
        checkBoxTopMost.Location = new Point(10, 20);
        checkBoxTopMost.AutoSize = true;
        checkBoxTopMost.Checked = false;
        groupBox3.Controls.Add(checkBoxTopMost);
        
        // チェックボックス（AutoTrim）
        checkBoxAutoTrim.Text = "フォームに入力したファイル名の前後に空白が含まれている場合は自動で削除する";
        checkBoxAutoTrim.Location = new Point(10, 45);
        checkBoxAutoTrim.AutoSize = true;
        checkBoxAutoTrim.Checked = false;
        groupBox3.Controls.Add(checkBoxAutoTrim);
        
        // チェックボックス（Minimize）
        checkBoxMinimize.Text = "ファイルを開くと同時に自動でフォームを最小化する";
        checkBoxMinimize.Location = new Point(10, 70);
        checkBoxMinimize.AutoSize = true;
        checkBoxMinimize.Checked = false;
        groupBox3.Controls.Add(checkBoxMinimize);
        
        // チェックボックス（Parallel）
        checkBoxParallel.Text = "ファイルの検索処理を検索フォルダごとに並列に処理する (処理速度が向上する場合があります)";
        checkBoxParallel.Location = new Point(10, 95);
        checkBoxParallel.AutoSize = true;
        checkBoxParallel.Checked = false;
        groupBox3.Controls.Add(checkBoxParallel);
        
        // チェックボックス（Lookahead）
        checkBoxLookahead.Text = "アプリケーション起動時に検索フォルダに格納されているファイルの一覧を先読みしてメモリに記憶する";
        checkBoxLookahead.Location = new Point(10, 120);
        checkBoxLookahead.AutoSize = true;
        checkBoxLookahead.Checked = false;
        checkBoxLookahead.CheckedChanged += CheckedChangedCheckBoxLookahead;
        groupBox3.Controls.Add(checkBoxLookahead);
        
        // チェックボックス（Tasktray）
        checkBoxTaskTray.Text = "アプリケーション起動時にタスクトレイに常駐する (次回アプリケーション起動時に有効となります)";
        checkBoxTaskTray.Location = new Point(10, 145);
        checkBoxTaskTray.AutoSize = true;
        checkBoxTaskTray.Checked = false;
        groupBox3.Controls.Add(checkBoxTaskTray);
        
        // ラベル（HotKey）
        var labelHotKey = new Label();
        labelHotKey.Location = new Point(20, 630);
        labelHotKey.Size = new Size(100, 20);
        labelHotKey.Text = "ショートカットキー";
        labelHotKey.BorderStyle = BorderStyle.Fixed3D;
        labelHotKey.TextAlign = ContentAlignment.MiddleCenter;
        this.Controls.Add(labelHotKey);
        
        // テキストボックス（HotKey）
        textBoxHotKey.Location = new Point(130, 630);
        textBoxHotKey.Size = new Size(160, 20);
        textBoxHotKey.Text = string.Empty;
        textBoxHotKey.TextAlign = HorizontalAlignment.Center;
        textBoxHotKey.ReadOnly = true;
        textBoxHotKey.BackColor = SystemColors.Window;
        textBoxHotKey.ShortcutsEnabled = false;
        textBoxHotKey.PreviewKeyDown += PreviewKeyDownTextBoxHotKey;
        this.Controls.Add(textBoxHotKey);
        
        // 保存ボタン
        var buttonSave = new Button();
        buttonSave.Location = new Point(650, 630);
        buttonSave.Size = new Size(60, 30);
        buttonSave.Text = "保存";
        buttonSave.Click += ClickButtonSave;
        toolTip.SetToolTip(buttonSave, "設定を保存して終了します。");
        this.Controls.Add(buttonSave);
        
        // 中止ボタン
        var buttonCancel  = new Button();
        buttonCancel.Location = new Point(720, 630);
        buttonCancel.Size = new Size(60, 30);
        buttonCancel.Text = "中止";
        buttonCancel.Click += ClickButtonCancel;
        toolTip.SetToolTip(buttonCancel, "設定を保存しないで終了します。");
        this.Controls.Add(buttonCancel);
        
        // データ読込
        LoadData();
    }
    
    private class Columns1
    {
        [DisplayName("フォルダ")]
        public string Path { get; set; }
        
        [Browsable(false)]
        public bool Disable { get; set; }
    }
    
    private class Columns2
    {
        [DisplayName("拡張子")]
        public string Extension { get; set; }
        
        [DisplayName("エディタ")]
        public string Editor { get; set; }
        
        [Browsable(false)]
        public bool Disable { get; set; }
    }
    
    private Label CreateLabel(int x, int y, string text)
    {
        var label = new Label();
        label.Location = new Point(x, y);
        label.Text = text;
        label.Size = new Size(label.PreferredWidth, 20);
        label.TextAlign = ContentAlignment.MiddleLeft;
        return label;
    }
    
    private void FormLoad(object sender, EventArgs e)
    {
        foreach(DataGridViewRow row in dataGridView1.Rows)
        {
            SetStrikeoutDataGridView1(row.Index);
        }
        foreach(DataGridViewRow row in dataGridView2.Rows)
        {
            SetStrikeoutDataGridView2(row.Index);
        }
    }
    
    private void SetStrikeoutDataGridView1(int index)
    {
        SetFontStyleStrikeout(dataGridView1, index, dataSource1[index].Disable);
    }
    
    private void SetStrikeoutDataGridView2(int index)
    {
        SetFontStyleStrikeout(dataGridView2, index, dataSource2[index].Disable);
    }
    
    private void SetFontStyleStrikeout(DataGridView dataGridView, int index, bool disable)
    {
        var style = FontStyle.Regular;
        if(disable)
        {
            style = FontStyle.Strikeout;
        }
        foreach(DataGridViewCell cell in dataGridView.Rows[index].Cells)
        {
            cell.Style.Font = new Font(dataGridView.DefaultCellStyle.Font, style);
        }
    }
    
    private void CellPaintingDataGridView1(object sender, DataGridViewCellPaintingEventArgs e)
    {
        if(e.RowIndex >= 0 && e.ColumnIndex < 0)
        {
            CellPaintingRowIndex(e);
        }
    }
    
    private void CellPaintingDataGridView2(object sender, DataGridViewCellPaintingEventArgs e)
    {
        if(e.RowIndex >= 0 && e.ColumnIndex < 0)
        {
            CellPaintingRowIndex(e);
        }
        if(e.RowIndex >= 0 && e.ColumnIndex == 1)
        {
            CellPaintingPlaceholderText(e, "(既定のアプリケーション)");
        }
    }
    
    private void CellPaintingRowIndex(DataGridViewCellPaintingEventArgs e)
    {
        e.Paint(e.ClipBounds, DataGridViewPaintParts.All);
        var indexRect = e.CellBounds;
        indexRect.Inflate(-2, -2);
        TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), e.CellStyle.Font, indexRect, e.CellStyle.ForeColor, TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
        e.Handled = true;
    }
    
    private void CellPaintingPlaceholderText(DataGridViewCellPaintingEventArgs e, string text)
    {
        if(e.Value != null)
        {
            if(e.Value.ToString() == "")
            {
                e.Paint(e.ClipBounds, DataGridViewPaintParts.All & ~(DataGridViewPaintParts.ContentForeground));
                TextRenderer.DrawText(e.Graphics, text, e.CellStyle.Font, e.CellBounds, SystemColors.GrayText, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
                e.Handled = true;
            }
        }
    }
    
    private void CellDoubleClickDataGridView1(object sender, DataGridViewCellEventArgs e)
    {
        if(dataGridView1.CurrentCell != null)
        {
            var row = dataSource1[dataGridView1.CurrentCell.RowIndex];
            textBoxPath.Text = row.Path;
        }
    }
    
    private void ClickButtonRef1(object sender, EventArgs e)
    {
        using(var dialog = new FolderBrowserDialog())
        {
            dialog.Description = "検索フォルダを指定してください。";
            dialog.SelectedPath = textBoxPath.Text;
            dialog.ShowNewFolderButton = false;
            if(dialog.ShowDialog() == DialogResult.OK)
            {
                textBoxPath.Text = dialog.SelectedPath;
            }
        }
    }
    
    private bool CheckDataGridView1()
    {
        if(String.IsNullOrWhiteSpace(textBoxPath.Text))
        {
            MessageBox.Show("検索フォルダを入力してください。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            textBoxPath.Select();
            return false;
        }
        if(dataSource1.Count >= 100)
        {
            MessageBox.Show("指定できる件数の上限を超過しました。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        return true;
    }
    
    private void ClickButtonAdd1(object sender, EventArgs e)
    {
        if(CheckDataGridView1())
        {
            var row = new Columns1();
            row.Path = textBoxPath.Text;
            dataSource1.Add(row);
            dataGridView1.CurrentCell = dataGridView1[0, dataGridView1.RowCount - 1];
            textBoxPath.Text = string.Empty;
            textBoxPath.Select();
            IsChanged = true;
        }
    }
    
    private void ClickButtonMod1(object sender, EventArgs e)
    {
        if(dataGridView1.CurrentCell != null && CheckDataGridView1())
        {
            var row = dataSource1[dataGridView1.CurrentCell.RowIndex];
            row.Path = textBoxPath.Text;
            dataGridView1.Refresh();
            textBoxPath.Text = string.Empty;
            IsChanged = true;
        }
    }
    
    private void ClickButtonDel1(object sender, EventArgs e)
    {
        if(dataGridView1.CurrentCell != null)
        {
            var index = dataGridView1.CurrentCell.RowIndex;
            if(dataGridView1.CurrentCell.RowIndex > 0)
            {
                dataGridView1.CurrentCell = dataGridView1[0, dataGridView1.CurrentCell.RowIndex - 1];
            }
            dataSource1.RemoveAt(index);
            IsChanged = true;
        }
    }
    
    private void ClickButtonDis1(object sender, EventArgs e)
    {
        if(dataGridView1.CurrentCell != null)
        {
            var row = dataSource1[dataGridView1.CurrentCell.RowIndex];
            row.Disable = !row.Disable;
            SetStrikeoutDataGridView1(dataGridView1.CurrentCell.RowIndex);
            IsChanged = true;
        }
    }
    
    private void ClickButtonUp1(object sender, EventArgs e)
    {
        if(dataGridView1.CurrentCell != null)
        {
            if(dataGridView1.CurrentCell.RowIndex > 0)
            {
                var row = dataSource1[dataGridView1.CurrentCell.RowIndex];
                dataSource1[dataGridView1.CurrentCell.RowIndex] = dataSource1[dataGridView1.CurrentCell.RowIndex - 1];
                dataSource1[dataGridView1.CurrentCell.RowIndex - 1] = row;
                SetStrikeoutDataGridView1(dataGridView1.CurrentCell.RowIndex);
                SetStrikeoutDataGridView1(dataGridView1.CurrentCell.RowIndex - 1);
                dataGridView1.CurrentCell = dataGridView1[0, dataGridView1.CurrentCell.RowIndex - 1];
                IsChanged = true;
            }
        }
    }
    
    private void ClickButtonDown1(object sender, EventArgs e)
    {
        if(dataGridView1.CurrentCell != null)
        {
            if(dataGridView1.CurrentCell.RowIndex < dataGridView1.RowCount - 1)
            {
                var row = dataSource1[dataGridView1.CurrentCell.RowIndex];
                dataSource1[dataGridView1.CurrentCell.RowIndex] = dataSource1[dataGridView1.CurrentCell.RowIndex + 1];
                dataSource1[dataGridView1.CurrentCell.RowIndex + 1] = row;
                SetStrikeoutDataGridView1(dataGridView1.CurrentCell.RowIndex);
                SetStrikeoutDataGridView1(dataGridView1.CurrentCell.RowIndex + 1);
                dataGridView1.CurrentCell = dataGridView1[0, dataGridView1.CurrentCell.RowIndex + 1];
                IsChanged = true;
            }
        }
    }
    
    private void CellDoubleClickDataGridView2(object sender, DataGridViewCellEventArgs e)
    {
        if(dataGridView2.CurrentCell != null)
        {
            var row = dataSource2[dataGridView2.CurrentCell.RowIndex];
            textBoxExtension.Text = row.Extension;
            textBoxEditor.Text = row.Editor;
        }
    }
    
    private void ClickButtonRef2(object sender, EventArgs e)
    {
        using(var dialog = new OpenFileDialog())
        {
            dialog.Title = "使用するアプリケーションを指定してください。";
            dialog.InitialDirectory = textBoxEditor.Text;
            dialog.Filter = "すべてのファイル (*.*)|*.*|アプリケーション (*.exe)|*.exe";
            dialog.FilterIndex = 2;
            if(dialog.ShowDialog() == DialogResult.OK)
            {
                textBoxEditor.Text = dialog.FileName;
            }
        }
    }
    
    private bool CheckDataGridView2()
    {
        if(String.IsNullOrWhiteSpace(textBoxExtension.Text))
        {
            MessageBox.Show("検索対象のファイルの拡張子を入力してください。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            textBoxExtension.Select();
            return false;
        }
        if(dataSource2.Count >= 100)
        {
            MessageBox.Show("指定できる件数の上限を超過しました。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        return true;
    }
    
    private void ClickButtonAdd2(object sender, EventArgs e)
    {
        if(CheckDataGridView2())
        {
            var row = new Columns2();
            row.Extension = textBoxExtension.Text;
            row.Editor = textBoxEditor.Text;
            dataSource2.Add(row);
            dataGridView2.CurrentCell = dataGridView2[0, dataGridView2.RowCount - 1];
            textBoxExtension.Text = string.Empty;
            textBoxEditor.Text = string.Empty;
            textBoxExtension.Select();
            IsChanged = true;
        }
    }
    
    private void ClickButtonMod2(object sender, EventArgs e)
    {
        if(dataGridView2.CurrentCell != null && CheckDataGridView2())
        {
            var row = dataSource2[dataGridView2.CurrentCell.RowIndex];
            row.Extension = textBoxExtension.Text;
            row.Editor = textBoxEditor.Text;
            dataGridView2.Refresh();
            textBoxExtension.Text = string.Empty;
            textBoxEditor.Text = string.Empty;
            IsChanged = true;
        }
    }
    
    private void ClickButtonDel2(object sender, EventArgs e)
    {
        if(dataGridView2.CurrentCell != null)
        {
            var index = dataGridView2.CurrentCell.RowIndex;
            if(dataGridView2.CurrentCell.RowIndex > 0)
            {
                dataGridView2.CurrentCell = dataGridView2[0, dataGridView2.CurrentCell.RowIndex - 1];
            }
            dataSource2.RemoveAt(index);
            IsChanged = true;
        }
    }
    
    private void ClickButtonDis2(object sender, EventArgs e)
    {
        if(dataGridView2.CurrentCell != null)
        {
            var row = dataSource2[dataGridView2.CurrentCell.RowIndex];
            row.Disable = !row.Disable;
            SetStrikeoutDataGridView2(dataGridView2.CurrentCell.RowIndex);
            IsChanged = true;
        }
    }
    
    private void ClickButtonUp2(object sender, EventArgs e)
    {
        if(dataGridView2.CurrentCell != null)
        {
            if(dataGridView2.CurrentCell.RowIndex > 0)
            {
                var row = dataSource2[dataGridView2.CurrentCell.RowIndex];
                dataSource2[dataGridView2.CurrentCell.RowIndex] = dataSource2[dataGridView2.CurrentCell.RowIndex - 1];
                dataSource2[dataGridView2.CurrentCell.RowIndex - 1] = row;
                SetStrikeoutDataGridView2(dataGridView2.CurrentCell.RowIndex);
                SetStrikeoutDataGridView2(dataGridView2.CurrentCell.RowIndex - 1);
                dataGridView2.CurrentCell = dataGridView2[0, dataGridView2.CurrentCell.RowIndex - 1];
                IsChanged = true;
            }
        }
    }
    
    private void ClickButtonDown2(object sender, EventArgs e)
    {
        if(dataGridView2.CurrentCell != null)
        {
            if(dataGridView2.CurrentCell.RowIndex < dataGridView2.RowCount - 1)
            {
                var row = dataSource2[dataGridView2.CurrentCell.RowIndex];
                dataSource2[dataGridView2.CurrentCell.RowIndex] = dataSource2[dataGridView2.CurrentCell.RowIndex + 1];
                dataSource2[dataGridView2.CurrentCell.RowIndex + 1] = row;
                SetStrikeoutDataGridView2(dataGridView2.CurrentCell.RowIndex);
                SetStrikeoutDataGridView2(dataGridView2.CurrentCell.RowIndex + 1);
                dataGridView2.CurrentCell = dataGridView2[0, dataGridView2.CurrentCell.RowIndex + 1];
                IsChanged = true;
            }
        }
    }
    
    private void CheckedChangedCheckBoxLookahead(object sender, EventArgs e)
    {
        IsChanged = true;
    }
    
    private void PreviewKeyDownTextBoxHotKey(object sender, PreviewKeyDownEventArgs e)
    {
        var keyString = GetKeyString(e.KeyData);
        if(keyString == string.Empty)
        {
            hotKeyData = Keys.None;
            textBoxHotKey.Text = string.Empty;
        }
        else
        {
            hotKeyData = e.KeyData;
            textBoxHotKey.Text = keyString;
        }
    }
    
    private string GetKeyString(Keys keyData)
    {
        var modKey = (keyData & Keys.Modifiers);
        var modStr = string.Empty;
        if((modKey & Keys.Shift) == Keys.Shift)
        {
            modStr += modStr == string.Empty ? "Shift" : " + Shift";
        }
        if((modKey & Keys.Control) == Keys.Control)
        {
            modStr += modStr == string.Empty ? "Control" : " + Control";
        }
        if((modKey & Keys.Alt) == Keys.Alt)
        {
            modStr += modStr == string.Empty ? "Alt" : " + Alt";
        }
        
        var hotKey = (keyData & Keys.KeyCode);
        var hotStr = string.Empty;
        if(Keys.D0 <= hotKey && hotKey <= Keys.D9)
        {
            hotStr = ((int)hotKey - (int)Keys.D0).ToString();
        }
        else if(Keys.A <= hotKey && hotKey <= Keys.Z)
        {
            hotStr = hotKey.ToString();
        }
        else if(Keys.F1 <= hotKey && hotKey <= Keys.F12)
        {
            hotStr = hotKey.ToString();
        }
        
        if(modStr != string.Empty && hotStr != string.Empty)
        {
            return modStr + " + " + hotStr;
        }
        else
        {
            return string.Empty;
        }
    }
    
    private void ClickButtonSave(object sender, EventArgs e)
    {
        var disable = true;
        foreach(var data in dataSource1)
        {
            if(!data.Disable)
            {
                disable = false;
                break;
            }
        }
        if(disable)
        {
            MessageBox.Show("検索フォルダを指定してください。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        var config = new Config();
        config.ClearPath();
        foreach(var data in dataSource1)
        {
            config.AddPath(data.Path, data.Disable);
        }
        config.ClearEditor();
        foreach(var data in dataSource2)
        {
            config.AddEditor(data.Extension, data.Editor, data.Disable);
        }
        config.TopMost = checkBoxTopMost.Checked;
        config.AutoTrim = checkBoxAutoTrim.Checked;
        config.Parallel = checkBoxParallel.Checked;
        config.Minimize = checkBoxMinimize.Checked;
        config.Lookahead = checkBoxLookahead.Checked;
        config.TaskTray = checkBoxTaskTray.Checked;
        config.HotKey = hotKeyData;
        config.Save();
        this.Close();
    }
    
    private void ClickButtonCancel(object sender, EventArgs e)
    {
        this.Close();
    }
    
    private void LoadData()
    {
        var config = new Config();
        foreach(var path in config.Path)
        {
            var row1 = new Columns1();
            row1.Path = path.Value;
            row1.Disable = path.Disable;
            dataSource1.Add(row1);
        }
        foreach(var editor in config.Editor)
        {
            var row2 = new Columns2();
            row2.Extension = editor.Attribute;
            row2.Editor = editor.Value;
            row2.Disable = editor.Disable;
            dataSource2.Add(row2);
        }
        checkBoxTopMost.Checked = config.TopMost;
        checkBoxAutoTrim.Checked = config.AutoTrim;
        checkBoxParallel.Checked = config.Parallel;
        checkBoxMinimize.Checked = config.Minimize;
        checkBoxLookahead.Checked = config.Lookahead;
        checkBoxTaskTray.Checked = config.TaskTray;
        var keyString = GetKeyString(config.HotKey);
        if(keyString != string.Empty)
        {
            hotKeyData = config.HotKey;
            textBoxHotKey.Text = keyString;
        }
        IsChanged = false;
    }
}

class Config
{
    public readonly string FileName = "config.xml";
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int Count { get; set; }
    public bool TopMost { get; set; }
    public bool AutoTrim { get; set; }
    public bool Parallel { get; set; }
    public bool Minimize { get; set; }
    public bool Lookahead { get; set; }
    public bool TaskTray { get; set; }
    public Keys HotKey { get; set; }
    public List<Element> Path { get; set; }
    public List<Element> Editor { get; set; }
    public List<Element> History { get; set; }
    
    public Config()
    {
        Load();
    }
    
    public class Element
    {
        public string Attribute { get; set; }
        public string Value { get; set; }
        public bool Disable { get; set; }
        public Element()
        {
            Attribute = string.Empty;
            Value = string.Empty;
            Disable = false;
        }
    }
    
    public void Load()
    {
        var xml = GetDocument().Element("config");
        X = GetIntValue(xml, "x");
        Y = GetIntValue(xml, "y");
        Width = GetIntValue(xml, "width");
        Height = GetIntValue(xml, "height");
        Count = GetIntValue(xml, "count");
        TopMost = GetBoolValue(xml, "topmost");
        AutoTrim = GetBoolValue(xml, "trim");
        Parallel = GetBoolValue(xml, "parallel");
        Minimize = GetBoolValue(xml, "minimize");
        Lookahead = GetBoolValue(xml, "lookahead");
        TaskTray = GetBoolValue(xml, "tasktray");
        HotKey = GetKeysValue(xml, "hotkey");
        Path = GetListValue(xml, "path");
        Editor = GetListValue(xml, "editor", "extension");
        History = GetListValue(xml, "history");
    }
    
    public void Save()
    {
        var xml = GetDocument().Element("config");
        SetElementValue(xml, "x", X);
        SetElementValue(xml, "y", Y);
        SetElementValue(xml, "width", Width);
        SetElementValue(xml, "height", Height);
        SetElementValue(xml, "count", Count);
        SetElementValue(xml, "topmost", TopMost);
        SetElementValue(xml, "trim", AutoTrim);
        SetElementValue(xml, "parallel", Parallel);
        SetElementValue(xml, "minimize", Minimize);
        SetElementValue(xml, "lookahead", Lookahead);
        SetElementValue(xml, "tasktray", TaskTray);
        SetElementValue(xml, "hotkey", HotKey);
        SetElementValue(xml, "path", Path);
        SetElementValue(xml, "editor", "extension", Editor);
        SetElementValue(xml, "history", History);
        xml.Save(FileName);
    }
    
    public void ClearPath()
    {
        Path = new List<Element>();
    }
    
    public void ClearEditor()
    {
        Editor = new List<Element>();
    }
    
    public void ClearHistory()
    {
        History = new List<Element>();
    }
    
    public void AddPath(string val, bool disable)
    {
        Path.Add(ConvertElement(string.Empty, val, disable));
    }
    
    public void AddEditor(string attribute, string val, bool disable)
    {
        Editor.Add(ConvertElement(attribute, val, disable));
    }
    
    public void AddHistory(string val)
    {
        History.Add(ConvertElement(string.Empty, val, false));
    }
    
    private Element ConvertElement(string attribute, string val, bool disable)
    {
        var element = new Element();
        element.Attribute = attribute;
        element.Value = val;
        element.Disable = disable;
        return element;
    }
    
    private XDocument GetDocument()
    {
        if(File.Exists(FileName))
        {
            return XDocument.Load(FileName);
        }
        else
        {
            var xml = new XDocument(
                new XDeclaration("1.0", "utf-8", string.Empty),
                new XElement("config",
                    new XElement("x", "10"),
                    new XElement("y", "10"),
                    new XElement("width", "220"),
                    new XElement("height", "140"),
                    new XElement("count", "10"),
                    new XElement("topmost", "False"),
                    new XElement("trim", "False"),
                    new XElement("parallel", "False"),
                    new XElement("minimize", "False"),
                    new XElement("lookahead", "False"),
                    new XElement("tasktray", "False"),
                    new XElement("hotkey", "None"),
                    new XElement("path", Environment.GetFolderPath(Environment.SpecialFolder.Personal)),
                    new XElement("editor", @"C:\Windows\notepad.exe", new XAttribute("extension", "txt"))
                )
            );
            xml.Save(FileName);
            return xml;
        }
    }
    
    private int GetIntValue(XElement node, string name)
    {
        int val = 0;
        if(node.Element(name) != null)
        {
            Int32.TryParse(node.Element(name).Value, out val);
        }
        return val;
    }
    
    private bool GetBoolValue(XElement node, string name)
    {
        bool val = false;
        if(node.Element(name) != null)
        {
            Boolean.TryParse(node.Element(name).Value, out val);
        }
        return val;
    }
    
    private Keys GetKeysValue(XElement node, string name)
    {
        Keys val = Keys.None;
        if(node.Element(name) != null)
        {
            Enum.TryParse<Keys>(node.Element(name).Value, out val);
        }
        return val;
    }
    
    private List<Element> GetListValue(XElement node, string name, string attribute)
    {
        var items = new List<Element>();
        foreach(var element in node.Elements(name))
        {
            var item = new Element();
            item.Value = element.Value;
            if(String.IsNullOrWhiteSpace(attribute))
            {
                item.Attribute = string.Empty;
            }
            else if(element.Attribute(attribute) == null)
            {
                item.Attribute = string.Empty;
            }
            else
            {
                item.Attribute = element.Attribute(attribute).Value;
            }
            if(element.Attribute("disable") == null)
            {
                item.Disable = false;
            }
            else
            {
                bool disable;
                Boolean.TryParse(element.Attribute("disable").Value, out disable);
                item.Disable = disable;
            }
            items.Add(item);
        }
        return items;
    }
    
    private List<Element> GetListValue(XElement node, string name)
    {
        return GetListValue(node, name, string.Empty);
    }
    
    private void SetElementValue(XElement node, string name, string val)
    {
        if(node.Element(name) == null)
        {
            node.Add(new XElement(name, val));
        }
        else
        {
            node.Element(name).Value = val;
        }
    }
    
    private void SetElementValue(XElement node, string name, int val)
    {
        SetElementValue(node, name, val.ToString());
    }
    
    private void SetElementValue(XElement node, string name, bool val)
    {
        SetElementValue(node, name, val.ToString());
    }
    
    private void SetElementValue(XElement node, string name, Keys val)
    {
        SetElementValue(node, name, val.ToString());
    }
    
    private void SetElementValue(XElement node, string name, string attribute, List<Element> items)
    {
        if(node.Element(name) != null)
        {
            foreach(var element in node.Elements(name).ToList())
            {
                element.Remove();
            }
        }
        foreach(var item in items)
        {
            if(item.Disable)
            {
                if(String.IsNullOrWhiteSpace(attribute))
                {
                    node.Add(new XElement(name, item.Value, new XAttribute("disable", "True")));
                }
                else
                {
                    node.Add(new XElement(name, item.Value, new XAttribute(attribute, item.Attribute), new XAttribute("disable", "True")));
                }
            }
            else
            {
                if(String.IsNullOrWhiteSpace(attribute))
                {
                    node.Add(new XElement(name, item.Value));
                }
                else
                {
                    node.Add(new XElement(name, item.Value, new XAttribute(attribute, item.Attribute)));
                }
            }
        }
    }
    
    private void SetElementValue(XElement node, string name, List<Element> items)
    {
        SetElementValue(node, name, string.Empty, items);
    }
}

class HotKey : IDisposable
{
    public event EventHandler HotKeyPush;
    
    public enum ModKey
    {
        None = 0x0000,
        Alt = 0x0001,
        Control = 0x0002,
        Shift = 0x0004,
    }
    
    private HotKeyForm form;
    
    public HotKey(ModKey modKey, Keys key)
    {
        form = new HotKeyForm(modKey, key, raiseHotKeyPush);
    }
    
    private void raiseHotKeyPush()
    {
        if(HotKeyPush != null)
        {
            HotKeyPush(this, EventArgs.Empty);
        }
    }
    
    public void Dispose()
    {
        form.Dispose();
    }
    
    private class HotKeyForm : Form
    {
        [DllImport("user32.dll")]
        private extern static int RegisterHotKey(IntPtr hWnd, int id, ModKey modKey, Keys key);
        
        [DllImport("user32.dll")]
        private extern static int UnregisterHotKey(IntPtr hWnd, int id);
        
        private ThreadStart proc;
        private int hotKeyId = 0x000;
        
        public HotKeyForm(ModKey modKey, Keys key, ThreadStart proc)
        {
            this.proc = proc;
            for(int i = 0x000; i <= 0xbfff; i++)
            {
                if(RegisterHotKey(Handle, i, modKey, key) != 0)
                {
                    hotKeyId = i;
                    break;
                }
            }
        }
        
        protected override void WndProc(ref Message message)
        {
            base.WndProc(ref message);
            if(message.Msg == 0x0312)
            {
                if((int)message.WParam == hotKeyId)
                {
                    proc();
                }
            }
        }
        
        protected override void Dispose(bool disposing)
        {
            UnregisterHotKey(Handle, hotKeyId);
            base.Dispose(disposing);
        }
    }
}
