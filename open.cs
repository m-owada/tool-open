using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion ("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0")]
[assembly: AssemblyTitle("")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyProduct("OPEN")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyCopyright("Copyright (c) 2020 m-owada.")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            Application.EnableVisualStyles();
            Application.ThreadException += (object sender, ThreadExceptionEventArgs e) =>
            {
                throw new Exception(e.Exception.Message);
            };
            Application.Run(new MainForm());
        }
        catch(Exception e)
        {
            MessageBox.Show(e.Message, e.Source);
            Application.Exit();
        }
    }
}

class MainForm : Form
{
    private ComboBox comboBox = new ComboBox();
    private CheckBox checkBox = new CheckBox();
    private ListBox listBox = new ListBox();
    private Dictionary<string, List<string>> lookaheadFiles = new Dictionary<string, List<string>>();
    private bool autoTrim = false;
    
    public MainForm()
    {
        // フォーム
        this.Text = "OPEN";
        this.Size = new Size(220, 160);
        this.MinimumSize  = this.Size;
        this.Location = new Point(10, 10);
        this.StartPosition = FormStartPosition.Manual;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.FormClosing += OnFormClosing;
        
        // コンボボックス
        comboBox.KeyDown += KeyDownComboBox;
        comboBox.KeyPress += KeyPressComboBox;
        comboBox.Text = string.Empty;
        comboBox.Location = new Point(10, 10);
        comboBox.Size = new Size(185, 20);
        comboBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        SetComboBoxItems();
        this.Controls.Add(comboBox);
        
        // チェックボックス
        checkBox.CheckedChanged += ChangedCheckBox;
        checkBox.Text = "常に手前に表示する";
        checkBox.Location = new Point(10, 35);
        checkBox.Size = new Size(125, 20);
        checkBox.Checked = true;
        checkBox.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        this.Controls.Add(checkBox);
        
        // リストボックス
        listBox.DoubleClick += DoubleClickListBox;
        listBox.KeyDown += KeyDownListBox;
        listBox.KeyPress += KeyPressListBox;
        listBox.Leave += LeaveListBox;
        listBox.MouseDown += MouseDownListBox;
        listBox.Location = new Point(10, 60);
        listBox.Size = new Size(185, 52);
        listBox.MultiColumn = false;
        listBox.SelectionMode = SelectionMode.MultiExtended;
        listBox.IntegralHeight = false;
        listBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(listBox);
        
        LoadConfigSetting();
    }
    
    private void OnFormClosing(object sender, FormClosingEventArgs e)
    {
        if(this.WindowState == FormWindowState.Normal)
        {
            var xml = LoadXmlFile("config.xml");
            var config = xml.Element("config");
            SetElementValue(config, "x", this.Location.X.ToString());
            SetElementValue(config, "y", this.Location.Y.ToString());
            SetElementValue(config, "width", this.Width.ToString());
            SetElementValue(config, "height", this.Height.ToString());
            SetElementValue(config, "topmost", checkBox.Checked.ToString());
            xml.Save("config.xml");
        }
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
    
    private async void KeyDownComboBox(object sender, KeyEventArgs e)
    {
        if(e.KeyCode == Keys.Enter)
        {
            comboBox.Text = comboBox.Text.Replace(@"/", "").Replace(@"\", "");
            if(autoTrim) comboBox.Text = comboBox.Text.Trim();
            if(!String.IsNullOrWhiteSpace(comboBox.Text))
            {
                await EnterComboBox();
            }
        }
    }
    
    private async Task EnterComboBox()
    {
        comboBox.Enabled = false;
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
        if(listBox.Items.Count == 0)
        {
            comboBox.SelectionStart = 0;
            comboBox.SelectionLength = comboBox.Text.Length;
            comboBox.Focus();
        }
        else if(listBox.Items.Count == 1)
        {
            OpenFiles();
            comboBox.Focus();
        }
        else
        {
            listBox.Focus();
        }
    }
    
    private Dictionary<string, List<string>> GetEnumerateFiles(string keyword)
    {
        var items = new Dictionary<string, List<string>>();
        var config = LoadXmlFile("config.xml").Element("config");
        var extensions = new List<string>();
        foreach(var extension in config.Elements("extension"))
        {
            if(!String.IsNullOrWhiteSpace(extension.Value))
            {
                extensions.Add(extension.Value);
            }
        }
        foreach(var path in config.Elements("path"))
        {
            foreach(var file in EnumerateFilesAllDirectories(path.Value, keyword + "*", extensions))
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
        var xml = LoadXmlFile("config.xml");
        var config = xml.Element("config");
        foreach(var history in config.Elements("history"))
        {
            if(word == history.Value)
            {
                history.Remove();
            }
        }
        if(config.Elements("history").Count() >= 10)
        {
            config.Elements("history").First().Remove();
        }
        config.Add(new XElement("history", word));
        xml.Save("config.xml");
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
        var config = LoadXmlFile("config.xml").Element("config");
        foreach(var history in config.Elements("history"))
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
    
    private void ChangedCheckBox(object sender, EventArgs e)
    {
        this.TopMost = checkBox.Checked;
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
    }
    
    private void KeyPressListBox(object sender, KeyPressEventArgs e)
    {
        if(e.KeyChar == (char)Keys.Enter)
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
        ContextMenuStrip menu = new ContextMenuStrip();
        foreach(var item in listBox.SelectedItems)
        {
            var file = (KeyValuePair<string, List<string>>)item;
            if(file.Value.Count <= 50)
            {
                foreach(var path in file.Value)
                {
                    ToolStripMenuItem menuItem = new ToolStripMenuItem();
                    menuItem.Text = path;
                    menuItem.Enabled = true;
                    menuItem.Click += MenuClickListBox;
                    menu.Items.Add(menuItem);
                }
            }
            else
            {
                ToolStripMenuItem menuItem = new ToolStripMenuItem();
                menuItem.Text = "ファイル数が50個を超えています。";
                menuItem.Enabled = false;
                menu.Items.Add(menuItem);
            }
            break;
        }
        return menu;
    }
    
    private void MenuClickListBox(object sender, EventArgs e)
    {
        ToolStripMenuItem menuItem = (ToolStripMenuItem)sender;
        foreach(var item in listBox.SelectedItems)
        {
            var file = (KeyValuePair<string, List<string>>)item;
            OpenFile(file.Key, menuItem.Text);
            break;
        }
    }
    
    private void OpenFiles()
    {
        foreach(var item in listBox.SelectedItems)
        {
            var file = (KeyValuePair<string, List<string>>)item;
            OpenFile(file.Key, file.Value.First());
        }
    }
    
    private void OpenFile(string name, string path)
    {
        if(File.Exists(path))
        {
            var config = LoadXmlFile("config.xml").Element("config");
            var app = string.Empty;
            foreach(var editor in config.Elements("editor"))
            {
                var target = editor.Attribute("target") == null ? string.Empty : editor.Attribute("target").Value;
                if(String.IsNullOrWhiteSpace(target))
                {
                    app = String.IsNullOrWhiteSpace(app) ? editor.Value : app;
                }
                else
                {
                    if(path.EndsWith("." + target, true, null))
                    {
                        app = editor.Value;
                        break;
                    }
                }
            }
            if(String.IsNullOrWhiteSpace(app))
            {
                MessageBox.Show("ファイル \"" + name + "\" を起動するアプリケーションが指定されていません。", this.Text);
            }
            else
            {
                Process.Start(app, path);
            }
        }
        else
        {
            MessageBox.Show("ファイル \"" + name + "\" が見つかりません。", this.Text);
        }
    }
    
    private void LoadConfigSetting()
    {
        var config = LoadXmlFile("config.xml").Element("config");
        
        var x = this.Location.X;
        foreach(var item in config.Elements("x"))
        {
            Int32.TryParse(item.Value, out x);
            break;
        }
        var y = this.Location.Y;
        foreach(var item in config.Elements("y"))
        {
            Int32.TryParse(item.Value, out y);
            break;
        }
        this.Location = new Point(x, y);
        
        var width = this.Width;
        foreach(var item in config.Elements("width"))
        {
            Int32.TryParse(item.Value, out width);
            break;
        }
        var height = this.Height;
        foreach(var item in config.Elements("height"))
        {
            Int32.TryParse(item.Value, out height);
            break;
        }
        this.Size = new Size(width, height);
        
        var topmost = checkBox.Checked;
        foreach(var item in config.Elements("topmost"))
        {
            Boolean.TryParse(item.Value, out topmost);
            break;
        }
        checkBox.Checked = topmost;
        
        var lookahead = false;
        foreach(var item in config.Elements("lookahead"))
        {
            Boolean.TryParse(item.Value, out lookahead);
            break;
        }
        if(lookahead) SetLookaheadFiles();
        
        var trim = false;
        foreach(var item in config.Elements("trim"))
        {
            Boolean.TryParse(item.Value, out trim);
            break;
        }
        autoTrim = trim;
    }
    
    private async void SetLookaheadFiles()
    {
        comboBox.Enabled = false;
        await Task.Run(() =>
        {
            lookaheadFiles = GetEnumerateFiles(string.Empty);
        });
        comboBox.Enabled = true;
        comboBox.Focus();
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
    
    private XDocument LoadXmlFile(string file)
    {
        if(File.Exists(file))
        {
            return XDocument.Load(file);
        }
        else
        {
            var xml = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                new XElement("config",
                    new XElement("x", this.Location.X.ToString()),
                    new XElement("y", this.Location.Y.ToString()),
                    new XElement("width", this.Width.ToString()),
                    new XElement("height", this.Height.ToString()),
                    new XElement("topmost", "True"),
                    new XElement("path", Environment.GetFolderPath(Environment.SpecialFolder.Personal)),
                    new XElement("editor", @"C:\Windows\notepad.exe"),
                    new XElement("extension", "txt"),
                    new XElement("lookahead", "False"),
                    new XElement("trim", "False")
                )
            );
            xml.Save(file);
            return xml;
        }
    }
}
