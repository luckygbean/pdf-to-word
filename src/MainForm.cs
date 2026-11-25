using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

internal sealed class MainForm : Form
{
    private readonly TextBox input = new TextBox(), password = new TextBox();
    private readonly ComboBox mode = new ComboBox(), quality = new ComboBox();
    private readonly Label hint = new Label(), status = new Label();
    private readonly ProgressBar bar = new ProgressBar();
    private readonly Button browse = new Button(), convert = new Button(), open = new Button();
    private bool busy;
    private string lastFile;
    private static readonly Color Ink = Color.FromArgb(27,42,57), Accent = Color.FromArgb(17,92,87);
    private static readonly string[] Modes = { "fixed", "edit", "ocr" };
    private static readonly string[] Suffix = { "原版排版", "可编辑", "识别文字" };
    private static readonly string[] Hints = {
        "每页按原尺寸嵌入高清图片，保留字体、图章和表格位置。页内文字作为图片保存。",
        "含线框表格的文档按原页重建文字和可填写的表格，图章、扫描页保留图像；其他文档由桌面版 Word 重排。",
        "用 Windows OCR 将扫描页识别为可编辑文字。表格和印章需对照原件，数字与标点建议校对。"
    };

    public MainForm(string initial)
    {
        Text = "PDF 转 Word";
        Font = new Font("Microsoft YaHei UI",10);
        ForeColor = Ink;
        BackColor = Color.White;
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(666,479);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        var title = LabelAt("PDF 转 Word",28,20,350,47);
        title.Font = new Font(Font.FontFamily,22,FontStyle.Bold);
        LabelAt("PDF 文件",30,85,230,24);
        Add(input,30,111,493,29);
        SetButton(browse,"选择文件…",532,109,102,32,false);
        browse.Click += delegate { ChooseInput(); };
        LabelAt("转换方式",30,159,240,24);
        Add(mode,30,186,324,31);
        mode.DropDownStyle = ComboBoxStyle.DropDownList;
        mode.Items.AddRange(new object[]{"保留原版排版","可编辑文字和表格","扫描文字识别"});
        mode.SelectedIndexChanged += delegate { UpdateMode(); };
        LabelAt("页面清晰度",374,159,160,24);
        Add(quality,374,186,260,31);
        quality.DropDownStyle = ComboBoxStyle.DropDownList;
        quality.Items.AddRange(new object[]{"300 DPI","450 DPI","600 DPI"});
        quality.SelectedIndex = 0;
        Add(hint,30,227,604,48);
        hint.BackColor = Color.FromArgb(239,246,245);
        hint.ForeColor = Color.FromArgb(41,78,76);
        hint.Padding = new Padding(8,3,8,3);
        hint.TextAlign = ContentAlignment.MiddleLeft;
        LabelAt("PDF 打开密码（如有）",30,289,236,25);
        Add(password,265,286,369,28);
        password.UseSystemPasswordChar = true;
        Add(bar,30,342,604,6);
        Add(status,30,355,604,48);
        SetButton(convert,"导出 Word…",488,414,146,40,true);
        SetButton(open,"打开结果",365,414,110,40,false);
        open.Enabled = false;
        convert.Click += async delegate { await SaveAsync(); };
        open.Click += delegate {
            try { Process.Start(new ProcessStartInfo(lastFile){UseShellExecute=true}); }
            catch(Exception e) { MessageBox.Show(this,e.Message,Text); }
        };
        mode.SelectedIndex = 0;
        status.Text = "可将一个 PDF 文件拖入窗口。";
        AllowDrop = true;
        DragEnter += delegate(object sender,DragEventArgs e) {
            e.Effect = !busy && e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        };
        DragDrop += delegate(object sender,DragEventArgs e) {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (!busy && files!=null && files.Length==1 && String.Equals(Path.GetExtension(files[0]),".pdf",StringComparison.OrdinalIgnoreCase))
                SetInput(files[0]);
            else status.Text = "请一次拖入一个 PDF 文件。";
        };
        input.TextChanged += delegate { password.Clear(); lastFile=null; open.Enabled=false; };
        if (!String.IsNullOrEmpty(initial)) SetInput(initial);
        FormClosing += delegate(object sender, FormClosingEventArgs e) {
            if (busy) { e.Cancel=true; status.Text="正在生成文档，完成后可关闭窗口。"; }
        };
    }

    private void Add(Control c,int x,int y,int width,int height)
    {
        c.SetBounds(x,y,width,height);
        Controls.Add(c);
    }
    private Label LabelAt(string text,int x,int y,int width,int height)
    {
        var c = new Label { Text=text };
        Add(c,x,y,width,height);
        return c;
    }
    private void SetButton(Button b,string caption,int x,int y,int width,int height,bool primary)
    {
        b.Text=caption; b.FlatStyle=FlatStyle.Flat; b.Cursor=Cursors.Hand;
        b.ForeColor=primary?Color.White:Ink; b.BackColor=primary?Accent:Color.White;
        b.FlatAppearance.BorderColor=primary?Accent:Color.FromArgb(192,200,206);
        Add(b,x,y,width,height);
    }
    private void SetInput(string path) { input.Text=Path.GetFullPath(path); }

    private void ChooseInput()
    {
        using (var dialog=new OpenFileDialog {Filter="PDF 文件 (*.pdf)|*.pdf",Title="选择 PDF"})
            if (dialog.ShowDialog(this)==DialogResult.OK) SetInput(dialog.FileName);
    }
    private void UpdateMode()
    {
        hint.Text=Hints[mode.SelectedIndex];
        quality.Enabled=!busy && mode.SelectedIndex==0;
        password.Enabled=!busy && mode.SelectedIndex!=1;
        if(mode.SelectedIndex==1) password.Clear();
    }
    private void SetBusy(bool value)
    {
        busy=value;
        input.Enabled=browse.Enabled=mode.Enabled=convert.Enabled=!value;
        open.Enabled=!value && lastFile!=null && File.Exists(lastFile);
        UpdateMode();
    }
    private async Task SaveAsync()
    {
        string source=input.Text.Trim().Trim('"');
        if (!File.Exists(source)) { ChooseInput(); source=input.Text.Trim().Trim('"'); }
        if (!File.Exists(source)) return;
        int index=mode.SelectedIndex, dpi=new[]{300,450,600}[quality.SelectedIndex];
        string key=password.Text;
        using (var dialog=new SaveFileDialog{Filter="Word 文档 (*.docx)|*.docx",DefaultExt="docx",Title="导出到新的 Word 文件",
            InitialDirectory=Path.GetDirectoryName(Path.GetFullPath(source)),OverwritePrompt=false})
        {
            string stem=Path.GetFileNameWithoutExtension(source)+"-"+Suffix[index];
            string name=stem+".docx";
            for(int n=2; File.Exists(Path.Combine(dialog.InitialDirectory,name)); n++) name=stem+"-"+n+".docx";
            dialog.FileName=name;
            if(dialog.ShowDialog(this)!=DialogResult.OK) return;
            string target=dialog.FileName;
            SetBusy(true); bar.Value=0;
            bar.Style=index==1?ProgressBarStyle.Marquee:ProgressBarStyle.Continuous;
            status.Text=index==1?"正在识别表格线和文字位置…":"正在读取页面…";
            var progress = new Progress<Tuple<int,int>>(p=>{
                bar.Style=ProgressBarStyle.Continuous;
                bar.Value=(int)((long)p.Item1*100/p.Item2);
                status.Text="已处理 "+p.Item1+" / "+p.Item2+" 页";
            });
            try
            {
                string report=await Task.Run(()=>Converter.RunAsync(Modes[index],source,target,dpi,key,
                    (i,n)=>((IProgress<Tuple<int,int>>)progress).Report(Tuple.Create(i,n))));
                lastFile=target;
                status.Text="已导出 · "+Path.GetFileName(target);
                bar.Style=ProgressBarStyle.Continuous; bar.Value=100;
                if (!String.IsNullOrEmpty(report)) MessageBox.Show(this,report,"转换结果",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            catch(Exception e)
            {
                status.Text="转换未完成。";
                bar.Style=ProgressBarStyle.Continuous; bar.Value=0;
                string reason=e.HResult==unchecked((int)0x8007052B)||e.HResult==unchecked((int)0x80070056) ?
                    "PDF 需要正确的打开密码。" : e.Message;
                MessageBox.Show(this,reason,Text,MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
            finally { SetBusy(false); }
        }
    }
}
