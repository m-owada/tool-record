using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.MediaFoundation;
using NAudio.Wave;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion ("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0")]
[assembly: AssemblyTitle("")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyProduct("Record")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyCopyright("Copyright (c) 2022 m-owada.")]
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
        catch(Exception ex)
        {
            MessageBox.Show(ex.Message, ex.Source);
            Application.Exit();
        }
    }
}

class MainForm : Form
{
    private ListBox listBox = new ListBox();
    private TextBox textBox = new TextBox();
    private ComboBox comboBox = new ComboBox();
    private CheckBox checkBox = new CheckBox();
    private Button buttonPlay = new Button();
    private Button buttonRec = new Button();
    private System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
    private Point mousePoint = new Point();
    
    private Player player = new Player();
    private Recorder recorder = new Recorder();
    
    private const string formName = "Record";
    private const string saveDirectory = "save";
    
    public MainForm()
    {
        // フォーム
        this.Size = new Size(340, 220);
        this.MinimumSize = this.Size;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.FormClosing += OnFormClosing;
        
        // リストボックス
        listBox.MouseDown += MouseDownListBox;
        listBox.DoubleClick += DoubleClickListBox;
        listBox.Location = new Point(10, 10);
        listBox.Size = new Size(305, 100);
        listBox.IntegralHeight = false;
        listBox.SelectionMode = SelectionMode.One;
        listBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(listBox);
        
        // テキストボックス
        textBox.Text = recorder.Device.ToString();
        textBox.Location = new Point(10, 120);
        textBox.Size = new Size(305, 20);
        textBox.ReadOnly = true;
        textBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(textBox);
        
        // ラベル
        var label = new Label();
        label.Text = "ビットレート";
        label.Location = new Point(10, 150);
        label.Size = new Size(label.PreferredWidth, 20);
        label.TextAlign = ContentAlignment.MiddleLeft;
        label.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(label);
        
        // コンボボックス
        comboBox.Location = new Point(70, 150);
        comboBox.Size = new Size(70, 20);
        comboBox.Items.Add("96kbps");
        comboBox.Items.Add("128kbps");
        comboBox.Items.Add("192kbps");
        comboBox.Items.Add("320kbps");
        comboBox.SelectedIndex = 0;
        comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
        comboBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(comboBox);
        
        // チェックボックス
        checkBox.CheckedChanged += CheckedChangedCheckBox;
        checkBox.Text = "最前面";
        checkBox.Location = new Point(150, 150);
        checkBox.Size = new Size(70, 20);
        checkBox.Checked = true;
        checkBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(checkBox);
        
        // 再生ボタン
        buttonPlay.Click += ClickButtonPlay;
        buttonPlay.Text = "再生";
        buttonPlay.Location = new Point(230, 150);
        buttonPlay.Size = new Size(40, 20);
        buttonPlay.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        this.Controls.Add(buttonPlay);
        
        // 録音ボタン
        buttonRec.Click += ClickButtonRec;
        buttonRec.Text = "録音";
        buttonRec.Location = new Point(275, 150);
        buttonRec.Size = new Size(40, 20);
        buttonRec.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        this.Controls.Add(buttonRec);
        
        // タイマー
        timer.Tick += TickTimer;
        timer.Interval = 100;
        timer.Start();
        
        player.Stopped += StoppedPlayer;
        recorder.Stopped += StoppedRecorder;
        AddItemsListBox();
    }
    
    private void OnFormClosing(object sender, FormClosingEventArgs e)
    {
        player.Dispose();
        recorder.Dispose();
    }
    
    private void MouseDownListBox(object sender, MouseEventArgs e)
    {
        mousePoint = e.Location;
    }
    
    private void DoubleClickListBox(object sender, EventArgs e)
    {
        int index = listBox.IndexFromPoint(mousePoint);
        if(index >= 0)
        {
            listBox.SelectedIndex = index;
            PlayStart();
        }
    }
    
    private void CheckedChangedCheckBox(object sender, EventArgs e)
    {
        this.TopMost = checkBox.Checked;
    }
    
    private void ClickButtonPlay(object sender, EventArgs e)
    {
        if(buttonPlay.Text == "再生")
        {
            PlayStart();
        }
        else
        {
             PlayStop();
        }
    }
    
    private void PlayStart()
    {
        var path = saveDirectory + @"\" + listBox.Text;
        if(File.Exists(path))
        {
            player.Start(saveDirectory + @"\" + listBox.Text);
            buttonPlay.Text = "停止";
            ClickButtonPlayEnabled(false);
        }
        else
        {
            MessageBox.Show("ファイルが存在しません。", formName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            AddItemsListBox();
        }
    }
    
    private void PlayStop()
    {
        player.Stop();
        buttonPlay.Text = "再生";
        ClickButtonPlayEnabled(true);
    }
    
    private void ClickButtonPlayEnabled(bool enabled)
    {
        listBox.Enabled = enabled;
        comboBox.Enabled = enabled;
        buttonPlay.Enabled = true;
        buttonRec.Enabled = enabled;
        buttonPlay.Select();
        this.Refresh();
    }
    
    private void StoppedPlayer(object sender, EventArgs e)
    {
        PlayStop();
    }
    
    private void ClickButtonRec(object sender, EventArgs e)
    {
        if(buttonRec.Text == "録音")
        {
            RecordStart();
        }
        else
        {
            RecordStop();
        }
    }
    
    private void RecordStart()
    {
        recorder.Start(saveDirectory, GetBitRate());
        buttonRec.Text = "停止";
        ClickButtonRecEnabled(false);
    }
    
    private void RecordStop()
    {
        recorder.Stop();
        buttonRec.Text = "録音";
        ClickButtonRecEnabled(true);
    }
    
    private void ClickButtonRecEnabled(bool enabled)
    {
        listBox.Enabled = enabled;
        comboBox.Enabled = enabled;
        buttonPlay.Enabled = enabled;
        buttonRec.Enabled = true;
        buttonRec.Select();
        this.Refresh();
    }
    
    private void StoppedRecorder(object sender, EventArgs e)
    {
        AddItemsListBox();
    }
    
    private void AddItemsListBox()
    {
        var text = listBox.Text;
        listBox.Items.Clear();
        if(Directory.Exists(saveDirectory))
        {
            foreach(var path in Directory.EnumerateFiles(saveDirectory, "*.mp3"))
            {
                listBox.Items.Add(Path.GetFileName(path));
            }
            if(listBox.Items.Count > 0)
            {
                var index = listBox.FindStringExact(text);
                if(index >= 0)
                {
                    listBox.SelectedIndex = index;
                }
                else
                {
                    listBox.SelectedIndex = 0;
                }
            }
        }
    }
    
    private int GetBitRate()
    {
        var bitRate = 0;
        switch(comboBox.Text)
        {
            case "96kbps":
                bitRate = 96000;
                break;
            case "128kbps":
                bitRate = 128000;
                break;
            case "192kbps":
                bitRate = 192000;
                break;
            case "320kbps":
                bitRate = 320000;
                break;
        }
        return bitRate;
    }
    
    private void TickTimer(object sender, EventArgs e)
    {
        var time = new TimeSpan();
        if(player.IsPlaying)
        {
            time = player.GetTime();
        }
        else if(recorder.IsRecording)
        {
            time = recorder.GetTime();
        }
        this.Text = formName + " (" + time.ToString(@"hh\:mm\:ss") + ")";
    }
}

class Player : IDisposable
{
    public bool IsPlaying { get; private set; }
    public event EventHandler Stopped;
    
    private AudioFileReader audioReader;
    private WaveOut waveOut;
    
    public Player()
    {
        IsPlaying = false;
    }
    
    private void Init(string path)
    {
        audioReader = new AudioFileReader(path);
        audioReader.Position = 0;
        waveOut = new WaveOut();
        waveOut.Init(audioReader);
        waveOut.PlaybackStopped += PlaybackStopped;
    }
    
    private void PlaybackStopped(object sender, StoppedEventArgs e)
    {
        Stop();
        Dispose();
        GC.Collect();
        Stopped(this, EventArgs.Empty);
    }
    
    public void Start(string path)
    {
        if(File.Exists(path))
        {
            Init(path);
            waveOut.Play();
            IsPlaying = true;
        }
    }
    
    public void Stop()
    {
        if(IsPlaying)
        {
            waveOut.Stop();
            IsPlaying = false;
        }
    }
    
    public TimeSpan GetTime()
    {
        if(IsPlaying)
        {
            return audioReader.CurrentTime;
        }
        else
        {
            return new TimeSpan();
        }
    }
    
    public TimeSpan GetTotalTime()
    {
        if(IsPlaying)
        {
            return audioReader.TotalTime;
        }
        else
        {
            return new TimeSpan();
        }
    }
    
    public void Dispose()
    {
        if(waveOut != null)
        {
            waveOut.PlaybackStopped -= PlaybackStopped;
            waveOut.Dispose();
            waveOut = null;
        }
        if(audioReader != null)
        {
            audioReader.Dispose();
            audioReader = null;
        }
    }
}

class Recorder : IDisposable
{
    public MMDevice Device { get; private set; }
    public string SaveDirectory { get; private set; }
    public int BitRate { get; private set; }
    public bool IsRecording { get; private set; }
    public event EventHandler Stopped;
    
    private WasapiLoopbackCapture audioCapture;
    private MemoryStream audioStream;
    private WaveFileWriter audioWriter;
    
    public Recorder()
    {
        MediaFoundationApi.Startup();
        Device = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        IsRecording = false;
    }
    
    private void Init(string directory, int bitRate)
    {
        audioCapture = new WasapiLoopbackCapture(Device);
        audioStream = new MemoryStream(1024);
        audioWriter = new WaveFileWriter(audioStream, audioCapture.WaveFormat);
        audioCapture.DataAvailable += DataAvailable;
        audioCapture.RecordingStopped += RecordingStopped;
        SaveDirectory = directory;
        BitRate = bitRate;
    }
    
    private void DataAvailable(object sender, WaveInEventArgs e)
    {
        audioWriter.Write(e.Buffer, 0, e.BytesRecorded);
    }
    
    private void RecordingStopped(object sender, StoppedEventArgs e)
    {
        WriteMp3();
        Dispose();
        GC.Collect();
        Stopped(this, EventArgs.Empty);
    }
    
    private void WriteMp3()
    {
        audioWriter.Flush();
        audioStream.Flush();
        audioStream.Position = 0;
        
        Directory.CreateDirectory(SaveDirectory);
        var path = SaveDirectory + @"\" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".mp3";
        
        if(audioWriter.Position > 0)
        {
            using(var reader = new WaveFileReader(audioStream))
            {
                MediaFoundationEncoder.EncodeToMp3(reader, path, BitRate);
            }
        }
    }
    
    public void Start(string directory, int bitRate)
    {
        Init(directory, bitRate);
        audioCapture.StartRecording();
        IsRecording = true;
    }
    
    public void Stop()
    {
        if(IsRecording)
        {
            audioCapture.StopRecording();
            IsRecording = false;
        }
    }
    
    public TimeSpan GetTime()
    {
        if(IsRecording)
        {
            var time = (int)Math.Min(Int32.MaxValue, audioWriter.Position / audioWriter.WaveFormat.AverageBytesPerSecond);
            return new TimeSpan(0, 0, time);
        }
        else
        {
            return new TimeSpan();
        }
    }
    
    public void Dispose()
    {
        if(audioCapture != null)
        {
            audioCapture.DataAvailable -= DataAvailable;
            audioCapture.RecordingStopped -= RecordingStopped;
            audioCapture.Dispose();
            audioCapture = null;
        }
        if(audioWriter != null)
        {
            audioWriter.Dispose();
            audioWriter = null;
        }
        if(audioStream != null)
        {
            audioStream.Dispose();
            audioStream = null;
        }
    }
}
