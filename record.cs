using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Speech.AudioFormat;
using System.Speech.Recognition;
using System.Text;
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
    private ContextMenuStrip contextMenu = new ContextMenuStrip();
    private TextBox textBox = new TextBox();
    private TrackBar trackBar = new TrackBar();
    private ComboBox comboBox = new ComboBox();
    private CheckBox checkBox1 = new CheckBox();
    private CheckBox checkBox2 = new CheckBox();
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
        this.Size = new Size(400, 250);
        this.MinimumSize = this.Size;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.FormClosing += OnFormClosing;
        this.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);
        
        // リストボックス
        listBox.MouseDown += MouseDownListBox;
        listBox.DoubleClick += DoubleClickListBox;
        listBox.Location = new Point(10, 10);
        listBox.Size = new Size(365, 100);
        listBox.IntegralHeight = false;
        listBox.SelectionMode = SelectionMode.One;
        listBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(listBox);
        
        // コンテキストメニュー
        contextMenu.Items.Add("再生", null, Menu1ClickListBox);
        contextMenu.Items.Add("解析", null, Menu2ClickListBox);
        contextMenu.Items.Add("削除", null, Menu3ClickListBox);
        
        // テキストボックス
        textBox.Text = recorder.Device.ToString();
        textBox.Location = new Point(10, 120);
        textBox.Size = new Size(365, 20);
        textBox.ReadOnly = true;
        textBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(textBox);
        
        // トラックバー
        trackBar.Location = new Point(10, 150);
        trackBar.Size = new Size(365, 20);
        trackBar.AutoSize = false;
        trackBar.TickStyle = TickStyle.None;
        trackBar.Minimum = 0;
        trackBar.TickFrequency = 1;
        trackBar.SmallChange = 1;
        trackBar.LargeChange = 10;
        trackBar.Enabled = false;
        trackBar.Scroll += ScrollTrackBar;
        trackBar.EnabledChanged += EnabledChangedTrackBar;
        trackBar.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(trackBar);
        
        // ラベル
        var label = new Label();
        label.Text = "ビットレート";
        label.Location = new Point(10, 180);
        label.Size = new Size(label.PreferredWidth, 20);
        label.TextAlign = ContentAlignment.MiddleLeft;
        label.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(label);
        
        // コンボボックス
        comboBox.Location = new Point(70, 180);
        comboBox.Size = new Size(70, 20);
        comboBox.Items.Add("96kbps");
        comboBox.Items.Add("128kbps");
        comboBox.Items.Add("192kbps");
        comboBox.Items.Add("320kbps");
        comboBox.SelectedIndex = 0;
        comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
        comboBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(comboBox);
        
        // チェックボックス（マイク入力）
        checkBox1.Text = "マイク入力";
        checkBox1.Location = new Point(150, 180);
        checkBox1.Size = new Size(75, 20);
        checkBox1.Checked = false;
        checkBox1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(checkBox1);
        
        // チェックボックス（最前面）
        checkBox2.CheckedChanged += CheckedChangedCheckBox2;
        checkBox2.Text = "最前面";
        checkBox2.Location = new Point(225, 180);
        checkBox2.Size = new Size(65, 20);
        checkBox2.Checked = true;
        checkBox2.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(checkBox2);
        
        // 再生ボタン
        buttonPlay.Click += ClickButtonPlay;
        buttonPlay.Text = "再生";
        buttonPlay.Location = new Point(290, 180);
        buttonPlay.Size = new Size(40, 20);
        buttonPlay.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        this.Controls.Add(buttonPlay);
        
        // 録音ボタン
        buttonRec.Click += ClickButtonRec;
        buttonRec.Text = "録音";
        buttonRec.Location = new Point(335, 180);
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
        if(e.Button == MouseButtons.Right)
        {
            listBox.ContextMenuStrip = new ContextMenuStrip();
            int index = listBox.IndexFromPoint(mousePoint);
            if(index >= 0)
            {
                listBox.SelectedIndex = index;
                listBox.ContextMenuStrip = contextMenu;
            }
        }
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
    
    private void Menu1ClickListBox(object sender, EventArgs e)
    {
        PlayStart();
    }
    
    private void Menu2ClickListBox(object sender, EventArgs e)
    {
        var path = saveDirectory + @"\" + listBox.Text;
        if(File.Exists(path))
        {
            var subForm = new SubForm(formName, this.TopMost, path);
            subForm.ShowDialog();
        }
        else
        {
            MessageBox.Show("ファイルが存在しません。", formName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            AddItemsListBox();
        }
    }
    
    private void Menu3ClickListBox(object sender, EventArgs e)
    {
        if(MessageBox.Show("ファイルを削除しますか？", formName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
        {
            var path = saveDirectory + @"\" + listBox.Text;
            var fi = new FileInfo(path);
            if(fi.Exists)
            {
                if(!fi.IsReadOnly)
                {
                    fi.Delete();
                }
                else
                {
                    MessageBox.Show("ファイルが読み取り専用です。", formName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("ファイルが存在しません。", formName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            AddItemsListBox();
        }
    }
    
    private void ScrollTrackBar(object sender, EventArgs e)
    {
        player.SetCurrentTime(new TimeSpan(0, 0, trackBar.Value));
    }
    
    private void EnabledChangedTrackBar(object sender, EventArgs e)
    {
        trackBar.Value = 0;
    }
    
    private void CheckedChangedCheckBox2(object sender, EventArgs e)
    {
        this.TopMost = checkBox2.Checked;
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
            player.Start(path);
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
        trackBar.Enabled = !enabled;
        comboBox.Enabled = enabled;
        checkBox1.Enabled = enabled;
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
        recorder.Start(saveDirectory, GetBitRate(), checkBox1.Checked);
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
        trackBar.Enabled = false;
        comboBox.Enabled = enabled;
        checkBox1.Enabled = enabled;
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
            trackBar.Maximum = (int)player.GetTotalTime().TotalSeconds;
            trackBar.Value = (int)time.TotalSeconds;
        }
        else if(recorder.IsRecording)
        {
            time = recorder.GetTime();
        }
        this.Text = formName + " (" + time.ToString(@"hh\:mm\:ss") + ")";
    }
}

class SubForm : Form
{
    private Recognition recognition = new Recognition();
    private TextBox textBox = new TextBox();
    private string formName = string.Empty;
    
    public SubForm(string name, bool topMost, string path)
    {
        // 音声認識
        recognition.Start(path);
        recognition.Recognized += RecognitionRecognized;
        recognition.Completed += RecognitionCompleted;
        
        // フォーム
        this.Text = name + " (解析中)";
        this.Size = new Size(480, 360);
        this.MinimumSize = this.Size;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.TopMost = topMost;
        this.FormClosing += OnFormClosing;
        this.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);
        
        // テキストボックス
        textBox.Text = string.Empty;
        textBox.Location = new Point(10, 10);
        textBox.Size = new Size(445, 270);
        textBox.ReadOnly = true;
        textBox.Multiline = true;
        textBox.ScrollBars = ScrollBars.Vertical;
        textBox.WordWrap = true;
        textBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(textBox);
        
        // ラベル（サンプルレート）
        var label1 = new Label();
        label1.Text = "サンプルレート";
        label1.Location = new Point(10, 290);
        label1.Size = new Size(label1.PreferredWidth, 20);
        label1.TextAlign = ContentAlignment.MiddleLeft;
        label1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(label1);
        
        // テキストボックス（サンプルレート）
        var textBox1 = new TextBox();
        textBox1.Text = recognition.SampleRate.ToString("#,0");
        textBox1.Location = new Point(85, 290);
        textBox1.Size = new Size(50, 20);
        textBox1.ReadOnly = true;
        textBox1.TextAlign = HorizontalAlignment.Center;
        textBox1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(textBox1);
        
        // ラベル（ビット/サンプル）
        var label2 = new Label();
        label2.Text = "ビット/サンプル";
        label2.Location = new Point(145, 290);
        label2.Size = new Size(label2.PreferredWidth, 20);
        label2.TextAlign = ContentAlignment.MiddleLeft;
        label2.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(label2);
        
        // テキストボックス（ビット/サンプル）
        var textBox2 = new TextBox();
        textBox2.Text = recognition.BitsPerSample.ToString("#,0");
        textBox2.Location = new Point(220, 290);
        textBox2.Size = new Size(30, 20);
        textBox2.ReadOnly = true;
        textBox2.TextAlign = HorizontalAlignment.Center;
        textBox2.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(textBox2);
        
        // ラベル（チャンネル）
        var label3 = new Label();
        label3.Text = "チャンネル";
        label3.Location = new Point(260, 290);
        label3.Size = new Size(label3.PreferredWidth, 20);
        label3.TextAlign = ContentAlignment.MiddleLeft;
        label3.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(label3);
        
        // テキストボックス（チャンネル）
        var textBox3 = new TextBox();
        textBox3.Text = recognition.Channels == 1 ? "モノラル" : "ステレオ";
        textBox3.Location = new Point(315, 290);
        textBox3.Size = new Size(50, 20);
        textBox3.ReadOnly = true;
        textBox3.TextAlign = HorizontalAlignment.Center;
        textBox3.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(textBox3);
        
        // ラベル（長さ）
        var label4 = new Label();
        label4.Text = "長さ";
        label4.Location = new Point(375, 290);
        label4.Size = new Size(label4.PreferredWidth, 20);
        label4.TextAlign = ContentAlignment.MiddleLeft;
        label4.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(label4);
        
        // テキストボックス（長さ）
        var textBox4 = new TextBox();
        textBox4.Text = recognition.TotalTime.ToString(@"hh\:mm\:ss");
        textBox4.Location = new Point(405, 290);
        textBox4.Size = new Size(50, 20);
        textBox4.ReadOnly = true;
        textBox4.TextAlign = HorizontalAlignment.Center;
        textBox4.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(textBox4);
        
        formName = name;
    }
    
    private void RecognitionRecognized(object sender, SpeechRecognizedEventArgs e)
    {
        if(textBox.Text != string.Empty)
        {
            textBox.Text += Environment.NewLine;
        }
        textBox.Text += e.Result.Text;
    }
    
    private void RecognitionCompleted(object sender, RecognizeCompletedEventArgs e)
    {
        this.Text = formName + " (完了)";
    }
    
    private void OnFormClosing(object sender, FormClosingEventArgs e)
    {
        if(recognition.IsRecognizing)
        {
            this.Text = formName + " (停止中)";
        }
        recognition.Dispose();
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
        if(Stopped != null)
        {
            Stopped(this, EventArgs.Empty);
        }
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
    
    public void SetCurrentTime(TimeSpan time)
    {
        if(IsPlaying)
        {
            audioReader.CurrentTime = time;
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
    
    private WaveIn waveIn;
    private WaveOut waveOut;
    private BufferedWaveProvider bufferedWaveProvider;
    
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
        
        waveIn = new WaveIn();
        waveIn.WaveFormat = audioCapture.WaveFormat;
        waveIn.DataAvailable += DataAvailableWaveIn;
        waveOut = new WaveOut();
        bufferedWaveProvider = new BufferedWaveProvider(waveIn.WaveFormat);
        bufferedWaveProvider.DiscardOnBufferOverflow = true;
        waveOut.Init(bufferedWaveProvider);
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
        if(Stopped != null)
        {
            Stopped(this, EventArgs.Empty);
        }
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
    
    private void DataAvailableWaveIn(object sender, WaveInEventArgs e)
    {
        bufferedWaveProvider.AddSamples(e.Buffer, 0, e.BytesRecorded);
    }
    
    public void Start(string directory, int bitRate, bool micIn = false)
    {
        Init(directory, bitRate);
        audioCapture.StartRecording();
        IsRecording = true;
        
        if(micIn)
        {
            waveIn.StartRecording();
            waveOut.Play();
        }
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
        if(bufferedWaveProvider != null)
        {
            bufferedWaveProvider = null;
        }
        if(waveOut != null)
        {
            waveOut.Dispose();
            waveOut = null;
        }
        if(waveIn != null)
        {
            waveIn.DataAvailable -= DataAvailableWaveIn;
            waveIn.Dispose();
            waveIn = null;
        }
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

class Recognition : IDisposable
{
    public int SampleRate { get; private set; }
    public int BitsPerSample { get; private set; }
    public int Channels { get; private set; }
    public TimeSpan TotalTime { get; private set; }
    public bool IsRecognizing { get; private set; }
    
    public event EventHandler<SpeechRecognizedEventArgs> Recognized;
    public event EventHandler<RecognizeCompletedEventArgs> Completed;
    
    private SpeechRecognitionEngine engine;
    
    public Recognition()
    {
        IsRecognizing = false;
    }
    
    public void Init(string path)
    {
        using(var reader = new MediaFoundationReader(path))
        {
            SampleRate = reader.WaveFormat.SampleRate;
            BitsPerSample = reader.WaveFormat.BitsPerSample;
            Channels = reader.WaveFormat.Channels;
            TotalTime = reader.TotalTime;
            engine = new SpeechRecognitionEngine(Application.CurrentCulture);
            engine.LoadGrammar(new DictationGrammar());
            engine.SetInputToAudioStream(reader, new SpeechAudioFormatInfo(SampleRate, ConvBitsPerSample(BitsPerSample), ConvChannels(Channels)));
            engine.SpeechRecognized += SpeechRecognized;
            engine.RecognizeCompleted += RecognizeCompleted;
            engine.RecognizeAsync(RecognizeMode.Multiple);
        }
    }
    
    private AudioBitsPerSample ConvBitsPerSample(int bps)
    {
        if(bps == 8)
        {
            return AudioBitsPerSample.Eight;
        }
        else
        {
            return AudioBitsPerSample.Sixteen;
        }
    }
    
    private AudioChannel ConvChannels(int channels)
    {
        if(channels == 1)
        {
            return AudioChannel.Mono;
        }
        else
        {
            return AudioChannel.Stereo;
        }
    }
    
    public void Start(string path)
    {
        Init(path);
        IsRecognizing = true;
    }
    
    public void Stop()
    {
        if(IsRecognizing)
        {
            engine.RecognizeAsyncStop();
            IsRecognizing = false;
        }
    }
    
    private void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
    {
        if(e.Result != null && Recognized != null)
        {
            Recognized(this, e);
        }
    }
    
    private void RecognizeCompleted(object sender, RecognizeCompletedEventArgs e)
    {
        if(e.Result != null && Completed != null)
        {
            Stop();
            Completed(this, e);
        }
    }
    
    public void Dispose()
    {
        if(engine != null)
        {
            engine.SpeechRecognized -= SpeechRecognized;
            engine.RecognizeCompleted -= RecognizeCompleted;
            engine.Dispose();
            engine = null;
        }
    }
}
