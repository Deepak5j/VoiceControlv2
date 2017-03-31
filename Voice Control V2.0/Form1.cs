using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.Drawing.Drawing2D;
namespace Voice_Control_V2._0
{
    public partial class Form1 : Form
    {
        //object declaration of synthesiser, promptbuilder, recognizer engine, choices
        SpeechSynthesizer ss = new SpeechSynthesizer();
        PromptBuilder pb = new PromptBuilder();
        SpeechRecognitionEngine sre = new SpeechRecognitionEngine();
        Choices clist = new Choices();

        //initializing componets
        public Form1()
        {

            //ss.SpeakAsync("Initializing the main Component");
            InitializeComponent();
        }
        //-----------------------changing form size--------------------------------
        /*
        protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
        {
            GraphicsPath wantedshape = new GraphicsPath();
            wantedshape.AddEllipse(0, 0, this.Width, this.Height);
            this.Region = new Region(wantedshape);
        }*/
        //-------------------------------------------------------------------------
        
        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.DrawRectangle(new Pen(Color.Black, 20),this.DisplayRectangle);
        }
        
        //foram loading
        private void Form1_Load(object sender, EventArgs e)
        {
            MinimizeBox = false;//disabling minimize button
            MaximizeBox = false;//disabling maximize button
            //System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            //path.AddEllipse(0, 0, this.Width, this.Height);
            //Region region = new Region(path);
            //this.Region = region;
            //ss.SpeakAsync("foram loading");
        }
        
        //button start code contents
        private void butstart_Click(object sender, EventArgs e)
        {
            //on click start button recognition start
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            //Choices commands = new Choices(File.ReadAllLines("commands.txt"));
            clist.Add(new string[] { "open my computer", "close my computer", "open control panel","close control panel", "open recycle bin", "open desktop", "open my videos",
                "open my documents", "open my musics", "open cmd", "open c drive", "open d drive", "open e drive", "open f drive", "open g drive", "open h drive", "open i drive",
                "stop recognition", "what is the current time", "what is time right now", "close the application", "close application", "open chrome", "open calculator",
                "open notepad", "left", "right", "up", "down", "d 1", "d 2", "d 3", "d 4", "c 1", "c 2", "c 3", "c 4", "c 5", "right click", "left click", "close window",
                "close cmd", "enter", "window", "tab", "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "f10", "f11", "f 12", "f13", "f14", "f15", "f16",
                "backspace", "break", "capslock", "delete", "down arrow", "end", "escape", "help", "home", "insert", "left arrow", "num lock", "page down", "page up",
                "print screen", "right arrow", "scroll lock", "up arrow", "open paint", "open my pictures"});
            Grammar gr = new Grammar(new GrammarBuilder(clist));
            try
            {
                sre.RequestRecognizerUpdate();
                sre.LoadGrammar(gr);
                sre.SpeechRecognized += sre_SpeechRecognized;
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Multiple);

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            //word recognize
        }

        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //event handler
            if (e.Result.Confidence > 0.0)
            {
                switch (e.Result.Text.ToString())
                {
                    

                    //_______________________________administrative utility______________________________
                    case "open control panel":
                        Process.Start("control.exe");
                        break;

                    case "open cmd":
                        Process.Start("cmd");
                        break;
                    
                    case "close cmd":
                        try
                        {
                            String p11 = GetForegroundProcessName();
                            Process[] proc1 = Process.GetProcessesByName(p11);
                            if (p11 == "cmd")
                                proc1[0].Kill();
                            else
                            {
                                Process[] prs = Process.GetProcesses();
                                int flg = 0;
                                foreach (Process p in prs)
                                {
                                    if (p.ProcessName == "cmd")
                                        flg = 1;
                                }
                                if (flg == 1)
                                    ss.SpeakAsync("command prompt is not active");
                                else
                                    ss.SpeakAsync("there is no command prompt running");
                            }
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Error!");
                        }
                        break;
                    
                    //______________________________explorer utility_____________________________
                    case "close window":
                        try
                        {
                            String p1 = GetForegroundProcessName();
                            Process[] proc = Process.GetProcessesByName(p1);
                            proc[0].Kill();
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("error!");
                        }
                        break;

                    case "open my computer":
                        System.Diagnostics.Process.Start("::{20d04fe0-3aea-1069-a2d8-08002b30309d}");
                        break;

                    case "open recycle bin":
                        System.Diagnostics.Process.Start("::{645FF040-5081-101B-9F08-00AA002F954E}");
                        break;

                    case "open c drive":
                        try
                        {
                            System.Diagnostics.Process.Start("c:");
                        }
                        catch (Exception)
                        {
                            ss.SpeakAsync("there is no dive c");
                        }
                        break;

                    case "open d drive":
                        try
                        {
                            System.Diagnostics.Process.Start("d:");
                        }
                        catch (Exception)
                        {
                            ss.SpeakAsync("there is no dive d");
                        }
                        break;

                    case "open e drive":
                        try
                        {
                            System.Diagnostics.Process.Start("e:");
                        }
                        catch (Exception)
                        {
                            ss.SpeakAsync("there is no dive e");
                        }
                        break;

                    case "open f drive":
                        try
                        {
                            System.Diagnostics.Process.Start("f:");
                        }
                        catch (Exception)
                        {
                            ss.SpeakAsync("there is no dive f");
                        }
                        break;

                    case "open g drive":
                        try
                        {
                            System.Diagnostics.Process.Start("g:");
                        }
                        catch (Exception)
                        {
                            ss.SpeakAsync("there is no dive g");
                        }
                        break;

                    case "open h drive":
                        try
                        {
                            System.Diagnostics.Process.Start("h:");
                        }
                        catch (Exception)
                        {
                            ss.SpeakAsync("there is no dive h");
                        }
                        break;

                    case "open i drive":
                        try
                        {
                            System.Diagnostics.Process.Start("i:");
                        }
                        catch (Exception)
                        {
                            ss.SpeakAsync("there is no dive i");
                        }
                        break;

                    case "open my documents":
                        System.Diagnostics.Process.Start("explorer", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                        break;

                    case "open my musics":
                        System.Diagnostics.Process.Start("explorer", Environment.GetFolderPath(Environment.SpecialFolder.MyMusic));
                        break;

                    case "open desktop":
                        System.Diagnostics.Process.Start("explorer", Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
                        break;

                    case "open my videos":                        
                        System.Diagnostics.Process.Start("explorer", Environment.GetFolderPath(Environment.SpecialFolder.MyVideos));
                        break;

                    case "open my pictures" :
                        System.Diagnostics.Process.Start("explorer", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures));
                        break;
                    //_____________________________web browser______________________________
                    case "open chrome": //open google chrome webbrowser
                        System.Diagnostics.Process.Start("chrome");
                        break;

                    //_______________________________general utility_______________________________
                    case "what is the current time":
                        ss.SpeakAsync("Current time is" + DateTime.Now.ToLongTimeString());
                        break;

                    case "what is time right now":
                        ss.SpeakAsync("time is" + DateTime.Now.ToLongTimeString());
                        break;

                    case "open notepad":
                        System.Diagnostics.Process.Start("notepad");
                        break;

                    case "open calculator":
                        System.Diagnostics.Process.Start("calc");
                        break;

                    case "open paint" :
                        System.Diagnostics.Process.Start("mspaint");
                        break;
                        
                    //_____________________________system command____________________________________
                    case "stop recognition":
                        sre.RecognizeAsyncStop();
                        btnStart.Enabled = true;
                        btnStop.Enabled = false;
                        break;

                    case "close the application": //exits applecattino
                        Application.Exit();
                        break;

                    case "close application": //exits applecattino
                        Application.Exit();
                        break;
                    //_______________________________________ mouse keys______________________________________________
                    case "left": //move cursor left
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X - 15, Cursor.Position.Y);
                        break;

                    case "right": //move cursor right
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X + 15, Cursor.Position.Y);
                        break;

                    case "up": //move cursor up
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X, Cursor.Position.Y - 15);
                        break;

                    case "down": //move cursor down
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X, Cursor.Position.Y + 15);
                        break;

                    case "d 1": //move cursor diagonaly up-left
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X - 30, Cursor.Position.Y - 30);
                        break;

                    case "d 2": //move cursor diagonaly up-right
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X + 30, Cursor.Position.Y - 30);
                        break;

                    case "d 3": //move cursor diagonaly down-left
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X - 30, Cursor.Position.Y + 30);
                        break;

                    case "d 4": //move cursor diagonaly down-right
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X + 30, Cursor.Position.Y + 30);
                        break;

                    case "c 1": //move cursor top-left corner
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X - 1366, Cursor.Position.Y - 768);
                        break;

                    case "c 2": //move top-right corner
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X + Screen.PrimaryScreen.Bounds.Width - 20, Cursor.Position.Y - Screen.PrimaryScreen.Bounds.Height);
                        break;

                    case "c 3": //move top-right corner
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X - Screen.PrimaryScreen.Bounds.Width, Cursor.Position.Y + Screen.PrimaryScreen.Bounds.Height - 20);
                        break;

                    case "c 4": //move top-right corner
                        this.Cursor = new Cursor(Cursor.Current.Handle);
                        Cursor.Position = new Point(Cursor.Position.X + Screen.PrimaryScreen.Bounds.Width - 20, Cursor.Position.Y + Screen.PrimaryScreen.Bounds.Height - 20);
                        break;

                    case "left click": //mose left click
                        txtContents.Text += e.Result.Text.ToString() + Environment.NewLine;
                        break;
                    //_____________________________________KEY BOARD KEYS_______________________________________________
                    case "enter" :
                        SendKeys.Send("{ENTER}");
                        break;

                    case "window":
                        txtContents.Text += e.Result.Text.ToString() + Environment.NewLine;
                        break;

                    case "tab":
                        SendKeys.Send("{TAB}");
                        break;

                    case "backspace":
                        SendKeys.Send("{BACKSPACE}");
                        break;

                    case "break":
                        SendKeys.Send("{BREAK}");
                        break;

                    case "capslock":
                        SendKeys.Send("{CAPSLOCK}");
                        break;

                    case "delete":
                        SendKeys.Send("{DELETE}");
                        break;

                    case "down arrow":
                        SendKeys.Send("{DOWN}");
                        break;

                    case "end":
                        SendKeys.Send("{END}");
                        break;

                    case "escape":
                        SendKeys.Send("{ESC}");
                        break;

                    case "help":
                        SendKeys.Send("{HELP}");
                        break;

                    case "home":
                        SendKeys.Send("{HOME}");
                        break;

                    case "insert":
                        SendKeys.Send("{INSERT}");
                        break;

                    case "left arrow":
                        SendKeys.Send("{LEFT}");
                        break;

                    case "num lock":
                        SendKeys.Send("{NUMLOCK}");
                        break;

                    case "page down":
                        SendKeys.Send("{PGDN}");
                        break;

                    case "page up":
                        SendKeys.Send("{PGUP}");
                        break;

                    case "print screen":
                        SendKeys.Send("{PRTSC}");
                        break;

                    case "right arrow":
                        SendKeys.Send("{RIGHT}");
                        break;

                    case "scroll lock":
                        SendKeys.Send("{SCROLLLOCK}");
                        break;

                    case "up arrow":
                        SendKeys.Send("{UP}");
                        break;

                    case "f1":
                        SendKeys.Send("{F1}");
                        break;

                    case "f2":
                        SendKeys.Send("{F2}");
                        break;

                    case "f3":
                        SendKeys.Send("{F3}");
                        break;

                    case "f4":
                        SendKeys.Send("{F4}");
                        break;

                    case "f5":
                        SendKeys.Send("{F5}");
                        break;

                    case "f6":
                        SendKeys.Send("{F6}");
                        break;

                    case "f7":
                        SendKeys.Send("{F7}");
                        break;

                    case "f8":
                        SendKeys.Send("{F8}");
                        break;

                    case "f9":
                        SendKeys.Send("{F9}");
                        break;

                    case "f10":
                        SendKeys.Send("{F10}");
                        break;

                    case "f11":
                        SendKeys.Send("{F11}");
                        break;

                    case "f 12":
                        SendKeys.Send("{F12}");
                        break;

                    case "f13":
                        SendKeys.Send("{F13}");
                        break;

                    case "f14":
                        SendKeys.Send("{F14}");
                        break;

                    case "f15":
                        SendKeys.Send("{F15}");
                        break;

                    case "f16":
                        SendKeys.Send("{F16}");
                        break;
                    //_____________________________________KEY BOARD ENDS!!!_____________________________________
                    /*
                    case "ding":
                        txtContents.Text += e.Result.Text.ToString() + Environment.NewLine;
                        SendKeys.Send("CTRL down");
                        SendKeys.Send("CTRL up");
                        break;*/
                }
                //______________________________________________switch ends_______________________________________
                txtContents.Text += e.Result.Text.ToString() + Environment.NewLine;
                //MessageBox.Show(Environment.UserName);
            }
            //______________________________________________if confidence ENDS!!!_________________________________
        }
        //___________________________________________return process name_________________________________
        // The GetForegroundWindow function returns a handle to the foreground window
        // (the window  with which the user is currently working).
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        // The GetWindowThreadProcessId function retrieves the identifier of the thread
        // that created the specified window and, optionally, the identifier of the
        // process that created the window.
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern Int32 GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // Returns the name of the process owning the foreground window.
        private string GetForegroundProcessName()
        {
            IntPtr hwnd = GetForegroundWindow();

            // The foreground window can be NULL in certain circumstances, 
            // such as when a window is losing activation.
            if (hwnd == null)
                return "Unknown";
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcesses())
            {
                if (p.Id == pid)
                    return p.ProcessName;
            }
            return "Unknown";
        }
        //_________________________________________stop button_______________________________________________
        private void btnstop_Click(object sender, EventArgs e)
        {
            //stop button
            sre.RecognizeAsyncStop();
            btnStart.Enabled = true;
            btnStop.Enabled = false;
        }
        //_________________________________radio button________________________________________________________
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            
        }
    }
}
