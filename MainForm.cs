using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;

namespace JARVISCS
{
    public partial class MainForm : Form
    {

        bool search;
        SpeechRecognitionEngine speechRecognitionEngine = null;
        SpeechSynthesizer Jarvis = new SpeechSynthesizer();

        

        public static List<string> MsgList = new List<string>();
        public static List<string> MsgLink = new List<string>();
        // QEvent for check new emails
        public static String QEvent;
 
        public MainForm()
        {
            InitializeComponent();
            try
            {
               
                // Set the language for speech engine
                speechRecognitionEngine = SetLanguage("en-US");
                //Event handler for recognized text 
                speechRecognitionEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Mainevent_SpeechRecognized);
                //Event for load grammar for speech engine 
                LoadGrammarAndCommands();
                // Using the system's default microphone
                speechRecognitionEngine.SetInputToDefaultAudioDevice();
                // Start listening 
                speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
                VoiceGender voiceGender = VoiceGender.NotSet;

                voiceGender = VoiceGender.Female;

                Jarvis.SelectVoiceByHints(voiceGender);

                //welcome message
                Jarvis.Volume = 100;
                Jarvis.Speak("Welcome to Jarvis Application, Version one point oh!");
                Jarvis.Speak("Welcome back, Sir");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void LoadGrammarAndCommands()
        {
            try
            {
                string connectionstring = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionstring);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "SELECT * FROM DefaultTable";
                //sc.CommandType = CommandType.TableDirect;
                SqlDataReader sdr = sc.ExecuteReader();
                while (sdr.Read())
                {
                    var Loadcmd = sdr["Commands"].ToString();
                    Grammar commandgrammar = new Grammar(new GrammarBuilder(new Choices(Loadcmd)));
                    speechRecognitionEngine.LoadGrammarAsync(commandgrammar);

                }
                sdr.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your web commands, possibly a blank line. web commands will case to work until it is fixed." + ex.Message);
            }
        }

        //Adding Speech Recognition
        private void Mainevent_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string Name = Environment.UserName;
            //Recognized Spoken words result is e.Result.Text
            string speech = e.Result.Text;
          
            if (speech == "search for")
            {
                search = true;
                Jarvis.Speak("Tell me what to search for");

            }
            if (search)
            {
                Process.Start("https://www.google.com/search?q=" + speech);
                search = false;
            }
            if (search == false)
            {
                switch (speech)
                {
                    //Greetings
                    case "Hey Jarvis":
                    case "Hello Jarvis":
                        Jarvis.SpeakAsync("hello " + Name);
                        System.DateTime timenow = System.DateTime.Now;
                        if (timenow.Hour >= 5 && timenow.Hour < 12)
                        {
                            Jarvis.SpeakAsync("Goodmorning " + Name);
                        }
                        if (timenow.Hour >= 12 && timenow.Hour < 18)
                        {
                            Jarvis.SpeakAsync("Good afternoon " + Name);
                        }
                        if (timenow.Hour >= 18 && timenow.Hour < 24)
                        { Jarvis.SpeakAsync("Good evening " + Name); }
                        if (timenow.Hour < 5)
                        { Jarvis.SpeakAsync("Hello " + Name + ", you are still awake you should go to sleep, it's getting late"); }
                        break;
                    case "what time is it":
                        System.DateTime now = System.DateTime.Now;
                        string time = now.GetDateTimeFormats('t')[0];
                        Jarvis.SpeakAsync(time);
                        break;
                    case "what day is it":
                        string day = "Today is," + System.DateTime.Now.ToString("dddd");
                        Jarvis.SpeakAsync(day);
                        break;
                    case "what is the date":
                        string date = "The date is, " + System.DateTime.Now.ToString("dd MMM");
                        Jarvis.SpeakAsync(date);
                        date = "" + System.DateTime.Today.ToString(" yyyy");
                        Jarvis.SpeakAsync(date);
                        break;
                    case "who are you":
                        Jarvis.SpeakAsync("i am your personal assistant");
                        Jarvis.SpeakAsync("i can read email, weather report, i can search web for you, anything that you need like a personal assistant do, you can ask me question i will reply to you");
                        break;
                    case "what is my name":
                        Jarvis.SpeakAsync(Name);
                        break;
                    case "start reading":
                        Readbtn.PerformClick();
                        if (tabControl1.SelectedIndex == 1)
                        {
                            Jarvis.SpeakAsync(Readtxt.Text);
                        }
                        if (tabControl1.SelectedIndex == 2)
                        {
                            Jarvis.SpeakAsync(convertedtxt.Text);
                        }

                        break;
                    case "pause":
                        //PauseBtn.PerformClick();
                        Jarvis.Pause();
                        break;
                    case "resume":
                        Pausebtn.PerformClick();

                        if (Jarvis.State == SynthesizerState.Speaking)
                            Jarvis.Resume();

                        if (Jarvis.State == SynthesizerState.Speaking)
                            Jarvis.Resume();
                        break;
                    case "stop":
                        Stopbtn.PerformClick();
                        break;
                    case "Open Weather Reader":
                        tabControl1.SelectedIndex = 3;
                        break;
                    case "open text file":
                        Open_file.PerformClick();
                        break;

                    case "open text reader":
                        Jarvis.Speak("Here you go, Sir");
                        this.tabControl1.SelectedIndex = 1;
                        break;
                    case "open news reader":
                        Jarvis.Speak("Here you go, Sir");
                        this.tabControl1.SelectedIndex = 2;
                        break;

                    case "thank you jarvis":
                        Jarvis.SpeakAsync("Always at your service, Sir");
                        break;
                    case "open google":
                        Jarvis.Speak("opening google");
                        System.Diagnostics.Process.Start("https://www.google.com");
                        break;
                    case "open notepad":
                        Jarvis.Speak("opening notepad");
                        System.Diagnostics.Process.Start("notepad.exe");
                        break;
                    case "Get Full Screen":
                    case "full screen":
                        WindowState = FormWindowState.Maximized;
                        break;

                    case "Open Commands Input":
                    case "Commands Input":
                        this.tabControl1.SelectedIndex = 0;
                        break;

                    case "Get noemal Screen":
                    case "normal":
                    case "one":
                    case "original screen":
                        WindowState = FormWindowState.Normal;
                        break;
                    case "Mininmized":
                        WindowState = FormWindowState.Minimized;
                        break;

                    case "kill this":

                        break;
                    case "Open Word":
                        Process.Start("winword");
                        break;
                    case "bye Jarvis":
                        Jarvis.Speak("Bye Sir, may you have a great day ahead");
                        SendKeys.Send("%{F4}");
                        break;
                    case "open browser":
                        //    WindowState = FormWindowState.Minimized;
                        Jarvis.Speak("Here you go, Sir");
                        this.tabControl1.SelectedIndex = 4;

                        break;

                    case "Whats the weather like":
                        GetWeather();
                        Jarvis.Speak("according to Google, ");

                        break;

                    case "Whats is the latest news":
                        GetBingNews();
                        break;
                    case "get bing news":
                        GetBingNews();
                        break;
                }
            }
            Debug_Livetxt.AppendText(speech + '\n');

        }

        private SpeechRecognitionEngine SetLanguage(string preferredCulture)
        {
            //Checking for installed language and comparing with our given parameter preferredCulture to set speech recognition engine language
            foreach (RecognizerInfo config in SpeechRecognitionEngine.InstalledRecognizers())
            {
                if (config.Culture.ToString() == preferredCulture)
                {
                    speechRecognitionEngine = new SpeechRecognitionEngine(config);
                    break;
                }
            }

            // if the desired culture is not found, then load default
            if (speechRecognitionEngine == null)
            {
                MessageBox.Show("The desired languages is not installed on this machine, the speech-engine will continue using "
                    + SpeechRecognitionEngine.InstalledRecognizers()[0].Culture.ToString() + " as the default language.",
                    "Culture " + preferredCulture + " not found!");
                speechRecognitionEngine = new SpeechRecognitionEngine(SpeechRecognitionEngine.InstalledRecognizers()[0]);
            }

            return speechRecognitionEngine;
        }

        private void LeftSideButton_Click(object sender, EventArgs e)
        {
            if (LeftSideMenu.Width == 50)
            {
                LeftSideMenu.Width = 260;
                Persona_Assistant.Text = "Personal Assistant";


            }
            else
            {
                LeftSideMenu.Width = 50;
                Persona_Assistant.Text = "PA";

            }
        }

        private void TopPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LeftSideMenu.Width = 50;
            RightSideMenu.Width = 50;
            Persona_Assistant.Text = "PA";
            Debug_Livetxt.SelectionIndent += 20;
        }

        private void RightMenubtn_Click(object sender, EventArgs e)
        {
            if (RightSideMenu.Width == 50)
            {
                RightSideMenu.Width = 260;

            }
            else
            {
                RightSideMenu.Width = 50;

            }
        }
       

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

    
        //For Reading Text fro file
        private void Open_file_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
            Readtxt.Clear();
            Jarvis.SpeakAsync("choose a text file from your, drives");

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //Filter for type of the file we are going to open
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|rtf files (*.rtf)|*.rtf|All files (*.*)|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    //FileNames Gets the file names of all selected files in the dialog box
                    string strfilename = openFileDialog1.FileName;
                    //Here we are using System.IO class to reads the lines of a file
                    string filetext = File.ReadAllText(strfilename);
                    //than we are going to pass the filetxt to over textbox which is Readtxt.Text
                    Readtxt.Text = filetext;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void Readbtn_Click(object sender, EventArgs e)
        {

            if(Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
            Readbtn.Enabled = false;
            Pausebtn.Enabled = true;
            if (tabControl1.SelectedIndex == 2)
            {
                Jarvis.SpeakAsync(Readtxt.Text);
            }
            if (tabControl1.SelectedIndex == 3)
            {
                Jarvis.SpeakAsync(convertedtxt.Text);
            }
        }

        //NEWS SECTION

        public void GetBingNews()
        {
            string checkinternet = NetworkInterface.GetIsNetworkAvailable().ToString();
            if (checkinternet != "True")
            {
                Jarvis.SpeakAsync("Please check your internet connection, before the news broadcast panel, work properly");
            }
            else
            {
                Jarvis.SpeakAsync("todays latest news is");
                convertedtxt.Clear();
                //it is common methods for sending data to and receiving data from a resource identified by a URI.
                WebClient webclient = new WebClient();
                // Downloads the requested resource as a String. The resource to download is specified as a String containing the URI.
                string page = webclient.DownloadString("https://www.bing.com/news/search?q=World&nvaug=%5bNewsVertical+Category%3d%22rt_World%22%5d&FORM=NSBABR");
                webBrowser2.Navigate("https://www.bing.com/news/search?q=World&nvaug=%5bNewsVertical+Category%3d%22rt_World%22%5d&FORM=NSBABR");
                //than we parse the html div tag and we will store in string variable
                string news = "<div class=\"snippet\">(.*?)</div>";
                //Searches the specified input string for the first occurrence of the regular expression specified in the Regex constructor.
                foreach (Match match in Regex.Matches(page, news))
                {
                    //Gets a collection of groups matched by the regular expression
                    convertedtxt.Text += match.Groups[1].Value;
                }
               // ReadBtn.PerformClick();
            }
        }

        private void Pausebtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
            {
                Jarvis.Pause();
                Pausebtn.Text = "Resume";
            }
            else
            {
                Jarvis.Resume();
                Pausebtn.Text = "Pause";
            }
        }

        private void Stopbtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string connectionstring = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionstring);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "INSERT INTO Commands(commad, response) VALUES (@commad,@response)";
                sc.Parameters.Add("@commad", textBox1.Text);
                sc.Parameters.Add("@response", textBox2.Text);
                sc.ExecuteNonQuery();
                con.Close();
                Jarvis.SpeakAsync("Changes have been successfully saved");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "Commands are not saved");
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }


        //WEATHER SECTION

        public void GetWeather()
        {
            Process.Start("https://www.accuweather.com/en/pk/karachi/261158/weather-forecast/261158");
        }
        //For Web Browser

        private void button4_Click(object sender, EventArgs e)
        {
            webBrowser1.GoForward();
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                webBrowser1.Navigate(textBox3.Text);

            }
        }

        private void back_Click(object sender, EventArgs e)
        {
            webBrowser1.GoBack();
        }

        private void refersh_Click(object sender, EventArgs e)
        {
            webBrowser1.Refresh();
        }

        private void search_btn_Click(object sender, EventArgs e)
        {
         
                webBrowser1.Navigate(textBox3.Text);

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }
    }
}
