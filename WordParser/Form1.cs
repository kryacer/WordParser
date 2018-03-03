using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;
using xNet;
namespace WordParser
{
    public partial class WordParser : Form
    {
        public WordParser()
        {
            InitializeComponent();
            pBar.Visible = false;
            _cards = new List<Card>();
            _notFound = new List<string>();
        }
        private string _getHtml(string url)
        {
            HttpWebRequest webReq = WebRequest.Create(url) as HttpWebRequest;
                webReq.Method = "GET";
                webReq.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.94 YaBrowser/17.11.1.988 Yowser/2.5 Safari/537.36";
            try
            {
                using (WebResponse webRes = webReq.GetResponse())
                {
                    Stream st = webRes.GetResponseStream();
                    StreamReader sr = new StreamReader(st, Encoding.UTF8);
                    string html = sr.ReadToEnd();
                    st.Close();
                    sr.Close();
                    return html;
                }
            }
            catch
            {
                return null;
            }
            
        }
        private IList<Card> _cards;
        private List<string> _notFound;
        private void _export()
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Workbooks.Add(Type.Missing);
            //ExcelApp.Cells[1, 1] = "Front";
            //ExcelApp.Cells[1, 2] = "Transcript";
            //ExcelApp.Cells[1, 3] = "Back";
            //ExcelApp.Cells[1, 4] = "Image";
            //ExcelApp.Cells[1, 5] = "Examples";
            //ExcelApp.Cells[1, 6] = "AudioMary";
            //ExcelApp.Cells[1, 7] = "AudioHunt";
            //ExcelApp.Cells[1, 8] = "NotFound";
            int j = 1;
            foreach(var s in _cards)
            {   
                ExcelApp.Cells[j, 1] = s.Front;
                ExcelApp.Cells[j, 2] = s.Transcript;
                ExcelApp.Cells[j, 3] = s.Back;
                ExcelApp.Cells[j, 4] = s.Image;
                ExcelApp.Cells[j, 5] = s.Examples;
                ExcelApp.Cells[j, 6] = s.AudioMary;
                ExcelApp.Cells[j, 7] = s.AudioHunt;
                
                j++;
            }
            for(int i=0;i<_notFound.Count; i++)
            {
                ExcelApp.Cells[i+1, 8] = _notFound[i];
            }
            ExcelApp.Visible = true;
        }
        
        private string _downloadAudio(string url, string fileName)
        {
            using (WebClient client = new WebClient())
            {
                byte[] bytes = (!url.Contains("adhoc01w"))?client.DownloadData(url):null;
                if (bytes !=null&& !Encoding.UTF8.GetString(bytes).Contains("Упссс"))
                {
                    File.WriteAllBytes(@"\audio\" + fileName, bytes);
                    //client.DownloadFile(new Uri(url), @"\audio\" + fileName);
                    return "[sound:"+fileName+"]";
                }
                else
                {
                    return "";
                }
                //OR 
                //client.DownloadFileAsync(new Uri(url), @"c:\temp\image35.png");
            }
        }
        private string _wordHuntUrl = "http://wooordhunt.ru/word/";
        private string _merriamUrl = "https://www.merriam-webster.com/dictionary/";
        private string _wordHuntAudioUrl = "http://wooordhunt.ru/data/sound/word/us/ogg/"; //.ogg needed
        private string _merriamAudioUrl = "http://media.merriam-webster.com/soundc11/"; //letter/fileName.wav
        
        private void btnParse_Click(object sender, EventArgs e)
        {
            _cards.Clear();
            _notFound.Clear();
            string filename = "words.txt";
            string txt = File.ReadAllText(@"\audio\" + filename);
            string[] words = txt.Split(new string[] { Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);
            for(int i =0; i<words.Length; i++)
            {
                words[i] = words[i].Trim().ToLower().Replace(".","").Replace("�", "");
            }
            pBar.Minimum = 0;
            pBar.Maximum = words.Length;
            pBar.Value = 0;
            pBar.Step = 1;
            pBar.Visible = true;
            for (int i=0; i<words.Length; i++)
            {
                string huntHtml = _getHtml(_wordHuntUrl + words[i]);
                string maryHtml = _getHtml(_merriamUrl + words[i]);
                string maryHword = maryHtml.Substring("<h1 class=\"hword\">", "</h1>").ToLower().Replace(".","");
                string huntBack = huntHtml.Substring("t_inline_en\">", "</span>");
                string transcript = huntHtml.Substring("transcription\"> ", "</span>"); //transcript
                if(maryHtml ==null || maryHword !=words[i] || transcript == "") {_notFound.Add(words[i]); continue; }
                string[] huntExamples = huntHtml.Substrings("ex_o\">", "<span"); //examples
                string examples = "";
                for(int q=0; q<huntExamples.Length; q++)
                {
                    examples += huntExamples[q] + Environment.NewLine;
                }
                string marySL = maryHtml.Substring("simple-learners", "data-source=\"elementary\"");
                string[] ems = marySL.Substrings("<em>", "</em>");
                string[] defs = marySL.Substrings("definition-block def-text", "</div>");
                string maryBack = "";
                for(int q=0; q<defs.Length; q++)
                {
                    maryBack += ems[q] + Environment.NewLine;
                    string[] items = defs[q].Substrings("definition-inner-item\">", "</p>");
                    for(int j = 0; j < items.Length; j++)
                    {
                       items[j] =  items[j].Replace("<span>", "").Replace("<span class=\"intro-colon\">", "").Replace("</span>", "").Trim().TrimEnd('\r', '\n');
                       maryBack += items[j] + Environment.NewLine;
                    }
                }
                string back = huntBack + Environment.NewLine + maryBack; //back final
                string maryFileName = maryHtml.Substring("data-file=\"", "\"");//for mary audio
                string maryDir = maryHtml.Substring("data-dir=\"", "\"");//for mary audio
                Card card = new Card();
                card.Front = words[i];
                card.Transcript = transcript;
                card.Back = back;
                card.Image = "";
                card.Examples = examples;
                card.AudioMary = _downloadAudio(_merriamAudioUrl + maryDir + "/" + maryFileName + ".wav", words[i] + " mary" + ".wav");
                card.AudioHunt = _downloadAudio(_wordHuntAudioUrl + words[i] + ".ogg", words[i] + " hunt" + ".ogg");
                _cards.Add(card);
                pBar.Value = i;
            }
            _export();
        }
    }
}
