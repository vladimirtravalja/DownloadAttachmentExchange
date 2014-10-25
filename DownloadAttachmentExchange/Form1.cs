using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Globalization;
using Microsoft.Exchange;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;

namespace DownloadAttachmentExchange
{
    public partial class Form1 : Form
    {
        private string path = "";
        public string subject = "";
        public string attachment = "";
        public string date = "";
        public string sender = "";
        private string logDir = "";
        private string logName = "";
        private bool automatic = true;

        public Form1()
        {
            InitializeComponent();
        }
        ExchangeClass exchange = new ExchangeClass();
        //button3 represents - getting path for attachment so it could be stored...
        private void button3_Click(object sender, EventArgs e)
        {
            if(folderBrowserDialog1.ShowDialog(this)== DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;   
                exchange.Path = path;
                pathTxt.Text = path;
            }
        }
        //starting manually application for getting attachments
        private void onBtn_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
        }
        //tip for exchange input
        private void exchangeTxt_MouseHover(object sender, EventArgs e)
        {
            tipLbl.Visible = true;
            tipLbl.Text = "It is usually same as email address...";
        }
        //remote tip for exchange input
        private void exchangeTxt_MouseLeave(object sender, EventArgs e)
        {
            tipLbl.Visible = false;
        }
        //another thread for accessing class manualy - for downloading attachment.
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (automatic)
            {
                exchange.GetAttachments(this);
            }
            else
            {
                //this needs to be modified and changed, automation is working flawlessly but some properties for manual version needs to be added and tested.
                exchange.EmailFetch = Convert.ToInt32(fetchUpDown.Value);
                exchange.FilterSender = fromFilterTxt.Text;
                exchange.Exchange = exchangeTxt.Text;
                exchange.Password = passwdTxt.Text;
                exchange.Username = userTxt.Text;
                exchange.LogPath = logDir;
                exchange.LogName = logName;
                exchange.Path = pathTxt.Text;
                exchange.ExchangeType = ExchangeValidation();
                exchange.FilterExtension = comboBox1.Text;
                exchange.GetAttachments(this);
                exchange.MailBoxSizeLimit = Convert.ToInt32(mailboxLimitNumUpDwn.Value);
            }
        }
        //method for manual getting exchange server type
        private string ExchangeValidation()
        {
            string type = "";

            if (exchange07rb.Checked)
            {
                type =  "Exchange2007_SP1";
            }
            else if(exchange10rb.Checked)
            {
                type =  "Exchange2010";
            }
            else if (exchange10sp1rb.Checked)
            {
                type = "Exchange2010_SP1";
            }
            else if (exchange10sp2rb.Checked)
            {
                type = "Exchange2010_SP2";
            }
            else if (exchange13rb.Checked)
            {
                type = "Exchange2013";
            }
            return type;
        }
        //load from appconfig file
        private void Form1_Load(object sender, EventArgs e)
        {
            onBtn.Enabled = false;
            timer1.Interval = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["timer"]);
            comboBox1.SelectedIndex = 1;
            exchange.Username = System.Configuration.ConfigurationManager.AppSettings["username"];
            userTxt.Text = System.Configuration.ConfigurationManager.AppSettings["username"];
            exchange.Password = System.Configuration.ConfigurationManager.AppSettings["password"];
            passwdTxt.Text = System.Configuration.ConfigurationManager.AppSettings["password"];
            exchange.Path = System.Configuration.ConfigurationManager.AppSettings["path"];
            pathTxt.Text = System.Configuration.ConfigurationManager.AppSettings["path"];
            exchange.Exchange = System.Configuration.ConfigurationManager.AppSettings["exchange"];
            exchangeTxt.Text = System.Configuration.ConfigurationManager.AppSettings["exchange"];
            exchange.FilterSender = System.Configuration.ConfigurationManager.AppSettings["fromFilter"];
            fromFilterTxt.Text = System.Configuration.ConfigurationManager.AppSettings["fromFilter"];
            mailboxLimitNumUpDwn.Value = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["mailboxLimit"]);
            exchange.FilterExtension = comboBox1.Text;
            fetchUpDown.Value = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["fetch"]);
            exchange.EmailFetch = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["fetch"]);
            exchange.LogName = System.Configuration.ConfigurationManager.AppSettings["logName"];
            exchange.LogPath = System.Configuration.ConfigurationManager.AppSettings["logPath"];
            logPathTxt.Text = System.Configuration.ConfigurationManager.AppSettings["logPath"] + "\\" + System.Configuration.ConfigurationManager.AppSettings["logName"];
            switch (System.Configuration.ConfigurationManager.AppSettings["exchangeType"])
            {
                case "Exchange2007_SP1":
                    exchange07rb.Checked = true;
                    exchange.ExchangeType = "Exchange2007_SP1";
                    break;
                case "Exchange2010":
                    exchange10rb.Checked = true;
                    exchange.ExchangeType = "Exchange2010";
                    break;
                case "Exchange2010_SP1":
                    exchange10sp1rb.Checked = true;
                    exchange.ExchangeType = "Exchange2010_SP1";
                    break;
                case "Exchange2010_SP2":
                    exchange10sp2rb.Checked = true;
                    exchange.ExchangeType = "Exchange2010_SP2";
                    break;
                case "Exchange2013":
                    exchange13rb.Checked = true;
                    exchange.ExchangeType = "Exchange2013";
                    break;
            }
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            backgroundWorker1.RunWorkerAsync();
            timer1.Start();
        }
        //another event - progress changed for refreshing log output textbox.(if there are many attachment ready for download)
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            logOutputTxt.Text += "[Execution date/time] - " + date + "\t Message subject: " + subject + "\t Attachment name: " + attachment+"\r\n";
            logOutputTxt.SelectionStart = logOutputTxt.Text.Length;
            logOutputTxt.ScrollToCaret();
        }
        //timer used for automatically check and download attachment from exchange server
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            if (!backgroundWorker1.IsBusy) //check if backgroundWorker is bussy, if not run backgroundWorker, else skip to timer start
            {
                backgroundWorker1.RunWorkerAsync();//exchange.GetAttachments(this);
            }
            timer1.Start();
        }
        //filedialog used for setting path for output log file
        private void fileDialogBtn_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                logDir = saveFileDialog1.InitialDirectory;
                logName = saveFileDialog1.FileName;
                logPathTxt.Text = saveFileDialog1.InitialDirectory + "\\" + saveFileDialog1.FileName;
            }
        }
        //close entire application
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
            Dispose();
        }
        //stop backgroundworker and timer
        private void StopBtn_Click(object sender, EventArgs e)
        {
            onBtn.Enabled = true;
            automatic = false;
            timer1.Stop();
            backgroundWorker1.CancelAsync();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            {
                Hide();
                notifyIcon1.ShowBalloonTip(5000, "Attachment Downloader", "Attachment Downloader application is still running, it has been brought to system tray...", ToolTipIcon.Info);
            }
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }

        private void restoreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
            Dispose();
        }
    }
}
