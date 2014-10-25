namespace DownloadAttachmentExchange
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.exchange13rb = new System.Windows.Forms.RadioButton();
            this.exchange10sp2rb = new System.Windows.Forms.RadioButton();
            this.exchange10sp1rb = new System.Windows.Forms.RadioButton();
            this.exchange10rb = new System.Windows.Forms.RadioButton();
            this.exchange07rb = new System.Windows.Forms.RadioButton();
            this.label9 = new System.Windows.Forms.Label();
            this.logPathTxt = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.fileDialogBtn = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.fetchUpDown = new System.Windows.Forms.NumericUpDown();
            this.tipLbl = new System.Windows.Forms.Label();
            this.fromFilterTxt = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.pathTxt = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.exchangeTxt = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.passwdTxt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.userTxt = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.onBtn = new System.Windows.Forms.Button();
            this.logOutputTxt = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.cancelBtn = new System.Windows.Forms.Button();
            this.stopBtn = new System.Windows.Forms.Button();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.restoreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mailboxLimitLbl = new System.Windows.Forms.Label();
            this.mailboxLimitNumUpDwn = new System.Windows.Forms.NumericUpDown();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fetchUpDown)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.mailboxLimitNumUpDwn)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.mailboxLimitLbl);
            this.groupBox1.Controls.Add(this.mailboxLimitNumUpDwn);
            this.groupBox1.Controls.Add(this.exchange13rb);
            this.groupBox1.Controls.Add(this.exchange10sp2rb);
            this.groupBox1.Controls.Add(this.exchange10sp1rb);
            this.groupBox1.Controls.Add(this.exchange10rb);
            this.groupBox1.Controls.Add(this.exchange07rb);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.logPathTxt);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.fileDialogBtn);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.fetchUpDown);
            this.groupBox1.Controls.Add(this.tipLbl);
            this.groupBox1.Controls.Add(this.fromFilterTxt);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.pathTxt);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.exchangeTxt);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.passwdTxt);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.userTxt);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(16, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(636, 206);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Server settings";
            // 
            // exchange13rb
            // 
            this.exchange13rb.AutoSize = true;
            this.exchange13rb.Location = new System.Drawing.Point(495, 139);
            this.exchange13rb.Name = "exchange13rb";
            this.exchange13rb.Size = new System.Drawing.Size(100, 17);
            this.exchange13rb.TabIndex = 37;
            this.exchange13rb.TabStop = true;
            this.exchange13rb.Text = "Exchange 2013";
            this.exchange13rb.UseVisualStyleBackColor = true;
            // 
            // exchange10sp2rb
            // 
            this.exchange10sp2rb.AutoSize = true;
            this.exchange10sp2rb.Location = new System.Drawing.Point(495, 117);
            this.exchange10sp2rb.Name = "exchange10sp2rb";
            this.exchange10sp2rb.Size = new System.Drawing.Size(123, 17);
            this.exchange10sp2rb.TabIndex = 36;
            this.exchange10sp2rb.TabStop = true;
            this.exchange10sp2rb.Text = "Exchange 2010 SP2";
            this.exchange10sp2rb.UseVisualStyleBackColor = true;
            // 
            // exchange10sp1rb
            // 
            this.exchange10sp1rb.AutoSize = true;
            this.exchange10sp1rb.Location = new System.Drawing.Point(495, 94);
            this.exchange10sp1rb.Name = "exchange10sp1rb";
            this.exchange10sp1rb.Size = new System.Drawing.Size(123, 17);
            this.exchange10sp1rb.TabIndex = 35;
            this.exchange10sp1rb.TabStop = true;
            this.exchange10sp1rb.Text = "Exchange 2010 SP1";
            this.exchange10sp1rb.UseVisualStyleBackColor = true;
            // 
            // exchange10rb
            // 
            this.exchange10rb.AutoSize = true;
            this.exchange10rb.Location = new System.Drawing.Point(495, 71);
            this.exchange10rb.Name = "exchange10rb";
            this.exchange10rb.Size = new System.Drawing.Size(100, 17);
            this.exchange10rb.TabIndex = 34;
            this.exchange10rb.TabStop = true;
            this.exchange10rb.Text = "Exchange 2010";
            this.exchange10rb.UseVisualStyleBackColor = true;
            // 
            // exchange07rb
            // 
            this.exchange07rb.AutoSize = true;
            this.exchange07rb.Location = new System.Drawing.Point(495, 48);
            this.exchange07rb.Name = "exchange07rb";
            this.exchange07rb.Size = new System.Drawing.Size(100, 17);
            this.exchange07rb.TabIndex = 33;
            this.exchange07rb.TabStop = true;
            this.exchange07rb.Text = "Exchange 2007";
            this.exchange07rb.UseVisualStyleBackColor = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(392, 50);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(90, 13);
            this.label9.TabIndex = 32;
            this.label9.Text = "Exchange server:";
            // 
            // logPathTxt
            // 
            this.logPathTxt.Location = new System.Drawing.Point(117, 90);
            this.logPathTxt.Name = "logPathTxt";
            this.logPathTxt.ReadOnly = true;
            this.logPathTxt.Size = new System.Drawing.Size(301, 20);
            this.logPathTxt.TabIndex = 31;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(22, 93);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(67, 13);
            this.label8.TabIndex = 30;
            this.label8.Text = "Set log path:";
            // 
            // fileDialogBtn
            // 
            this.fileDialogBtn.Location = new System.Drawing.Point(424, 89);
            this.fileDialogBtn.Name = "fileDialogBtn";
            this.fileDialogBtn.Size = new System.Drawing.Size(31, 20);
            this.fileDialogBtn.TabIndex = 29;
            this.fileDialogBtn.Text = "...";
            this.fileDialogBtn.UseVisualStyleBackColor = true;
            this.fileDialogBtn.Click += new System.EventHandler(this.fileDialogBtn_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            ".csv",
            ".xls"});
            this.comboBox1.Location = new System.Drawing.Point(117, 160);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(80, 21);
            this.comboBox1.TabIndex = 28;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(392, 26);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(37, 13);
            this.label7.TabIndex = 17;
            this.label7.Text = "Fetch:";
            // 
            // fetchUpDown
            // 
            this.fetchUpDown.Location = new System.Drawing.Point(435, 23);
            this.fetchUpDown.Name = "fetchUpDown";
            this.fetchUpDown.Size = new System.Drawing.Size(47, 20);
            this.fetchUpDown.TabIndex = 16;
            // 
            // tipLbl
            // 
            this.tipLbl.AutoSize = true;
            this.tipLbl.ForeColor = System.Drawing.Color.Firebrick;
            this.tipLbl.Location = new System.Drawing.Point(319, 72);
            this.tipLbl.Name = "tipLbl";
            this.tipLbl.Size = new System.Drawing.Size(35, 13);
            this.tipLbl.TabIndex = 14;
            this.tipLbl.Text = "label7";
            this.tipLbl.Visible = false;
            // 
            // fromFilterTxt
            // 
            this.fromFilterTxt.Location = new System.Drawing.Point(117, 114);
            this.fromFilterTxt.Name = "fromFilterTxt";
            this.fromFilterTxt.Size = new System.Drawing.Size(198, 20);
            this.fromFilterTxt.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(22, 117);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(55, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "From filter:";
            // 
            // pathTxt
            // 
            this.pathTxt.Location = new System.Drawing.Point(117, 137);
            this.pathTxt.Name = "pathTxt";
            this.pathTxt.ReadOnly = true;
            this.pathTxt.Size = new System.Drawing.Size(249, 20);
            this.pathTxt.TabIndex = 11;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(372, 137);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(32, 21);
            this.button3.TabIndex = 10;
            this.button3.Text = "...";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(22, 140);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Save file to:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(22, 163);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Filter extnsion:";
            // 
            // exchangeTxt
            // 
            this.exchangeTxt.Location = new System.Drawing.Point(117, 68);
            this.exchangeTxt.Name = "exchangeTxt";
            this.exchangeTxt.Size = new System.Drawing.Size(198, 20);
            this.exchangeTxt.TabIndex = 5;
            this.exchangeTxt.MouseLeave += new System.EventHandler(this.exchangeTxt_MouseLeave);
            this.exchangeTxt.MouseHover += new System.EventHandler(this.exchangeTxt_MouseHover);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(90, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Exchange server:";
            // 
            // passwdTxt
            // 
            this.passwdTxt.Location = new System.Drawing.Point(117, 45);
            this.passwdTxt.Name = "passwdTxt";
            this.passwdTxt.PasswordChar = '*';
            this.passwdTxt.Size = new System.Drawing.Size(198, 20);
            this.passwdTxt.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Password:";
            // 
            // userTxt
            // 
            this.userTxt.Location = new System.Drawing.Point(117, 22);
            this.userTxt.Name = "userTxt";
            this.userTxt.Size = new System.Drawing.Size(198, 20);
            this.userTxt.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Username / email:";
            // 
            // onBtn
            // 
            this.onBtn.Location = new System.Drawing.Point(573, 396);
            this.onBtn.Name = "onBtn";
            this.onBtn.Size = new System.Drawing.Size(78, 23);
            this.onBtn.TabIndex = 0;
            this.onBtn.Text = "Download";
            this.onBtn.UseVisualStyleBackColor = true;
            this.onBtn.Click += new System.EventHandler(this.onBtn_Click);
            // 
            // logOutputTxt
            // 
            this.logOutputTxt.Location = new System.Drawing.Point(16, 225);
            this.logOutputTxt.Multiline = true;
            this.logOutputTxt.Name = "logOutputTxt";
            this.logOutputTxt.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.logOutputTxt.Size = new System.Drawing.Size(636, 165);
            this.logOutputTxt.TabIndex = 2;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            // 
            // timer1
            // 
            this.timer1.Interval = 300000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // cancelBtn
            // 
            this.cancelBtn.Location = new System.Drawing.Point(21, 396);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(75, 23);
            this.cancelBtn.TabIndex = 3;
            this.cancelBtn.Text = "Cancel";
            this.cancelBtn.UseVisualStyleBackColor = true;
            this.cancelBtn.Click += new System.EventHandler(this.cancelBtn_Click);
            // 
            // stopBtn
            // 
            this.stopBtn.Location = new System.Drawing.Point(492, 396);
            this.stopBtn.Name = "stopBtn";
            this.stopBtn.Size = new System.Drawing.Size(75, 23);
            this.stopBtn.TabIndex = 4;
            this.stopBtn.Text = "Stop";
            this.stopBtn.UseVisualStyleBackColor = true;
            this.stopBtn.Click += new System.EventHandler(this.StopBtn_Click);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.ContextMenuStrip = this.contextMenuStrip1;
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "Attachment Downloader";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.DoubleClick += new System.EventHandler(this.notifyIcon1_DoubleClick);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.restoreToolStripMenuItem,
            this.closeToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(114, 48);
            // 
            // restoreToolStripMenuItem
            // 
            this.restoreToolStripMenuItem.Name = "restoreToolStripMenuItem";
            this.restoreToolStripMenuItem.Size = new System.Drawing.Size(113, 22);
            this.restoreToolStripMenuItem.Text = "Restore";
            this.restoreToolStripMenuItem.Click += new System.EventHandler(this.restoreToolStripMenuItem_Click);
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(113, 22);
            this.closeToolStripMenuItem.Text = "Close";
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
            // 
            // mailboxLimitLbl
            // 
            this.mailboxLimitLbl.AutoSize = true;
            this.mailboxLimitLbl.Location = new System.Drawing.Point(392, 179);
            this.mailboxLimitLbl.Name = "mailboxLimitLbl";
            this.mailboxLimitLbl.Size = new System.Drawing.Size(69, 13);
            this.mailboxLimitLbl.TabIndex = 39;
            this.mailboxLimitLbl.Text = "Mail box limit:";
            // 
            // mailboxLimitNumUpDwn
            // 
            this.mailboxLimitNumUpDwn.Location = new System.Drawing.Point(464, 176);
            this.mailboxLimitNumUpDwn.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.mailboxLimitNumUpDwn.Name = "mailboxLimitNumUpDwn";
            this.mailboxLimitNumUpDwn.Size = new System.Drawing.Size(69, 20);
            this.mailboxLimitNumUpDwn.TabIndex = 38;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(669, 431);
            this.Controls.Add(this.stopBtn);
            this.Controls.Add(this.cancelBtn);
            this.Controls.Add(this.logOutputTxt);
            this.Controls.Add(this.onBtn);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Attachment Downloader";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fetchUpDown)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.mailboxLimitNumUpDwn)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button onBtn;
        private System.Windows.Forms.TextBox exchangeTxt;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox passwdTxt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox userTxt;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox pathTxt;
        private System.Windows.Forms.TextBox fromFilterTxt;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label tipLbl;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        public System.Windows.Forms.TextBox logOutputTxt;
        public System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.NumericUpDown fetchUpDown;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBox1;
        public System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.TextBox logPathTxt;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button fileDialogBtn;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.RadioButton exchange07rb;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.RadioButton exchange13rb;
        private System.Windows.Forms.RadioButton exchange10sp2rb;
        private System.Windows.Forms.RadioButton exchange10sp1rb;
        private System.Windows.Forms.RadioButton exchange10rb;
        private System.Windows.Forms.Button cancelBtn;
        private System.Windows.Forms.Button stopBtn;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem restoreToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeToolStripMenuItem;
        private System.Windows.Forms.Label mailboxLimitLbl;
        private System.Windows.Forms.NumericUpDown mailboxLimitNumUpDwn;
    }
}

