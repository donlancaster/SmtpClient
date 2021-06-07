using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.Threading;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace quickmailer
{
    public partial class Form1 : Form
    {
        private Workbook eWorkBook;
        private Worksheet eWorkSheet;
        private Microsoft.Office.Interop.Excel.Application excelApplication;
        private static MailMessage[] mailMessages;
        private Mailer[] mailers;
        private NetworkCredential networkCredential;
        private long mail_count = 0;
        const long MAX_COUNT = 20;
        private StringBuilder attachmentsPaths = new StringBuilder();


        public static void DeleteMessage(int index)
        {
            mailMessages[index].Dispose();
        }

        public Form1()
        {
            InitializeComponent();
            SetCbNameElements();
            txtUsername.Text = "";
            txtPassword.Text = "";
            txtSubject.Text = "";
            networkCredential = new NetworkCredential(txtUsername.Text, txtPassword.Text);
            mailers = new Mailer[MAX_COUNT];
            mailMessages = new MailMessage[MAX_COUNT];
        }

        private void GetCredentials()
        {
            ofd.Filter = "Credentials|*.txt";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StreamReader sr;
                sr = new StreamReader(ofd.FileName);
                networkCredential.UserName = sr.ReadLine();
                networkCredential.Password = sr.ReadLine();

                txtUsername.Text = networkCredential.UserName;
                txtPassword.Text = networkCredential.Password;
                sr.Close();


            }
        }


        private void SetCbNameElements()
        {
            cbName.Items.Clear();
            StreamReader streamReader = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "names.txt");
            string name;
            while ((name = streamReader.ReadLine()) != null)
            {
                if (name.Trim() != "")
                    cbName.Items.Add(name);
            }
            streamReader.Close();
            cbName.SelectedIndex = 0;
        }


        private void btnAddAttachments_Click(object sender, EventArgs e)
        {
            ofd.Filter = "ALL|*";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
              //  attachmentsPaths.Append(ofd.FileName);
                //attachmentsPaths.Append("\r\n");
                //Console.WriteLine(attachmentsPaths.ToString());
                listAttachments.Items.Add(ofd.FileName + "\r\n");
            }
        }


        private void mnuOpenSource_Click(object sender, EventArgs e)
        {
            ofd.Filter = "Excel 2007|*.xlsx|Excel 2003|*.xls";

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                eWorkBook = excelApplication.Workbooks.Open(ofd.FileName);
                eWorkSheet = eWorkBook.Worksheets[1];
                int rs;

                mainlist.Items.Clear();
                ListViewItem listView;
                rs = 2;
                while (eWorkSheet.Cells[rs, 1].Value != null)
                {
                    String receiverName = (string)eWorkSheet.Cells[rs, 2].Value;
                    listView = mainlist.Items.Add(receiverName);

                    listView.SubItems.Add((string)eWorkSheet.Cells[rs, 3].Value);
                    Console.WriteLine("3 " + (string)eWorkSheet.Cells[rs, 3].Value);
                    listView.SubItems.Add((string)eWorkSheet.Cells[rs, 4].Value);
                    Console.WriteLine("4 " + (string)eWorkSheet.Cells[rs, 4].Value);
                    listView.SubItems.Add((string)eWorkSheet.Cells[rs, 5].Value);
                    Console.WriteLine("5 " + (string)eWorkSheet.Cells[rs, 5].Value);
                    listView.SubItems.Add((string)eWorkSheet.Cells[rs, 6].Value);
                    Console.WriteLine("6 " + (string)eWorkSheet.Cells[rs, 6].Value);
                    listView.SubItems.Add((string)eWorkSheet.Cells[rs, 7].Value);
                    Console.WriteLine("7 " + (string)eWorkSheet.Cells[rs, 7].Value);
                    listView.SubItems.Add("");

                    rs++;
                }

                eWorkBook.Close();
                excelApplication.Quit();

            }
        }

        private void mnuSend_Click(object sender, EventArgs e)
        {
            string s_message;
            string[] s_emails;
            s_message = (File.OpenText(txtContent.Text)).ReadToEnd();
            Console.WriteLine(s_message);

            mail_count = 0;
            foreach (ListViewItem item in mainlist.Items)
            {
                if (item.Checked)
                {
                    /*if (!File.Exists(lv.SubItems[2].Text))
                    {
                        lv.SubItems[3].Text = "No attachment!";
                        continue;
                    }*/

                    if (item.SubItems[5].Text.Length == 0)
                    {
                        item.SubItems[5].Text = "No email address!";
                        continue;
                    }


                    s_emails = item.SubItems[5].Text.Split(',', ';');




                    foreach (string to in s_emails)
                    {
                        item.SubItems[6].Text = "Sending...";
                        mailers[mail_count] = new Mailer(networkCredential);
                        mailMessages[mail_count] = new MailMessage(txtEmail.Text, to);
                        mailMessages[mail_count].From = new MailAddress(txtEmail.Text, cbName.Text);

                        mailMessages[mail_count].Body = s_message;

                        mailMessages[mail_count].Subject = txtSubject.Text + " FOR " + item.Text.ToUpper();
                        mailMessages[mail_count].IsBodyHtml = true;
                 
                        if (listAttachments.Items.Count != 0)
                        {
                            foreach (ListViewItem attachmentPath in listAttachments.Items)
                            {
                                Console.WriteLine("attach        =      "+attachmentPath.Text);
                                mailMessages[mail_count].Attachments.Add(new Attachment(attachmentPath.Text.Substring(0,attachmentPath.Text.Length-2)));
                               
                            }
                        }
                        Console.WriteLine(mailMessages[mail_count].Body);
                        ListViewX sending = new ListViewX();
                        sending.item = item;
                        sending.rowindex = (int)mail_count;
                        mailers[mail_count].SendMessage(mailMessages[mail_count], sending);
                        mail_count++;
                    }
                }
            }
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            networkCredential.UserName = txtUsername.Text;
            networkCredential.Password = txtPassword.Text;
            MessageBox.Show("Credentials changed!");
        }

        private void txtContent_Click(object sender, EventArgs e)
        {
            ofd.Filter = "HTML|*.html|HTM|*.htm";

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtContent.Text = ofd.FileName;
            }
        }

        private void cbName_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cbName.SelectedIndex)
            {
                case 0:
                    txtEmail.Text = "";
                    break;

                case 1:
                    txtEmail.Text = "";
                    break;
            }
        }

        private void mnuSelectAll_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in mainlist.Items)
            {
                item.Checked = true;
            }
        }

        private void btnGetCredentials_Click(object sender, EventArgs e)
        {
            GetCredentials();
        }

        private void btnSaveCredentials_Click(object sender, EventArgs e)
        {
            string fileName = "Credentials " + networkCredential.UserName + ".txt";
            FileInfo fileInfo = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + fileName);
            StreamWriter streamWriter;
            if (fileInfo.Exists)
            {
                fileInfo.Delete();
            }
            streamWriter = fileInfo.CreateText();
            streamWriter.WriteLine(networkCredential.UserName);
            streamWriter.WriteLine(networkCredential.Password);
            streamWriter.Close();
            MessageBox.Show("Credentials saved to " + fileInfo.FullName);
        }

        private void btnEditBody_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("notepad.exe", txtContent.Text);
        }

        private void mnuSelectNone_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in mainlist.Items)
            {
                item.Checked = false;
            }
        }

        private void btnRemoveAttachment_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listAttachments.Items)
            {
                if (item.Checked)
                {
                    item.Remove();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }

    class Mailer
    {
        private SmtpClient client;
        private Attachment attachment;

        public Mailer(NetworkCredential credentials)
        {
            client = new SmtpClient("smtp.googlemail.com");
            client.Port = 587;
            client.EnableSsl = true;
            client.Credentials = credentials;
            client.SendCompleted += new SendCompletedEventHandler(SendCompleted);
        }


        public void SendMessage(MailMessage message, ListViewX token)
        {
            client.SendAsync(message, token);
        }

        private static void SendCompleted(Object smtpClient, AsyncCompletedEventArgs @event)
        {
            ListViewX token;

            token = (ListViewX)@event.UserState;

            if (@event.Error != null)
            {
                token.item.SubItems[6].Text = @event.Error.Message;
            }
            else
            {
                token.item.SubItems[6].Text = "OK";
                token.item.Checked = false;
            }
            Form1.DeleteMessage(token.rowindex);
            ((SmtpClient)smtpClient).Dispose();
        }
    }


    class ListViewX
    {
        public ListViewItem item;
        public int rowindex;
    }
}
