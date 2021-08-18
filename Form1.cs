using IronOcr;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CheckTxt();
        }

        private void CheckTxt()
        {
            tuopan();
            Thread.Sleep(1000);
            var img = GetScreen();
            pictureBox1.Image = img;
            using (var input = new OcrInput())
            {
                input.AddImage(img);
                IronTesseract Ocr = new IronTesseract();
                Ocr.Language = OcrLanguage.ChineseSimplified;
                Ocr.AddSecondaryLanguage(OcrLanguage.English);
                var Result = Ocr.Read(input);
                this.textBox1.Text = DateTime.Now.ToString("yyyy年MM月dd日hh时mm分ss秒") + "\r\n" + Result.Text;
                if (Result.Text.Contains("生/成"))
                {
                    Console.WriteLine("TEST");
                    StringBuilder strInfo = new StringBuilder();
                    strInfo.Append(string.Format("Hello-----------------------------"));//邮件主体内容（自己拼接的）
                    string fromEmail = "m76812791@outlook.com";//发件邮箱 自己邮箱的服务
                    string emailPwd = "qq18741689356";//发件邮箱密码
                    string toEmail = "1273955317@qq.com";//收件邮箱
                    string emailTitle = "你好----------------------------";//邮件标题
                    string emailContent = strInfo.ToString() + this.textBox1.Text;//邮件主体内容
                    string SmtpHost = "smtp.office365.com";
                    int SmtpPort = 587;
                    if (SmtpMailSend(SmtpHost, SmtpPort, fromEmail, emailPwd, emailTitle, emailContent, true, toEmail))
                    {
                    }
                    else
                    {
                    }
                }
            }
        }
        public Bitmap GetScreen()
        {
            Rectangle ScreenArea = Screen.GetWorkingArea(this);
            Bitmap bmp = new Bitmap(ScreenArea.Width, ScreenArea.Height);

            var leftX = (ScreenArea.Width - 300) / 2;
            var leftY = (ScreenArea.Height - 300) / 2;
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(200, 150, 0, 0, new Size(500, 500));
            }
            return bmp;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            CheckTxt();
        }


        public bool SmtpMailSend(string smtpHost, int smtpPort, string fromAddress, string mailPassword, string title, string body, bool isBodyHtml, params string[] recipient)
        {
            try
            {

                MailMessage myMail = new MailMessage();
                myMail.From = new MailAddress(fromAddress);
                foreach (string item in recipient)
                {
                    if (!string.IsNullOrEmpty(item))
                        myMail.To.Add(new MailAddress(item));
                }
                myMail.Subject = title; //邮件标题
                myMail.SubjectEncoding = Encoding.UTF8;//标题编码
                myMail.Body = body;//邮件主体
                myMail.BodyEncoding = Encoding.UTF8;
                myMail.IsBodyHtml = isBodyHtml;//是否HTML
                SmtpClient smtp = new SmtpClient();
                smtp.Host = smtpHost;
                smtp.Port = smtpPort; //Gmail的smtp端口
                                      // mail.smtp.starttls.enable = true
                smtp.EnableSsl = true;//Gmail要求SSL连接
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = new System.Net.NetworkCredential(fromAddress, mailPassword);
                //smtp.EnableSsl = true; //Gmail要求SSL连接
                //smtp.DeliveryMethod = SmtpDeliveryMethod.Network; //Gmail的发送方式是通过网络的方式，需要指定
                try
                {
                    smtp.Send(myMail);
                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            //tuopan();
        }

        private void tuopan()
        {
            // if (this.WindowState == FormWindowState.Minimized)
            this.Hide();
            this.notifyIcon1.Visible = true;
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
            //this.notifyIcon1.Visible = false;
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F1:
                  
                    break;

                case Keys.F2:
                  
                    break;
            }

            // 组合键
            if (e.Modifiers == Keys.Alt && e.KeyCode == Keys.D0)
            {
                CheckTxt();
            }
            CheckTxt();

            //if (e.KeyCode == Keys.S && e.Modifiers == Keys.Control)         //Ctrl+F1
            //{
            //    CheckTxt();
            //}
        }
    }
}
