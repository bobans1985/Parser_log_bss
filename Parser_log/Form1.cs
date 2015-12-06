using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


using System.Web;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;


namespace Parser_log
{
    public partial class Form1 : Form
    {
        public int i1 = 0, i2 = 0, i3 = 0, i4 = 0, i5 = 1, i6 = 1, i7 = 1; 


        public Form1()
        {
            InitializeComponent();
        }

        bool CheckTime()
        {
            bool res = false;
            TimeSpan dt = DateTime.Now.TimeOfDay;
            TimeSpan dt1 = new TimeSpan(8, 00, 0);
            TimeSpan dt2 = new TimeSpan(21, 00, 0);
            TimeSpan diff1 = dt1 - dt;
            TimeSpan diff2 = dt2 - dt;
            if ((diff1 < TimeSpan.Zero) && (diff2 > TimeSpan.Zero)) res=true;
 //           MessageBox.Show(diff1.ToString());
//            MessageBox.Show(diff2.ToString());
            return res;
        }
        
        string CreateTableHtml(string str1,string str2,string str3,string str4)
        {
        string res="<table height=119 border=1 cellpadding=0 cellspacing=0 bordercolor=#004664 bgcolor=#E6EBFF>"
           +"<tr valign=middle>"
           +"<td height=34 colspan=3><div align=center><strong>"+str1+"</strong></div></td>"
           +"</tr>"
           +"<tr valign=middle>"
           +"<td height=34><div align=center>Общее время выполнения операции </div></td>"
           +"<td height=34><div align=center>Начало</div></td>"
           +"<td><div align=center>Конец</div></td>"
           +"</tr>"
           +"<tr>"
           +"<td width=139 height=43 valign=middle><div align=center>"+str2+"</div></td>"
           + "<td width=110 valign=middle><div align=center>"+str3+"</div></td>"
           +"<td width=112 valign=middle><div align=center>"+str4+"</div></td>"
           +"</tr>"
           +"</table>";
        return res;
         }


        private void SendEmail()
        {

            //Авторизация на SMTP сервере
            SmtpClient Smtp = new SmtpClient("IP", 25);
            Smtp.Credentials = new NetworkCredential("username", "password");  //
            //Smtp.EnableSsl = false;

            //Формирование письма
            MailMessage Message = new MailMessage();
            Message.From = new MailAddress("from@from");
            Message.To.Add(new MailAddress("to@to"));
            Message.Subject = "Результаты парсинга лога";

            Message.Body = "Дата и время: " + DateTime.Now + "\n"
                   + CreateTableHtml(" Утреняя выписка ", label24.Text, label28.Text, label27.Text) + "<br>"
                   + CreateTableHtml(" Подготовка к утренней выписке ", label23.Text, label26.Text, label25.Text) + "<br>"
                   + CreateTableHtml(" Максимальное время выписки по автопроцедуре [ГО за 1 день] ", label2.Text, label3.Text, label4.Text) + "<br>"
                   + CreateTableHtml(" Максимально время выписки по автопроцедуре [ФИЛИАЛ за 1 день] ", label9.Text, label8.Text, label7.Text) + "<br>"
                   + CreateTableHtml(" Максимальное время выписки по автопроцедуре [ГО за 4 дня] ", label36.Text, label35.Text, label34.Text) + "<br>"
                   + CreateTableHtml(" Максимально время выписки по автопроцедуре [ФИЛИАЛ за 4 дня] ", label40.Text, label38.Text, label39.Text) + "<br>"
                   + CreateTableHtml(" Максимальное время квитовки документов ", label12.Text, label11.Text, label10.Text) + "<br>"
                   + CreateTableHtml(" Максимальное время выгрузки документов ", label18.Text, label17.Text, label16.Text) + "<br>";
                   /*+ richTextBox6.Text + "<br>-------------------------<br>"
                   + richTextBox7.Text + "<br>-------------------------<br>"
                   + richTextBox1.Text + "<br>-------------------------<br>"
                   + richTextBox4.Text + "<br>-------------------------<br>"
                   + richTextBox5.Text + "<br>-------------------------<br>";*/
            Message.IsBodyHtml = true;    

            //Прикрепляем файл
            // string file = "C:\\file.zip";
            //Attachment attach = new Attachment(file, MediaTypeNames.Application.Octet);

            // Добавляем информацию для файла
            //ContentDisposition disposition = attach.ContentDisposition;
            //disposition.CreationDate = System.IO.File.GetCreationTime(file);
            //disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
            //disposition.ReadDate = System.IO.File.GetLastAccessTime(file);

            //            Message.Attachments.Add(attach);

              Smtp.Send(Message);//отправка
        }

        private void LoadData()
        {
            richTextBox1.Text = "";
            richTextBox2.Text = "";
            richTextBox3.Text = "";
            richTextBox4.Text = "";
            richTextBox5.Text = "";
            richTextBox6.Text = "";
            richTextBox7.Text = "";
            richTextBox8.Text = "";
            label2.Text = "00:00:00";
            label3.Text = "00.00.0000";
            label4.Text = "00.00.0000";
            label9.Text = "00:00:00";
            label7.Text = "00.00.0000";
            label8.Text = "00.00.0000";
            label12.Text = "00:00:00";
            label11.Text = "00.00.0000";
            label10.Text = "00.00.0000";
            label18.Text = "00:00:00";
            label17.Text = "00.00.0000";
            label16.Text = "00.00.0000";
            label23.Text = "00:00:00";
            label25.Text = "00.00.0000";
            label26.Text = "00.00.0000";
            label24.Text = "00:00:00";
            label27.Text = "00.00.0000";
            label28.Text = "00.00.0000";
            label36.Text = "00:00:00";
            label35.Text = "00.00.0000";
            label34.Text = "00.00.0000";
            label40.Text = "00:00:00";
            label38.Text = "00.00.0000";
            label39.Text = "00.00.0000"; 


           // dataGridView1.Rows.Clear();

            int counter = 0;
            string line;
            string time1 = "";
            string time2 = "";
            string time3 = "";
            string time4 = "";
            string time5 = "";
            string time6 = "";
            string time7 = "";
            string time8 = "";
            string time9 = "";
            string time10 = "";
            string time11 = "";
            string time12 = "";
            string time13 = "";
            string time14 = "";
            string time15 = "";
            string time16 = "";

            DateTime dt1, dt2,dt3,dt4,dt5,dt6,dt7,dt8,dt9,dt10,dt11,dt12,dt13,dt14,dt15,dt16,dt1_real=DateTime.Now, dt2_real=DateTime.Now,
                dt3_real = DateTime.Now, dt4_real = DateTime.Now, dt5_real = DateTime.Now, dt6_real = DateTime.Now,
                dt7_real = DateTime.Now, dt8_real = DateTime.Now, dt9_real = DateTime.Now, dt10_real = DateTime.Now,
                dt11_real = DateTime.Now, dt12_real = DateTime.Now, dt13_real = DateTime.Now, dt14_real = DateTime.Now,
                dt15_real = DateTime.Now, dt16_real = DateTime.Now; ;
            TimeSpan diff, diff_real=TimeSpan.Zero;
            TimeSpan diff1 = TimeSpan.Zero, diff1_real = TimeSpan.Zero;
            TimeSpan diff2 = TimeSpan.Zero, diff2_real = TimeSpan.Zero;
            TimeSpan diff3 = TimeSpan.Zero, diff3_real = TimeSpan.Zero;
            TimeSpan diff4 = TimeSpan.Zero, diff4_real = TimeSpan.Zero;
            TimeSpan diff5 = TimeSpan.Zero, diff5_real = TimeSpan.Zero;
            TimeSpan diff6 = TimeSpan.Zero, diff6_real = TimeSpan.Zero;
            TimeSpan diff7 = TimeSpan.Zero, diff7_real = TimeSpan.Zero;
            TimeSpan diff8 = TimeSpan.Zero, diff8_real = TimeSpan.Zero;

            string filename = @"C:\BSSystems\SUBSYS\Logs\Sheduler\" + DateTime.Today.ToString("yyyyMMdd") + ".log";
            label15.Text = filename;
            Encoding WIN = Encoding.GetEncoding(1251);
            System.IO.StreamReader file =
                new System.IO.StreamReader(@filename, WIN);
            while ((line = file.ReadLine()) != null)
            {


/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.Обработка информации из АБС (квитовка);") != -1)
                {
                    richTextBox3.Text = richTextBox3.Text + line.Substring(1, 20) + line.Substring(114, line.Length - 150) + "\n";
                    time6 = line.Substring(11, 9);
                    dt5 = Convert.ToDateTime(time5);
                    dt6 = Convert.ToDateTime(time6);
                    diff2 = dt6 - dt5;
                    if (diff2_real < diff2)
                    {
                        diff2_real = diff2;
                        dt5_real = dt5;
                        dt6_real = dt6;
                    }
                    label12.Text = diff2_real.ToString();
                    label11.Text = dt5_real.ToString("g");
                    label10.Text = dt6_real.ToString("g");
                }
                if (line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.Обработка информации из АБС (квитовка)' произведен.") != -1)
                {
                    richTextBox3.Text = richTextBox3.Text + line.Substring(1, 20) + line.Substring(87, line.Length - 88) + "\n";
                    time5 = line.Substring(11, 9);
                }
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto._8 [ГО] Запрос выписок из АБС(GroupID=1,DateInterval=1,TheatCounts=5)") != -1)
                {
                    richTextBox5.Text = richTextBox5.Text + line.Substring(1, 20) + line.Substring(114, line.Length - 150) + "\n";
                    time2 = line.Substring(11, 9);
                    dt1 = Convert.ToDateTime(time1);
                    dt2 = Convert.ToDateTime(time2);
                    diff = dt2 - dt1;
                    if (diff_real < diff) 
                    {
                        diff_real = diff;
                        dt1_real = dt1;
                        dt2_real = dt2;

                    }
                    label2.Text = diff_real.ToString();
                    label3.Text = dt1_real.ToString("g");
                    label4.Text = dt2_real.ToString("g");

                }
                if (line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto._8 [ГО] Запрос выписок из АБС(GroupID=1,DateInterval=1,TheatCounts=5)' произведен.") != -1)
                {
                    richTextBox5.Text = richTextBox5.Text + line.Substring(1, 20) + line.Substring(87, line.Length - 88) + "\n";
                    time1= line.Substring(11, 9);
                }
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.9 [ФИЛИАЛ] Запрос выписок из АБС(GroupID=2,DateInterval=1,TheatCounts=5)") != -1)
                {
                    richTextBox8.Text = richTextBox8.Text + line.Substring(1, 20) + line.Substring(114, line.Length - 150) + "\n";
                    time4 = line.Substring(11, 9);
                    dt3 = Convert.ToDateTime(time3);
                    dt4 = Convert.ToDateTime(time4);
                    diff1 = dt4 - dt3;
                    if (diff1_real < diff1)
                    {
                        diff1_real = diff1;
                        dt3_real = dt3;
                        dt4_real = dt4;
                    }
                    label9.Text = diff1_real.ToString();
                    label8.Text = dt3_real.ToString("g");
                    label7.Text = dt4_real.ToString("g");
                }
                if (line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.9 [ФИЛИАЛ] Запрос выписок из АБС(GroupID=2,DateInterval=1,TheatCounts=5)' произведен.") != -1)
                {
                    richTextBox8.Text = richTextBox8.Text + line.Substring(1, 20) + line.Substring(87, line.Length - 88) + "\n";
                    time3 = line.Substring(11, 9);
                }
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (line.IndexOf("Завершение выполнения операции DocOut.Выгрузка платежных поручений;") != -1)
                {
                    richTextBox2.Text = richTextBox2.Text + line.Substring(1, 20) + line.Substring(110, line.Length - 145) + "\n";
                    time8 = line.Substring(11, 9);
                    label18.Text = time8;
                    if (time7 != "")
                    {
                        dt7 = Convert.ToDateTime(time7);
                        dt8 = Convert.ToDateTime(time8);
                        diff3 = dt8 - dt7;
                        if (diff3_real < diff3)
                        {
                            diff3_real = diff3;
                            dt7_real = dt7;
                            dt8_real = dt8;
                        }
                        label18.Text = diff3_real.ToString();
                        label17.Text = dt7_real.ToString("g");
                        label16.Text = dt8_real.ToString("g");
                    }

                }
                if (line.IndexOf("Вызов операции 'DocOut.Выгрузка платежных поручений' произведен.") != -1)
                {
                    richTextBox2.Text = richTextBox2.Text + line.Substring(1, 20) + line.Substring(84, line.Length - 84) + "\n";
                    time7 = line.Substring(11, 9);
                }
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto._1 [ГО] Запрос выписок из АБС  (GroupID=1) [ПОДГОТОВКА К УТР ВЫПИСКЕ];") != -1)
                {
                    richTextBox6.Text = richTextBox6.Text + line.Substring(1, 20) + line.Substring(114, line.Length - 145) + "\n";
                    time10 = line.Substring(11, 9);
                    if (time9 != "")
                    {
                        dt9 = Convert.ToDateTime(time9);
                        dt10 = Convert.ToDateTime(time10);
                        diff4 = dt10 - dt9;
                        if (diff4_real < diff4)
                        {
                            diff4_real = diff4;
                            dt9_real = dt9;
                            dt10_real = dt10;
                        }
                        label23.Text = diff4_real.ToString();
                        label26.Text = dt9_real.ToString("g");
                        label25.Text = dt10_real.ToString("g");
                    }

                }
                if (line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto._1 [ГО] Запрос выписок из АБС  (GroupID=1) [ПОДГОТОВКА К УТР ВЫПИСКЕ]' произведен") != -1)
                {
                    richTextBox6.Text = richTextBox6.Text + line.Substring(1, 20) + line.Substring(86, line.Length - 86) + "\n";
                    time9 = line.Substring(11, 9);
                }
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto._2 [ГО] Запрос выписок из АБС  (GroupID=1) [УТР ВЫПИСКА];") != -1)
                {
                    richTextBox7.Text = richTextBox7.Text + line.Substring(1, 20) + line.Substring(114, line.Length - 150) + "\n";
                    time12 = line.Substring(11, 9);
                    if (time11 != "")
                    {
                        dt11 = Convert.ToDateTime(time11);
                        dt12 = Convert.ToDateTime(time12);
                        diff5 = dt12 - dt11;
                        if (diff5_real < diff5)
                        {
                            diff5_real = diff5;
                            dt11_real = dt11;
                            dt12_real = dt12;
                        }
                        label24.Text = diff5_real.ToString();
                        label28.Text = dt11_real.ToString("g");
                        label27.Text = dt12_real.ToString("g");
                    }

                }
                if (line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto._2 [ГО] Запрос выписок из АБС  (GroupID=1) [УТР ВЫПИСКА]' произведен.") != -1)
                {
                    richTextBox7.Text = richTextBox7.Text + line.Substring(1, 20) + line.Substring(86, line.Length - 86) + "\n";
                    time11 = line.Substring(11, 9);
                }
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto._3 [ГО] Запрос выписок из АБС  (GroupID=1)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto._4 [ГО] Запрос выписок из АБС  (GroupID=1)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto._5 [ГО] Запрос выписок из АБС  (GroupID=1)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto._6 [ГО] Запрос выписок из АБС  (GroupID=1)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto._7 [ГО] Запрос выписок из АБС  (GroupID=1)") != -1)
                {
                    richTextBox1.Text = richTextBox1.Text + line.Substring(1, 20) + line.Substring(114, line.Length - 150) + "\n";
                    time14 = line.Substring(11, 9);
                    dt13 = Convert.ToDateTime(time13);
                    dt14 = Convert.ToDateTime(time14);
                    diff7 = dt14 - dt13;
                    if (diff7_real < diff7)
                    {
                        diff7_real = diff7;
                        dt13_real = dt13;
                        dt14_real = dt14;
                    }
                    label36.Text = diff7_real.ToString();
                    label35.Text = dt13_real.ToString("g");
                    label34.Text = dt14_real.ToString("g");
                }
                if (line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto._3 [ГО] Запрос выписок из АБС  (GroupID=1)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto._4 [ГО] Запрос выписок из АБС  (GroupID=1)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto._5 [ГО] Запрос выписок из АБС  (GroupID=1)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto._6 [ГО] Запрос выписок из АБС  (GroupID=1)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto._7 [ГО] Запрос выписок из АБС  (GroupID=1)' произведен.") != -1)                                    
                {
                    richTextBox1.Text = richTextBox1.Text + line.Substring(1, 20) + line.Substring(87, line.Length - 88) + "\n";
                    time13 = line.Substring(11, 9);
                }
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.1 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.2 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.3 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.4 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.5 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.6 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.7 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)") != -1 ||
                    line.IndexOf("Завершение выполнения операции LinkABS.LinkABSAuto.8 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)") != -1)
                {
                    richTextBox4.Text = richTextBox4.Text + line.Substring(1, 20) + line.Substring(114, line.Length - 150) + "\n";
                    time16 = line.Substring(11, 9);
                    dt15 = Convert.ToDateTime(time15);
                    dt16 = Convert.ToDateTime(time16);
                    diff8 = dt16 - dt15;
                    if (diff8_real < diff8)
                    {
                        diff8_real = diff8;
                        dt15_real = dt15;
                        dt16_real = dt16;
                    }
                    label40.Text = diff8_real.ToString();
                    label38.Text = dt15_real.ToString("g");
                    label39.Text = dt16_real.ToString("g");
                }
                if (line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.1 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.2 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.3 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.4 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.5 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.6 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.7 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)' произведен.") != -1 ||
                    line.IndexOf("Вызов операции 'LinkABS.LinkABSAuto.8 [ФИЛИАЛ] Запрос выписок из АБС  (GroupID=2)' произведен.") != -1)                
                {
                    richTextBox4.Text = richTextBox4.Text + line.Substring(1, 20) + line.Substring(87, line.Length - 88) + "\n";
                    time15 = line.Substring(11, 9);
                }
                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   

         }
            file.Close();

        }


        private void Form1_Load(object sender, EventArgs e)
        {
//              LoadData();
        }

    

        private void button1_Click(object sender, EventArgs e)
        {
            if (CheckTime()) 
            LoadData();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (CheckTime()) 
            SendEmail();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (CheckTime()) 
            LoadData();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (button3.Text == "Обновлять данные автоматом")
            {
                timer1.Interval = Convert.ToInt16(numericUpDown1.Value) * 1000 * 60;
                timer1.Enabled = true;
                button3.Text = "Остановить обновление";
            }
            else 
            {
                timer1.Enabled = false;
                button3.Text = "Обновлять данные автоматом";
            }

            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (button4.Text == "Отправлять автоматом")
            {
                timer2.Interval = Convert.ToInt16(numericUpDown2.Value) * 1000 * 60;
                timer2.Enabled = true;
                button4.Text = "Остановить отправку";
            }
            else
            {
                timer2.Enabled = false;
                button4.Text = "Отправлять автоматом";
            }

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (CheckTime()) 
            SendEmail();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form Settings = new Form2();
            Settings.ShowDialog();
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            bool res = false;
            TimeSpan dt = DateTime.Now.TimeOfDay;
            TimeSpan dt1 = new TimeSpan(16, 00, 0);
            TimeSpan dt2 = new TimeSpan(16, 01, 0);
            TimeSpan diff1 = dt1 - dt;
            TimeSpan diff2 = dt2 - dt;
            if ((diff1 < TimeSpan.Zero) && (diff2 > TimeSpan.Zero)) res = true;
            if (res)
            {
                LoadData();
                SendEmail();
            }

            res = false;
            dt = DateTime.Now.TimeOfDay;
            dt1 = new TimeSpan(22, 00, 0);
            dt2 = new TimeSpan(22, 01, 0);
            diff1 = dt1 - dt;
            diff2 = dt2 - dt;
            if ((diff1 < TimeSpan.Zero) && (diff2 > TimeSpan.Zero)) res = true;
            if (res)
            {
                LoadData();
                SendEmail();
            }

        }

          }
}