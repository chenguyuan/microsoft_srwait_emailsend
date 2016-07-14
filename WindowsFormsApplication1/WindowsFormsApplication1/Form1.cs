using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Text;
using System.Web;
using System.Reflection;
using Microsoft.Exchange.WebServices.Data;

//考虑增加的功能：选择需要发送的工程师（checkedlistBox）

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public string filePath = "C:\\Users\\t-guch\\Downloads\\Case Wellness.xlsx";//excel的存储地址,需要变动
        public List<string> emplyee = new List<string>(); //员工姓名列表
        public List<Emplyee> EmplyeeList = new List<Emplyee>();
        public DataTable dt = new DataTable();//源数据
        //状态字符串，共10种，0-9
        public string[] SRWaitstate = new string[] { "Pending CTS", "Pending Customer", "Pending Development", "Pending Operations", "Mitigated-Pending RFC", "Solution Delivered - Pending Confirmation", "Solution Delivered - Solution Confirmed", "Pending Premier", "Pending 3rd party", "Recovery" };

        ExchangeService Exservice = new ExchangeService();//exchange连接

        public partial class Emplyee
        {
            public string alias;
            public string emailtable;
            public string emailBody;
            public Emplyee(string emplyeealias)
            {
                this.alias = emplyeealias;
            }

        }

        public Form1() //初始化
        {

            InitializeComponent();
            if (File.Exists(filePath))//判断默认位置是否有excel
            {
                EmplyeelistInitialize();
            }
            this.comboBox2.SelectedIndex = 0;
        }

        private void getTable() //获得源数据
        {
            try
            {
                FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet result = excelReader.AsDataSet();
                dt = result.Tables[0];
                stream.Close();
                //调试监控
                Console.WriteLine("datasource got");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void button1_Click(object sender, EventArgs e) //发送按钮
        {
            if (emplyee.Count > 2)
            {
                int num = 0;
                progressBar1.Visible = true;
                progressBar1.Value = 3;
                progressBar1.Maximum = emplyee.Count + 8;
                progressBar1.BringToFront();
                progressBar1.Show();
                try
                {
                    string toolUser = string.Empty;
                    toolUser = comboBox2.SelectedItem.ToString();
                    exserviceSet(toolUser);
                    progressBar1.Value += 3;
                    //实例化每个emplyee
                    foreach (string s in emplyee)
                    {
                        Emplyee temp = new Emplyee(s);
                        temp.emailtable = buildHtmlTable(temp.alias);
                        temp.emailBody = ReplaceText(temp.alias, toolUser, temp.emailtable);
                        EmplyeeList.Add(temp);
                    }
                    //正式发邮件
                    string emailTo = string.Empty;
                    foreach (Emplyee em in EmplyeeList)
                    {
                        emailTo = "t-guch" + "@microsoft.com";
                        if (sendEmailbyExchange(emailTo, em.emailBody)) { num++; };
                        progressBar1.Value++;

                    }
                    //MessageBox.Show(x.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                progressBar1.Hide();
                progressBar1.Visible = false;
                MessageBox.Show("send" + num.ToString() + "emails");
            }
            else { MessageBox.Show("Please get the Excel first!"); }
        }

        private void EmplyeelistInitialize() // 提取所有员工、并放入了下拉菜单（隐藏）
        {
            try
            {
                emplyee.Clear();
                getTable();
                foreach (DataRow dr in dt.Rows)
                {
                    string alias = dr["Column18"].ToString();
                    if (alias != "" && alias != "Owner Employee Email")
                    {
                        emplyee.Add(alias);
                    }
                }
                emplyee = emplyee.Distinct().ToList();//去重
                foreach (string emp in emplyee)
                {
                    comboBox1.Items.Add(emp);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            button1.Enabled = true;
        }
        public string buildHtmlTable(string emplyeename)//根据员工姓名建立html表格
        {
            string htmltable = string.Empty;
            
            StringBuilder sb = new StringBuilder();
            sb.Append("<table width=\"95%\" border=\"1\" cellpadding=\"2\" cellspacing=\"1\">");

            //表内容
            if (dt.Rows.Count > 1)
            {
                DataTable temptable = new DataTable();//临时存放某员工的所有case
                temptable = dt.Clone();
                foreach (DataRow dr in dt.Rows)
                {
                    string alias = dr["Column18"].ToString();
                    if (alias == emplyeename)
                    {
                        temptable.ImportRow(dr);
                    }
                }

                foreach (string statename in SRWaitstate)
                {
                    bool flag = false;//标记是不是写过标题栏了
                    foreach (DataRow dr in temptable.Rows)
                    {
                        string state = dr["Column6"].ToString();
                        if (state == statename)
                        {
                            if (flag == false)//写标题行
                            {
                                sb.Append("  <tr>");
                                sb.Append("    <th colspan=\"5\" bgcolor=\"#FFCC00\" scope=\"col\">SRWait State:" + statename + "</th>");
                                sb.Append("  </tr>");
                                sb.Append("  <tr>");
                                sb.Append("    <th width=\"130\" bgcolor=\"#00FFCC\" scope=\"col\">Service Request Number</th>");
                                sb.Append("    <th nowrap bgcolor=\"#00FFCC\" scope=\"col\">SRTitle Internal</th>");
                                sb.Append("    <th width=\"70\" bgcolor=\"#00FFCC\" scope=\"col\">Days Open</th>");
                                sb.Append("    <th width=\"70\" bgcolor=\"#00FFCC\" scope=\"col\">Total Labor Minutes</th>");
                                sb.Append("    <th width=\"100\" bgcolor=\"#00FFCC\" scope=\"col\"><p>Labor Idle");
                                sb.Append("    <strong>(days from last labor date)</strong></p></th>  ");
                                sb.Append("  </tr>");
                                sb.Append("<tr>");
                                sb.Append("    <td><a href=\"mssv://sr/?" + dr["Column1"] + "\">" + dr["Column1"] + "</a></td>");
                                sb.Append("    <td>" + dr["Column2"] + "</td>");
                                sb.Append("    <td>" + dr["Column3"] + "</td><td>" + dr["Column4"] + "</td><td>" + dr["Column5"] + "</td>");
                                sb.Append("  </tr>");
                                flag = true;
                            }
                            else
                            {
                                sb.Append("<tr>");
                                sb.Append("    <td><a href=\"mssv://sr/?"+dr["Column1"]+"\">" + dr["Column1"] + "</a></td>");
                                sb.Append("    <td>" + dr["Column2"] + "</td>");
                                sb.Append("    <td>" + dr["Column3"] + "</td><td>" + dr["Column4"] + "</td><td>" + dr["Column5"] + "</td>");
                                sb.Append("  </tr>");
                            }
                        }
                    }
                }
            }
            sb.Append("</table>");

            htmltable = htmltable+sb;

            return htmltable;
        }

        private bool exserviceSet(string sender)//设置Exchange Service
        {
            //Exservice.Credentials = new WebCredentials("t-guch@microsoft.com", "Mike704@ms");
            Exservice.UseDefaultCredentials = true;
            Exservice.TraceEnabled = true;
            Exservice.TraceFlags = TraceFlags.All;
            string myEmailaddress = sender + "@microsoft.com";
            Exservice.AutodiscoverUrl(myEmailaddress, RedirectionUrlValidationCallback);//发件人

            //if AutodiscoverUrl worked
            var y= Exservice.Url;
            if (y == null) return false;
            return true; }

        //理论上发送成功返回true
        private bool sendEmailbyExchange(string emailto,string emailbody)
        {
            try
            {
                EmailMessage email = new EmailMessage(Exservice);
                email.ToRecipients.Add(emailto);//收件人
                email.Subject = "TEST Email for SRwait alart";
                email.Body = new MessageBody(emailbody);
                email.Body.BodyType = Microsoft.Exchange.WebServices.Data.BodyType.HTML;
                email.Send();
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); return false;  }

            return true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        #region 管理文件路径
        ////没写好这里
        //private void filepathlog()
        //{
        //    string file = Application.ExecutablePath + "cache.txt";
        //    string content = "mike";
        //    if (!File.Exists(file) == true)
        //    {
        //        MessageBox.Show("存在此文件!");
        //    }
        //    else
        //    {
        //        FileStream myFs = new FileStream(file, FileMode.Create);
        //        StreamWriter mySw = new StreamWriter(myFs);
        //        mySw.Write(content);
        //        mySw.Close();
        //        myFs.Close();
        //        MessageBox.Show("写入成功");
        //    }            
        //}



        private void btnPath_Click(object sender, EventArgs e)
        {
            
        }
        #endregion

        /// <summary>     
        ///替换HTML模板中的字段值     
        /// </summary>     
        public string ReplaceText(String userName, string myName,string table)
        {

            string html = string.Empty;
            html=Resource1.emailTemplate;

                if (html == string.Empty)
            {
                return string.Empty;
            }

            html = html.Replace("$USER_NAME$", userName);
            html = html.Replace("$TABLEREPLACE$", table);
            html = html.Replace("$MY_NAME$", myName);
            return html;
        }

        /// <summary>     
        /// This validates whether redirected URLs returned by Autodiscover represent an HTTPS endpoint.     
        /// </summary>    
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }


        #region 参考邮件发送方法

        private void sendemail()
        {
            try
            {
                //确定smtp服务器地址。实例化一个Smtp客户端
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("smtp.163.com");
                //生成一个发送地址
                string strFrom = string.Empty;
                strFrom = "qingchuntongji@163.com";

                //构造一个发件人地址对象
                MailAddress from = new MailAddress(strFrom, "guyuan", Encoding.UTF8);
                //构造一个收件人地址对象
                MailAddress to = new MailAddress("330943592@qq.com", "guyuanchen", Encoding.UTF8);

                //构造一个Email的Message对象
                MailMessage message = new MailMessage(from, to);

                ////为 message 添加附件
                //foreach (TreeNode treeNode in treeViewFileList.Nodes)
                //{
                //    //得到文件名
                //    string fileName = treeNode.Text;
                //    //判断文件是否存在
                //    if (File.Exists(fileName))
                //    {
                //        //构造一个附件对象
                //        Attachment attach = new Attachment(fileName);
                //        //得到文件的信息
                //        ContentDisposition disposition = attach.ContentDisposition;
                //        disposition.CreationDate = System.IO.File.GetCreationTime(fileName);
                //        disposition.ModificationDate = System.IO.File.GetLastWriteTime(fileName);
                //        disposition.ReadDate = System.IO.File.GetLastAccessTime(fileName);
                //        //向邮件添加附件
                //        message.Attachments.Add(attach);
                //    }
                //    else
                //    {
                //        MessageBox.Show("文件" + fileName + "未找到！");
                //    }
                //}

                //添加邮件主题和内容
                message.Subject = "test";
                message.SubjectEncoding = Encoding.UTF8;
                message.Body = "test";
                message.BodyEncoding = Encoding.UTF8;

                //设置邮件的信息
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.IsBodyHtml = false;

                //如果服务器支持安全连接，则将安全连接设为true。
                //gmail支持，163不支持，如果是gmail则一定要将其设为true
                //if (cmbBoxSMTP.SelectedText == "smpt.163.com")
                client.EnableSsl = false;
                // else
                //client.EnableSsl = true;

                //设置用户名和密码。
                //string userState = message.Subject;
                client.UseDefaultCredentials = false;
                string username = "qingchuntongji@163.com";
                string passwd = "qingchunTJ";
                //用户登陆信息
                NetworkCredential myCredentials = new NetworkCredential(username, passwd);
                client.Credentials = myCredentials;
                //发送邮件
                client.Send(message);
                //提示发送成功
                MessageBox.Show("发送成功!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        } //普通email
        private void sendEmailbyHtml(string email_from, string email_to, string email_cc, string strbody)
        {
            try
            {
                // 建立一个邮件实体     
                MailAddress from = new MailAddress(email_from);


                MailAddress to = new MailAddress(email_to);
                MailMessage message = new MailMessage(from, to);


                if (email_cc.ToString() != string.Empty)
                {
                    foreach (string ccs in email_cc.Split(';'))
                    {
                        MailAddress cc = new MailAddress(ccs);
                        message.CC.Add(cc);
                    }
                }

                message.IsBodyHtml = true;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.Priority = MailPriority.High;
                message.Body = strbody;  //邮件BODY内容    
                message.Subject = "Subject";
                //微软内部的smtp怎么使用
                //System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("smtp.microsoft.com");
                //client.Credentials = new System.Net.NetworkCredential("t-guch@microsoft.com", "Mike704@ms");

                //外部邮件
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("smtp.163.com");
                client.Credentials = new System.Net.NetworkCredential("qingchuntongji@163.com", "qingchunTJ");

                client.Send(message); //发送邮件    

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


        #endregion
    }

    
}
