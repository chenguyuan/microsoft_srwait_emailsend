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
using mshtml;
using System.Threading;
using System.ComponentModel;

//考虑增加的功能：选择需要发送的工程师（checkedlistBox）

namespace SR_Wait_State_Summary
{
    public partial class Form1 : Form
    {
        
        string[] args= {"","",""};//外部参数
        bool test = false;//默认不是测试
        bool auto = false;//默认不是自动
        public string filePath = "";//excel的存储地址,需要变动
        public List<string> emplyee = new List<string>(); //员工姓名列表
        public List<Emplyee> EmplyeeList = new List<Emplyee>();
        public DataTable dt = new DataTable();//源数据
        //状态字符串，共10种，0-9
        public string[] SRWaitstate = new string[] { "Pending 3rd party", "Pending CTS", "Pending Customer", "Pending Development", "Pending Operations", "Pending Premier", "Recovery", "Mitigated-Pending RFC", "Solution Delivered - Pending Confirmation", "Solution Delivered - Solution Confirmed" };
        public string[] SRWaitstateexplain = new string[] { "Select when the responsible party of current key action in this SR is a non-CSS team (for collaborations) or an outside 3rd party.", "The default case status on SR creation. Select when the responsible party of current key action in this SR is case owner. Also use when no other allowed wait state category is appropriate.","Select when the responsible party of current key action in this SR is customer/partner.", "Select when the responsible party of current key action in this SR is engineering group (for example Bugs / RFCs / Hotfixes / CFLs)", "DO NOT USE – see Pending Development","DO NOT USE", "DO NOT USE", "DO NOT USE – see Pending Development", "Select when the solution to the problem is offered to the customer/partner and we are waiting for customer/partner confirmation.", "Select when the customer/partner has successfully confirmed the offered solution is accepted by the customer/partner." };
        ExchangeService Exservice = new ExchangeService();//exchange连接
        public string toolUser = "t-guch";//可能会删掉的参数
        public HtmlElement elem = null;
        public string elemstyle = string.Empty;
        public static string url = Environment.CurrentDirectory.ToString();
        public StreamWriter log = new StreamWriter(url + "\\log.txt", true);
        public string address = "http://gbs-sandbox/ReportServer/Pages/ReportViewer.aspx?/CTS%20Reports/GBSDBI/SR%20Wellness/Case%20Wellness&LaborMins=0&UserRole=Team%20Manager";
        public string line;
        public string add_alias; public string add_group;
        delegate void set_Elemstyle();
        set_Elemstyle Set_Elemstyle;

        //private Thread thread1;



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

        public Form1(string[] args) //初始化
        {
            this.args = args;
            foreach (string s in args)
            { if (s == "-test") test = true;
                if (s == "-auto") auto = true;
            }
            InitializeComponent();
            //从指定文件中读入参数
            System.IO.StreamReader file = new System.IO.StreamReader(url + "\\parameters.txt");
            int i;
            i = 0;
            while ((line = file.ReadLine()) != null && i<=10)
            {
                if (i == 0)
                {
                    add_alias = line;
                    i++;
                }
                else if (i == 1)
                {
                    add_group = line;
                    i++;
                }
            }
            file.Close();
            address = address + "&" + add_alias + "&" + add_group;
            webBrowser1.Navigate(address);

            //auto的话读location，不然的话读lastL
            if (auto)
            {
                if (File.Exists(url + "\\location.txt"))
                {
                    filePath = Read(url + "\\location.txt");
                    textBox1.Text = filePath;
                }
                else
                {
                    filePath = textBox1.Text.ToString();
                }
            }
            else
            {
                if (File.Exists(url + "\\lastL.txt"))
                {
                    filePath = Read(url + "\\lastL.txt");
                    textBox1.Text = filePath;
                }
                else
                {
                    filePath = textBox1.Text.ToString();
                }
            }
            
            if (File.Exists(filePath))//判断默认位置是否有excel
            {
                EmplyeelistInitialize();
            }
            this.comboBox2.SelectedIndex = 0;
            Set_Elemstyle = new set_Elemstyle(set_elemstylemike);

            //命令行执行自动发送
            if (auto)
            {
                FileInfo fi = new FileInfo(filePath);
                TimeSpan t1 = System.DateTime.Now - fi.LastWriteTime;
                if (Math.Abs(t1.Days) > 1)
                {
                    log.WriteLine("Please check the Subscription of the case wellness, the Excel hasn't been update in two days. log time: " + System.DateTime.Now.ToString());
                }
                else
                {
                    auto_Click();
                }
                System.Environment.Exit(0);
            }
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
        private void auto_Click()
        {
            if (emplyee.Count >=1)
            {
                int num = 0;//计算发送邮件数量
                progressBar1.Show();
                progressBar1.Visible = true;
                progressBar1.Maximum = emplyee.Count + 10;
                progressBar1.Value = 2;
                progressBar1.BringToFront();
                label6.Visible = true;

                try
                {
                    progressBar1.Value += 2;
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

                    //测试用
                    if (checkBox1.Checked || test)
                    {
                        //emailTo = toolUser + "@microsoft.com";
                        //if (sendEmailbyExchange(emailTo, EmplyeeList[4].emailBody)) { num++; };
                        //progressBar1.Value++;
                        foreach (Emplyee em in EmplyeeList)
                        {
                            emailTo = toolUser + "@microsoft.com";
                            if (sendEmailbyExchange(emailTo, em.emailBody)) { num++; };
                            progressBar1.Value++;

                        }
                    }
                    //正式代码
                    else if (!(checkBox1.Checked))
                    {
                        foreach (Emplyee em in EmplyeeList)
                        {
                            //emailTo = toolUser + "@microsoft.com";
                            emailTo = em.alias + "@microsoft.com";//正式发布时使用
                            if (sendEmailbyExchange(emailTo, em.emailBody)) { num++; };
                            progressBar1.Value++;
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                progressBar1.Hide();
                progressBar1.Visible = false;
                label6.Visible = false;
                log.WriteLine("send " + num.ToString() + " emails at "+ System.DateTime.Now.ToString());
                log.Close();
            }
            else { MessageBox.Show("Please make sure excel and data exist!"); 
            }
        }

        private void button1_Click(object sender, EventArgs e) //发送按钮
        {
            if (emplyee.Count > 2)
            {
                int num = 0;//计算发送邮件数量
                progressBar1.Show();
                progressBar1.Visible = true;
                progressBar1.Maximum = emplyee.Count + 10;
                progressBar1.Value = 2;
                progressBar1.BringToFront();
                label6.Visible = true;
                
                try
                {
                    progressBar1.Value +=2;
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

                    //测试用
                    if (checkBox1.Checked||test)
                    {
                        //emailTo = toolUser + "@microsoft.com";
                        //if (sendEmailbyExchange(emailTo, EmplyeeList[4].emailBody)) { num++; };
                        //progressBar1.Value++;
                        foreach (Emplyee em in EmplyeeList)
                        {
                            emailTo = toolUser + "@microsoft.com";
                            if (sendEmailbyExchange(emailTo, em.emailBody)) { num++; };
                            progressBar1.Value++;

                        }
                    }
                    //正式代码
                    else if (!(checkBox1.Checked))
                    {
                        foreach (Emplyee em in EmplyeeList)
                        {
                            //emailTo = toolUser + "@microsoft.com";
                            emailTo = em.alias + "@microsoft.com";//正式发布时使用
                            if (sendEmailbyExchange(emailTo, em.emailBody)) { num++; };
                            progressBar1.Value++;



                        }
                    }
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                progressBar1.Hide();
                progressBar1.Visible = false;
                label6.Visible = false;
                MessageBox.Show("send " + num.ToString() + " emails");
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
            sb.Append("<table width=\"98%\" border=\"1\" cellpadding=\"0\" cellspacing=\"0\" align=\"center\" style=\"border - collapse:collapse; \">");

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
                //按照最近交流时间进行排序
                DataTable temptableSort = temptable.Copy();
                DataView dv = temptable.DefaultView;
                dv.Sort = "Column7";
                temptableSort = dv.ToTable();
                bool flag2 = true;//标记是不是第一行了
                int tempcount = 0;//确定到第几个state了
                foreach (string statename in SRWaitstate)
                {
                    string stateexplain=SRWaitstateexplain[tempcount];
                    bool flag = false;//标记是不是写过标题栏了

                    string longidle = string.Empty;
                    foreach (DataRow dr in temptableSort.Rows)
                    {
                        string state = dr["Column6"].ToString();
                        if (state == statename)
                        {

                            if (flag == false)//写标题行 
                            {
                                if (flag2 == false)
                                { 
                                sb.Append("  <tr>");
                                sb.Append("    <th colspan=\"6\" height=\"2\" bgcolor=\"#cccccc\" scope=\"col\"></th>");
                                sb.Append("  </tr>");
                                    
                                }
                                sb.Append("  <tr>");
                                sb.Append("    <th colspan=\"6\" height=\"35\" bgcolor=\"#4682B4\" class=\"titlemike\" scope=\"col\">" + statename + "</th>");
                                sb.Append("  </tr>");
                                sb.Append("  <tr>");
                                sb.Append("    <th width=\"130\" bgcolor=\"#42B0B9\" class=\"titlemike3\" scope =\"col\">Case ID</th>");
                                sb.Append("    <th nowrap bgcolor=\"#42B0B9\" class=\"titlemike\" scope=\"col\">SR Title</th>");
                                sb.Append("    <th width=\"60\" bgcolor=\"#42B0B9\" class=\"titlemike\" scope=\"col\">Last Communication</th>");
                                sb.Append("    <th width=\"70\" bgcolor=\"#42B0B9\" class=\"titlemike\" scope=\"col\">Days Open</th>");
                                sb.Append("    <th width=\"60\" bgcolor=\"#42B0B9\" class=\"titlemike\" scope=\"col\">Labor</th>");
                                sb.Append("    <th width=\"70\" bgcolor=\"#42B0B9\" class=\"titlemike\" scope=\"col\">Idle Days</th>");
                                sb.Append("  </tr>");

                                if (Convert.ToInt32(dr["Column5"]) >= 5)
                                { longidle = "class=\"notemike\""; }
                                else { longidle = ""; }
                                sb.Append("<tr "+longidle+ " >");
                                sb.Append("    <td align=\"center\" class=\"mike\"><a href=\"mssv://sr/?" + dr["Column1"] + "\">" + dr["Column1"] + "</a></td>");
                                sb.Append("    <td align=\"center\" class=\"mike\">" + dr["Column2"] + "</td><td align=\"center\" class=\"mike\">" + dr["Column7"] + "</td>");
                                sb.Append("    <td align=\"center\" class=\"mike\">" + dr["Column3"] + "</td><td align=\"center\" class=\"mike\">" + dr["Column4"] + "</td><td align=\"center\" class=\"mike\">" + dr["Column5"] + "</td>");
                                sb.Append("  </tr>");
                                flag = true;
                                flag2 = false;
                            }
                            else
                            {
                                if (Convert.ToInt32(dr["Column5"]) >= 5)
                                { longidle = "class=\"notemike\""; }
                                else { longidle = ""; }
                                sb.Append("<tr " + longidle + " >");
                                sb.Append("    <td align=\"center\" class=\"mike\"><a href=\"mssv://sr/?" + dr["Column1"] + "\">" + dr["Column1"] + "</a></td>");
                                sb.Append("    <td align=\"center\" class=\"mike\">" + dr["Column2"] + "</td><td align=\"center\" class=\"mike\">" + dr["Column7"] + "</td>");
                                sb.Append("    <td align=\"center\" class=\"mike\">" + dr["Column3"] + "</td><td align=\"center\" class=\"mike\">" + dr["Column4"] + "</td><td align=\"center\" class=\"mike\">" + dr["Column5"] + "</td>");
                                sb.Append("  </tr>");
                            }
                        }
                    }
                    tempcount++;
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
                email.Subject = "SR Wait State Summary was executed at "+System.DateTime.Now.ToString();
                email.Body = new MessageBody(emailbody);
                email.Body.BodyType = Microsoft.Exchange.WebServices.Data.BodyType.HTML;
                email.Send();
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); return false;  }

            return true;
        }

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

        #region 管理文件路径
        public string Read(string path)//读文件
        {
            FileStream sr = File.Open(path, FileMode.Open);
            char[] chs = new char[sr.Length];
            for (int i = 0; i < sr.Length; i++)
            {
                chs[i] = (char)sr.ReadByte();
            }
            string ss = new string(chs);
            sr.Close();
            return ss;

        }

        private void btnPath_Click(object sender, EventArgs e)//change file path
        {
            try
            {
                HtmlElement elem = this.webBrowser1.Document.GetElementById("ReportViewerControl_AsyncWait_Wait");
                if (elem != null)
                {
                    String nameStr = elem.Style;
                    Console.WriteLine(nameStr);
                }
                OpenFileDialog open = new OpenFileDialog();
                if (open.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = open.FileName;
                }
                filePath = textBox1.Text.ToString();
                EmplyeelistInitialize();
                EditFile(filePath, url + "\\lastL.txt"); 
            }
            catch (Exception Ex) { throw Ex; }
        }
        public static void EditFile(string newLineValue, string patch)//修改patch
        {
            FileStream fs = new FileStream(patch, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(newLineValue);
            sw.Close();
            fs.Close();
        }
        #endregion

        private void getExcelbutton_Click(object sender, EventArgs e)//获取excel
        {
            HtmlElementCollection elemList = this.webBrowser1.Document.GetElementsByTagName("a");
            foreach (HtmlElement elem in elemList)
            {
                String nameStr = elem.GetAttribute("title");
                if (nameStr == "Excel")
                {
                    elem.InvokeMember("click");
                }
            }
        }
        public static int countnum = 0;
        string content = string.Empty;
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (countnum == 3)
            {
                webBrowser1.Document.GetElementById("ReportViewerControl_ctl04_ctl00").InvokeMember("click");
                countnum++;
                Console.WriteLine("press search");
            }
            IHTMLDocument2 doc = webBrowser1.Document.DomDocument as IHTMLDocument2;
            content = doc.body.innerText;
        }

        private void WebBrowser1_NewWindow(Object sender, CancelEventArgs e)
        {

            System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
            messageBoxCS.AppendFormat("{0} = {1}", "Cancel", e.Cancel);
            messageBoxCS.AppendLine();
            MessageBox.Show(messageBoxCS.ToString(), "NewWindow Event");
        }


        int navigatecount = 1;
        void browser_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            if (navigatecount == 3)
            {
                webBrowser1.Document.GetElementById("ReportViewerControl_ctl04_ctl00").InvokeMember("click");
                Console.WriteLine("press search");
            }
            navigatecount++;
            Console.WriteLine(navigatecount.ToString());
        }


        private void exitBtn_Click(object sender, EventArgs e)
        {
            
            log.Close();
            try
            {
                //if (thread1.IsAlive)
                //{ thread1.Abort(); }
                if (File.Exists(filePath))//判断默认位置是否有excel
                {
                    //File.Delete(filePath);
                }
            }
            catch (Exception ex)
            { throw ex; }
            System.Environment.Exit(0);
        }//退出键
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void set_elemstylemike()
        {
            elem = this.webBrowser1.Document.GetElementById("ReportViewerControl_AsyncWait_Wait");
            if (elem != null)
            {
                elemstyle = elem.Style;
                Console.WriteLine(elemstyle + "this is another thread");
            }
        }

        //public void forthread()
        //{
        //    Console.WriteLine("thread start");
        //    for (int i = 1; i <= 100; i++)
        //    {
        //        Thread.Sleep(2000);
        //        this.Invoke(Set_Elemstyle);
        //    }
        //}

        private void button2_Click(object sender, EventArgs e)
        {
            //Uri uri = new Uri("http://gbs-sandbox/ReportServer?/CTS%20Reports/GBSDBI/SR%20Wellness/Case%20Wellness&LaborMins=0&UserRole=Team%20Manager&Alias=nichshen&Workgroup=GBS.OLSV.CN.APGC.CLOUD.CORE.SE.MS");
            //webBrowser1.Navigate(uri);
            //thread1 = new Thread(new ThreadStart(forthread));
            //thread1.Start();
        }
    }


}
