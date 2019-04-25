using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Xls;
using System.Threading;
using System.Net;
using System.Collections.Specialized;

namespace 电网自动化
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public class ExcelHelpers
        {
            #region 导入
            /// <summary>
            /// 将Excel以文件流转换DataTable
            /// </summary>
            /// <param name="hasTitle">是否有表头</param>
            /// <param name="path">文件路径</param>
            /// <param name="tableindex">文件簿索引</param>
            public  DataTable ExcelToDataTableFormPath(bool hasTitle = true, string path = "", int tableindex = 0)
            {
                //新建Workbook
                Workbook workbook = new Workbook();
                //将当前路径下的文件内容读取到workbook对象里面
                workbook.LoadFromFile(path);
                //得到第一个Sheet页
                Worksheet sheet = workbook.Worksheets[tableindex];
                return SheetToDataTable(hasTitle, sheet);
            }
            /// <summary>
            /// 将Excel以文件流转换DataTable
            /// </summary>
            /// <param name="hasTitle">是否有表头</param>
            /// <param name="stream">文件流</param>
            /// <param name="tableindex">文件簿索引</param>
            public  DataTable ExcelToDataTableFormStream(bool hasTitle = true, Stream stream = null, int tableindex = 0)
            {
                //新建Workbook
                Workbook workbook = new Workbook();
                //将文件流内容读取到workbook对象里面
                workbook.LoadFromStream(stream);
                //得到第一个Sheet页
                Worksheet sheet = workbook.Worksheets[tableindex];
                int iRowCount = sheet.Rows.Length;
                int iColCount = sheet.Columns.Length;
                DataTable dt = new DataTable();
                //生成列头
                for (int i = 0; i < iColCount; i++)
                {
                    var name = "column" + i;
                    if (hasTitle)
                    {
                        var txt = sheet.Range[1, i + 1].Text;
                        if (!string.IsNullOrEmpty(txt)) name = txt;
                    }
                    while (dt.Columns.Contains(name)) name = name + "_1";//重复行名称会报错。
                    dt.Columns.Add(new DataColumn(name, typeof(string)));
                }
                //生成行数据
                int rowIdx = hasTitle ? 2 : 1;
                for (int iRow = rowIdx; iRow <= iRowCount; iRow++)
                {
                    DataRow dr = dt.NewRow();
                    for (int iCol = 1; iCol <= iColCount; iCol++)
                    {
                        dr[iCol - 1] = sheet.Range[iRow, iCol].Text;
                    }
                    dt.Rows.Add(dr);
                }
                return SheetToDataTable(hasTitle, sheet);
            }
            private  DataTable SheetToDataTable(bool hasTitle, Worksheet sheet)
            {
                int iRowCount = sheet.Rows.Length;
                int iColCount = sheet.Columns.Length;
                var dt = new DataTable();
                //生成列头
                for (var i = 0; i < iColCount; i++)
                {
                    var name = "column" + i;
                    if (hasTitle)
                    {
                        var txt = sheet.Range[1, i + 1].Text;
                        if (!string.IsNullOrEmpty(txt)) name = txt;
                    }
                    while (dt.Columns.Contains(name)) name = name + "_1";//重复行名称会报错。
                    dt.Columns.Add(new DataColumn(name, typeof(String)));
                }
                //生成行数据
                // ReSharper disable once SuggestVarOrType_BuiltInTypes
                var rowIdx = hasTitle ? 2 : 1;
                for (var iRow = rowIdx; iRow <= iRowCount; iRow++)
                {
                    var dr = dt.NewRow();
                    for (var iCol = 1; iCol <= iColCount; iCol++)
                    {
                       
                         dr[iCol - 1] = sheet.Range[iRow, iCol].Value; //原来是Text
                        
                    }
                    dt.Rows.Add(dr);
                }
                return dt;
            }
            #endregion
            #region 导出
            /// <summary>
            /// 将DaTaTable转成byte[]类型
            /// </summary>
            /// <param name="dt"></param>
            /// <param name="hasTitle">是否有表头</param>
            /// <returns></returns>
            public void GetDataTableToByte(DataTable dt, bool hasTitle, string filename)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    Workbook workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];//第一个工作簿
                    //sheet.Name = "";
                    if (hasTitle) //表头
                    {
                        for (var j = 0; j < dt.Columns.Count; j++)
                        {
                            sheet.Range[1, j + 1].Text = dt.Columns[j].ColumnName;
                            sheet.Range[1, j + 1].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;//边框
                            sheet.Range[1, j + 1].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;//边框
                            sheet.Range[1, j + 1].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;//边框
                            sheet.Range[1, j + 1].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;//边框
                        }
                    }
                    //循环表数据
                    for (var i = 0; i < dt.Rows.Count; i++)//循环赋值
                    {
                        for (var j = 0; j < dt.Columns.Count; j++)
                        {
                            var dyg = sheet.Range[i + 2, j + 1];
                            dyg.Text = dt.Rows[i][j].ToString();
                            dyg.Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;//边框
                            dyg.Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
                            dyg.Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
                            dyg.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                        }
                    }
                    sheet.AllocatedRange.AutoFitColumns();//自动调整列的宽度去适应单元格的数据
                    //保存到物理路径
                    workbook.SaveToFile(filename, Spire.Xls.FileFormat.Version2007);
                    MessageBox.Show(filename+"导出ok!");
                }
            }
            #endregion

            #region 合并
            public void ExcelInsert(DataTable dt,string newpath,bool hasTitle = true, string path = "", int tableindex = 0)
            {
                //新建Workbook
                Workbook workbook = new Workbook();
                //将当前路径下的文件内容读取到workbook对象里面
                workbook.LoadFromFile(path);
                //得到第一个Sheet页
                Worksheet sheet = workbook.Worksheets[tableindex];
                
                int org_RowsLength = sheet.Rows.Length;
                //循环表数据
                for (var i = 0; i < dt.Rows.Count; i++)//循环赋值
                {
                    for (var j = 0; j < dt.Columns.Count; j++)
                    {
                        var dyg = sheet.Range[i + org_RowsLength + 1, j + 1];
                        dyg.Text = dt.Rows[i][j].ToString();
                        dyg.Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;//边框
                        dyg.Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
                        dyg.Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
                        dyg.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                    }
                }
                sheet.AllocatedRange.AutoFitColumns();//自动调整列的宽度去适应单元格的数据
                 //保存到物理路径

               workbook.SaveToFile(newpath, Spire.Xls.FileFormat.Version2007);
               MessageBox.Show(newpath + "导出ok!");
            }

            #endregion
        }
        DataTable dt,edt;
        string[] wordfiles = null;
        private void Form1_Load(object sender, EventArgs e)
        {
            //构建表格
            dt = new DataTable("dt");
            dt.Columns.Add("收货方", typeof(String));
            dt.Columns.Add("订单编号", typeof(String));
            dt.Columns.Add("工程账号", typeof(String));
            dt.Columns.Add("项目名称", typeof(String));
            dt.Columns.Add("货物名称", typeof(String));
            dt.Columns.Add("数量", typeof(String));
            dt.Columns.Add("合计金额", typeof(String));

            edt= new DataTable("edt");
            edt.Columns.Add("contractcode", typeof(String));
            edt.Columns.Add("projectcode", typeof(String));
            edt.Columns.Add("projectname", typeof(String));
            edt.Columns.Add("a", typeof(String));
            edt.Columns.Add("b", typeof(String));
            edt.Columns.Add("g", typeof(String));
            edt.Columns.Add("c", typeof(String));
            edt.Columns.Add("d", typeof(String));
            edt.Columns.Add("e", typeof(String));
            edt.Columns.Add("f", typeof(String));
            edt.Columns.Add("totalsum", typeof(String));
            edt.Columns.Add("info1", typeof(String));
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //定义打开的默认文件夹位置
            openFileDialog1.Filter = "Word文件(*.doc,docx)|*.doc;*.docx";
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label4.Text = "0/0";
                label5.Text = "0";
                label7.Text = "共:";
                label3.Text = "0";
                wordfiles = null;
                this.progressBar1.Value = 0;
                this.textBox1.Text = null;
                dt.Rows.Clear();
                this.dataGridView1.DataSource = null;
                this.label1.Text = openFileDialog1.FileName.Replace(openFileDialog1.SafeFileName, "");
                wordfiles = Directory.GetFiles(this.label1.Text, "*.docx");
                this.label3.Text = wordfiles.Count().ToString() + "个";
                this.progressBar1.Maximum = wordfiles.Count();//设置最大长度值
                this.progressBar1.Value = 0;//设置当前值
                this.progressBar1.Step = 1;//设置没次增长多少
                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
                sw.Start();

                for (int i = 0; i < wordfiles.Count(); i++)
                {
                    this.textBox1.Text += wordfiles[i].Replace(this.label1.Text, "") + "\r\n";
                    uiadd(i);
                    Thread thread = new Thread(() => Loadword(i));
                    thread.Start();
                    thread.Join();
                    thread.Abort();
                }
                this.dataGridView1.DataSource = dt;
                //宽度自适应
                this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
                label7.Text = "共:" + dt.Rows.Count.ToString() + "行";
                sw.Stop();
                TimeSpan ts2 = sw.Elapsed;
                label5.Text = (ts2.TotalMilliseconds / 1000).ToString() + "秒";
            }
        }
        public void Loadword(int i)
        {
            try
            {
                Document document = new Document(wordfiles[i]);
                string Str = document.GetText().Trim();
                Table tb1 = document.Sections[0].Tables[1] as Table;
                //MessageBox.Show(tb1.Rows.Count.ToString());//加表头最小3
                for (int j = 0; j < tb1.Rows.Count - 2; j++)
                {
                    int start = Str.IndexOf("订单编号：");
                    string 订单编号 = Str.Substring(start + 5, 10);
                    //  MessageBox.Show(tb1.Rows[0].Cells[3].Paragraphs[0].Text);
                    string 收货方 = tb1.Rows[j + 1].Cells[1].Paragraphs[0].Text.Replace("上海市电力公司", "").Replace("供电公司", "");
                    string 工程账号 = "";
                    string 项目名称 = tb1.Rows[j + 1].Cells[2].Paragraphs[0].Text;
                    string 货物名称 = tb1.Rows[j + 1].Cells[3].Paragraphs[0].Text;
                    string 数量 = tb1.Rows[j + 1].Cells[5].Paragraphs[0].Text;
                    string 合计金额 = tb1.Rows[j + 1].Cells[9].Paragraphs[0].Text;
                    lock (dt.Rows.SyncRoot)
                    {
                        dt.Rows.Add(收货方, 订单编号, 工程账号, 项目名称, 货物名称, 数量, 合计金额);
                    }

                }
            }
            catch (Exception exp) { MessageBox.Show("i:" + i.ToString() + " " + exp.ToString()); }
        }
        public void uiadd(int i)
        {
            this.progressBar1.Invoke(new MethodInvoker(() =>
            {
                this.progressBar1.Value += this.progressBar1.Step;
            }));
            this.label4.Invoke(new MethodInvoker(() =>
            {
                this.label4.Text = (i + 1).ToString() + "/" + this.label3.Text;
            }));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
                ExcelHelpers excelHelpers = new ExcelHelpers();
                excelHelpers.GetDataTableToByte(dt, true, this.label1.Text + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            }
            else
            {
                MessageBox.Show("数据不能为空");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //定义打开的默认文件夹位置
                openFileDialog1.Filter = "Excel文件(*.xls,xlsx)|*.xls;*.xlsx";
                openFileDialog1.RestoreDirectory = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    this.label8.Text = openFileDialog1.FileName;
                    ExcelHelpers excelHelpers = new ExcelHelpers();
                    string path = System.IO.Path.GetDirectoryName(openFileDialog1.FileName);
                    string filename = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
                    string extension = System.IO.Path.GetExtension(openFileDialog1.FileName);
                    string newnfilenames = path + '\\' + filename + DateTime.Now.ToString("yyyyMMddHHmmss") + extension;
                    //MessageBox.Show(newnfilenames);
                    excelHelpers.ExcelInsert(dt, newnfilenames, false, this.label8.Text);

                }
               
            }
            else
            {
                MessageBox.Show("数据不能为空");
            }


        }

        public void wf5dd(DataTable exceldataTable)
        {
            var client = new CookieAwareWebClient();
            client.BaseAddress = @"http://60.60.60.202/OA/";
            var loginData = new NameValueCollection();
            loginData.Add("name", "621957");
            loginData.Add("pwd", "S56650088");
            client.UploadValues("login/login", "POST", loginData);
            //Now you are logged in and can request pages
           // Encoding ec = Encoding.Default;
           // byte[] btArr = ec.GetBytes(client.DownloadString("index/index"));
           // string strBuffer = Encoding.UTF8.GetString(btArr);
           // this.textBox3.Text = strBuffer;

            for (int i = 0; i < exceldataTable.Rows.Count; i++)
            {

                var values = new NameValueCollection();
                values["department"] = "电网市场部";
                values["handler"] = "杨朔";
                values["contractcode"] = exceldataTable.Rows[i]["contractcode"].ToString();
                values["projectcode"] = exceldataTable.Rows[i]["projectcode"].ToString();
                values["projectname"] = exceldataTable.Rows[i]["projectname"].ToString();
                values["prognum"] = "";
                values["taxrate"] = "16";
                values["company2"] = "国网上海市电力公司";
                values["taxnumber2"] = "91310101132224671B";
                values["address2"] = "上海市浦东新区源深路1122号";
                values["bank2"] = "中国工商银行股份有限公司上海市分行营业部";
                values["tel2"] = "021-28925222";
                values["account2"] = "1001254029003452681";
                values["people2"] = "";
                values["company1"] = "上海电力环保设备总厂有限公司";
                values["taxnumber1"] = "91310113133030473E";
                values["address1"] = "上海市宝山区真陈路889号";
                values["bank1"] = "中国银行上海市桃浦支行";
                values["tel1"] = "56655880";
                values["account1"] = "442968467193";
                values["people1"] = "";

                //判断多行
                int j = i;
                int n = 0;
                for (j = i; j < edt.Rows.Count; j++)
                {
                    if (j < edt.Rows.Count - 1)
                    {
                        if (edt.Rows[j]["info1"].ToString() == edt.Rows[j + 1]["info1"].ToString())
                        { n++; }
                        else { break; }
                    }
                    else
                    { break; }
                }
                this.textBox3.Text += edt.Rows[i]["info1"].ToString() + " " + (n + 1).ToString() + "条\r\n";

                for (int k = 0; k < n+1; k++)
                {
                    values["a[" + k + "]"] = exceldataTable.Rows[i + k]["a"].ToString();
                    values["b[" + k + "]"] = exceldataTable.Rows[i + k]["b"].ToString();
                    values["c[" + k + "]"] = exceldataTable.Rows[i + k]["c"].ToString();
                    values["d[" + k + "]"] = exceldataTable.Rows[i + k]["d"].ToString();
                    values["e[" + k + "]"] = exceldataTable.Rows[i + k]["e"].ToString();
                    values["f[" + k + "]"] = exceldataTable.Rows[i + k]["f"].ToString();
                    values["g[" + k + "]"] = exceldataTable.Rows[i + k]["g"].ToString();
                }
                /*
                int k = n;
                int l = 0;
                do
                {
                    values["a[" + l + "]"] = exceldataTable.Rows[i + l]["a"].ToString();
                    values["b[" + l + "]"] = exceldataTable.Rows[i + l]["b"].ToString();
                    values["c[" + l + "]"] = exceldataTable.Rows[i + l]["c"].ToString();
                    values["d[" + l + "]"] = exceldataTable.Rows[i + l]["d"].ToString();
                    values["e[" + l + "]"] = exceldataTable.Rows[i + l]["e"].ToString();
                    values["f[" + l + "]"] = exceldataTable.Rows[i + l]["f"].ToString();
                    values["g[" + l + "]"] = exceldataTable.Rows[i + l]["g"].ToString();
                    k--; l++;
                } while (k > 0);
                */
                if (n > 0)
                {
                    i = i + n;
                }


                values["totalsum"] = exceldataTable.Rows[i]["totalsum"].ToString();
                values["xiaozhang"] = "0";
                values["info1"] = exceldataTable.Rows[i]["info1"].ToString();
                values["depboss"] = "赵一鸣";
                values["level2"] = "施宏";
                values["save"] = "0";
                values["submit1"] = "启动流程";
                var response=client.UploadValues("wf5/addwf5", "POST", values);
                //string responseString = Encoding.UTF8.GetString(response);
                // string responseString = Encoding.Default.GetString(response);
                //MessageBox.Show(responseString);
                // this.label9.Text = responseString;
                //this.textBox3.Text = responseString;
                this.label9.Invoke(new MethodInvoker(() =>
                {
                    this.label9.Text = (i + 1).ToString() + "/" + exceldataTable.Rows.Count;
                }));
            }
        }

        public void wf5dd(DataTable exceldataTable,string cookie)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");//采取POST方式必须加的header，如果改为GET方式的话就去掉这句话即可  
                client.Headers.Add("User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.75 Safari/537.36");
                client.Headers.Add(HttpRequestHeader.Cookie, cookie);
                for (int i = 0; i < exceldataTable.Rows.Count; i++)
                {
                    var values = new NameValueCollection();
                    values["department"] = "电网市场部";
                    values["handler"] = "杨朔";
                    values["contractcode"] = exceldataTable.Rows[i]["contractcode"].ToString();
                    values["projectcode"] = exceldataTable.Rows[i]["projectcode"].ToString();
                    values["projectname"] = exceldataTable.Rows[i]["projectname"].ToString();
                    values["prognum"] = "";
                    values["taxrate"] = "16";
                    values["company2"] = "国网上海市电力公司";
                    values["taxnumber2"] = "91310101132224671B";
                    values["address2"] = "上海市浦东新区源深路1122号";
                    values["bank2"] = "中国工商银行股份有限公司上海市分行营业部";
                    values["tel2"] = "021-28925222";
                    values["account2"] = "1001254029003452681";
                    values["people2"] = "";
                    values["company1"] = "上海电力环保设备总厂有限公司";
                    values["taxnumber1"] = "91310113133030473E";
                    values["address1"] = "上海市宝山区真陈路889号";
                    values["bank1"] = "中国银行上海市桃浦支行";
                    values["tel1"] = "56655880";
                    values["account1"] = "442968467193";
                    values["people1"] = "";
                    values["a[]"] = exceldataTable.Rows[i]["a"].ToString();
                    values["b[]"] = exceldataTable.Rows[i]["b"].ToString();
                   
                    values["c[]"] = exceldataTable.Rows[i]["c"].ToString();
                    values["d[]"] = exceldataTable.Rows[i]["d"].ToString();
                    values["e[]"] = exceldataTable.Rows[i]["e"].ToString();
                    values["f[]"] = exceldataTable.Rows[i]["f"].ToString();
                    values["g[]"] = exceldataTable.Rows[i]["g"].ToString();
                    values["totalsum"] = exceldataTable.Rows[i]["totalsum"].ToString();
                    values["xiaozhang"] = "0";
                    values["info1"] = exceldataTable.Rows[i]["info1"].ToString();
                    values["depboss"] = "赵一鸣";
                    values["level2"] = "施宏";
                    values["save"] = "0";
                    values["submit1"] = "启动流程";
                    //client.UploadValuesCompleted += (sender, e) => { this.label9.Text = (i).ToString() + "/" + exceldataTable.Rows.Count; };
                    client.UploadValues("http://60.60.60.202/OA/wf5/addwf5", values);
                    

                   // string responseString = Encoding.Default.GetString(response);
                    //MessageBox.Show(responseString);
                   // this.label9.Text = responseString;
                    this.label9.Invoke(new MethodInvoker(() =>
                    {
                        this.label9.Text = (i+1).ToString() + "/" + exceldataTable.Rows.Count;
                    }));
                }
               
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            wf5dd(edt);
            MessageBox.Show("ok");
        }

        public class CookieAwareWebClient : WebClient
        {
            private CookieContainer cookie = new CookieContainer();
            protected override WebRequest GetWebRequest(Uri address)
            {
                WebRequest request = base.GetWebRequest(address);
                if (request is HttpWebRequest)
                {
                    (request as HttpWebRequest).CookieContainer = cookie;
                }
                
                return request;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

            var client = new CookieAwareWebClient();
            client.BaseAddress = @"http://60.60.60.202/OA/";
            var loginData = new NameValueCollection();
            loginData.Add("name", "621957");
            loginData.Add("pwd", "S56650088");
            client.UploadValues("login/login", "POST", loginData);
            //Now you are logged in and can request pages
            Encoding ec = Encoding.Default;
            byte[] btArr = ec.GetBytes(client.DownloadString("index/index"));
            
            string strBuffer = Encoding.UTF8.GetString(btArr);

            this.textBox3.Text = strBuffer;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < edt.Rows.Count; i++)
            {
                //判断多行
                int j = i;
                int n = 0;
                for(j=i;j< edt.Rows.Count;j++)
                {
                    if (j < edt.Rows.Count-1)
                    {
                        if (edt.Rows[j]["info1"].ToString() == edt.Rows[j + 1]["info1"].ToString())
                        { n++; }
                        else{break;}
                    }
                    else
                    { break;}
                }
                this.textBox3.Text += edt.Rows[i]["info1"].ToString()+" "+(n+1).ToString()+"条\r\n";
                if (n > 0)
                {
                    i = i + n;
                }  
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //定义打开的默认文件夹位置
            openFileDialog1.Filter = "Excel文件(*.xls,xlsx)|*.xls;*.xlsx";
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                edt.Rows.Clear();
                ExcelHelpers excelHelpers = new ExcelHelpers();
              
                edt = excelHelpers.ExcelToDataTableFormPath(true, openFileDialog1.FileName, 0);
                this.dataGridView2.DataSource= edt;

            }
        }
    }
}
