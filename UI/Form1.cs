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
using System.Threading;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Data.SQLite;
using System.Web;
using System.Collections;
using System.Diagnostics;
using System.Windows.Forms.DataVisualization.Charting;
//for AI Regression Predict
using Microsoft.ML;


namespace CrawlerCSV
{
    public partial class Form1 : Form
    {
       
        public String connStr = "Data Source=stock.db;Pooling=true;FailIfMissing=false";
        Dictionary<String, String> stock_list = new Dictionary<string, String>();

        //右鍵選單新增歷史k線圖
        public string clickedNode = "";
        public MenuItem item1 = new MenuItem("深度分析");
        public MenuItem item2 = new MenuItem("技術分析");
        public ContextMenu menu = new ContextMenu();

        public Form1()
        {
            InitializeComponent();
            //預設先取得所有目前股票列表
            String[] db_stock_list = DB_SQL("select stock_code,stock_name from stock_profile order by stock_code ","stock_code,stock_name").Split(';');

            menu.MenuItems.Add(item1);
            menu.MenuItems.Add(item2);
            item1.Click += new EventHandler(item1_Click);
            item2.Click += new EventHandler(item2_Click);

            //新增至列表
            TreeNode etfNode = new TreeNode("ETF");
            TreeNode publicNode = new TreeNode("上市公司");
            treeView1.Nodes.Add(etfNode);
            treeView1.Nodes.Add(publicNode);
            int etfCount = 0;
            int publicCount = 0;
            for (int i = 0; i < db_stock_list.Length - 1; i++)
            {
                stock_list.Add(db_stock_list[i].Split('@')[0], db_stock_list[i].Split('@')[1]);
                if (db_stock_list[i].Split('@')[0].Length > 4 || db_stock_list[i].Split('@')[0].StartsWith("00"))
                {
                    etfNode.Nodes.Add(String.Format("[{0}]{1}", db_stock_list[i].Split('@')[0], db_stock_list[i].Split('@')[1]));
                    etfCount++;
                }
                else
                {
                    publicNode.Nodes.Add(String.Format("[{0}]{1}", db_stock_list[i].Split('@')[0], db_stock_list[i].Split('@')[1]));
                    publicCount++;
                }
            }
            //預設展開列表
            publicNode.Expand();
            //統計目前共有多少檔
            etfNode.Text = String.Format("ETF[共{0}檔]",etfCount);
            publicNode.Text = String.Format("上市公司[共{0}檔]", publicCount);

            ProcessStartInfo Info = new ProcessStartInfo();
            Info.FileName = "intro.vbs";//執行的檔案名稱
                                        //Info.WorkingDirectory = @"E:\NVR";//檔案所在的目錄
                                        //Process.Start(Info);// RUN bat

            AutoCompleteStringCollection acc = new AutoCompleteStringCollection();
            foreach (var item in stock_list)
            {
                acc.Add(item.Key+" "+item.Value);
                acc.Add(item.Value+" "+item.Key);
            }
            textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            textBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox1.AutoCompleteCustomSource = acc;

        }

        void item1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage2;
            label1.Visible = true;
            //button1.Visible = true;
            button9.Visible = true;
            textBox1.Visible = true;

            label2.Visible = false;
            
            textBox3.Visible = false;
            label3.Visible = false;
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            textBox1.Text = clickedNode.Replace("[", "").Replace("]", " ");
        }

        void item2_Click(object sender, EventArgs e)
        {
            textBox1.Text = clickedNode.Replace("[", "").Replace("]", " ");
            String stock_code = textBox1.Text.Split(' ')[0];
            MessageBox.Show(stock_code, textBox1.Text.Split(' ')[1]);

            label7.Text = textBox1.Text;
            
            //錯誤寫法 ： String sql = "select max(close) max_price,min(close) min_price from stock_price_new where stock_code = '" + stock_code + "' order by compare_date desc limit 0,90";
            String sql = "select max(close) max_price, min(close) min_price from(select* from stock_price_new where stock_code = '" + stock_code + "'  limit 0,3000)";
            //取得最大價格、最低價格
            String max_min = DB_SQL(sql, "max_price,min_price");
            //MessageBox.Show(max_min);
            //設定最大值及最小值
            chart1.ChartAreas["ChartArea1"].AxisY.Maximum = Convert.ToDouble(max_min.Split('@')[0])*1.03;
            chart1.ChartAreas["ChartArea1"].AxisY.Minimum = Convert.ToDouble(max_min.Split('@')[1])*0.97;

            clickedNode = "";
           
            tabControl1.SelectedTab = tabPage6;
            chart1.ChartAreas["ChartArea1"].AxisX.MajorGrid.LineWidth = 1;
            chart1.ChartAreas["ChartArea1"].AxisY.MajorGrid.LineWidth = 1;

            chart1.ChartAreas["ChartArea1"].AxisX.Interval = 240;
            //chart1.ChartAreas["ChartArea1"].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.WordWrap;
            //chart1.ChartAreas["ChartArea1"].AxisX.IsLabelAutoFit = true;
            //chart1.ChartAreas["ChartArea1"].AxisX.LabelStyle.Enabled = true;

            chart1.Series[0].XValueMember = "stock_date";
            chart1.Series[0].YValueMembers = "high,low,open,close";//順序不可以改!!!!!
            chart1.Series[0].XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Date;
            chart1.Series[0].CustomProperties = "PriceDownColor=Green,PriceUpColor=Red";

            chart1.Series[0]["OpenCloseStyle"] = "Triangle";
            chart1.Series[0]["ShowOpenClose"] = "Both";
            chart1.DataManipulator.IsStartFromFirst = true;

            //取得最近九十個交易日的資料
            sql = "select * from stock_price_new where stock_code='"+ stock_code + "' order by compare_date desc limit 0,3000 ";
            //灌入資料時發現時間日期順序不對 -> 調整sql語法
            sql = "select * from (select * from stock_price_new where stock_code='" + stock_code + "' order by compare_date desc limit 0,3000 ) order by compare_date asc";

            DataTable newTable = new DataTable();
            //table = newTable;
            using (SQLiteConnection conn = new SQLiteConnection(connStr))
            {
                conn.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                {
                    using (SQLiteDataReader dr = cmd.ExecuteReader())
                    {
                        newTable.Load(dr);
                    }
                }
            }

            chart1.DataSource = newTable;
            chart1.DataBind();
            
        }


        Point? prevPosition = null;
        ToolTip tooltip = new ToolTip();

        //參考資料https://stackoverflow.com/questions/33978447/display-tooltip-when-mouse-over-the-line-chart/35593876
        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            Point mousePoint = new Point(e.X, e.Y);
            chart1.ChartAreas["ChartArea1"].CursorX.SetCursorPixelPosition(mousePoint, true);
            chart1.ChartAreas["ChartArea1"].CursorY.SetCursorPixelPosition(mousePoint, true);

            var pos = e.Location;
            if (prevPosition.HasValue && pos == prevPosition.Value)
                return;
            tooltip.RemoveAll();
            prevPosition = pos;
            //判斷是否移到點位並取得點位的資料
            var results = chart1.HitTest(pos.X, pos.Y, false, ChartElementType.DataPoint); // set ChartElementType.PlottingArea for full area, not only DataPoints
            foreach (var result in results)
            {
                if (result.ChartElementType == ChartElementType.DataPoint) // set ChartElementType.PlottingArea for full area, not only DataPoints
                {
                    //取得點位資料 
                    var data = result.Object as DataPoint;
                    tooltip.Show(String.Format("High:{0}\nLow:{1}\nOpen:{2}\nClose:{3}", data.YValues[0], data.YValues[1], data.YValues[2], data.YValues[3]), chart1, pos.X, pos.Y - 15);
                }
            }
            
        }




        private void TreeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                clickedNode = e.Node.Text;

                menu.Show(treeView1, e.Location);
            }
        }

        public void updateMsgNew(String msg)
        {
            if (textBox4.IsDisposed || !textBox4.Parent.IsHandleCreated) return;//C# 在建立視窗控制代碼之前,不能在控制元件上呼叫 Invoke 或 BeginInvoke
            this.Invoke((MethodInvoker)delegate
            {
                if (msg == "clear")
                    textBox2.Text = "";
                else
                {
                    textBox4.AppendText(msg);
                    textBox4.AppendText(Environment.NewLine);
                }
            });
        }

        public void updateMsg(String msg)
        {
            if (textBox2.IsDisposed || !textBox2.Parent.IsHandleCreated) return;//C# 在建立視窗控制代碼之前,不能在控制元件上呼叫 Invoke 或 BeginInvoke
            this.Invoke((MethodInvoker)delegate
            {
                if (msg == "clear")
                    textBox2.Text = "";
                else
                {
                    textBox2.AppendText(msg);
                    textBox2.AppendText(Environment.NewLine);
                }
            });
        }

        //DB_SQL insert/update/delete
        public String DB_SQL(String sql)
        {
            int result = -1;
            String strResult = String.Empty;
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    if (conn.State == ConnectionState.Open) conn.Close();
                    conn.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                    {
                        result = cmd.ExecuteNonQuery();
                        strResult = result.ToString();
                    }
                }
            }
            catch (Exception e)
            {
                strResult = e.Message;
                //strResult = "0";
            }

            return strResult;
        }

        //DB_SQL read data
        public String DB_SQL(String sql, String column)
        {
            String strResult = String.Empty;
            StringBuilder sb = new StringBuilder();
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                    {
                        using (SQLiteDataReader DB_Reader = cmd.ExecuteReader())
                        {
                            try
                            {
                                if (column.Contains(','))
                                {
                                    String[] cols = column.Split(',');
                                    String item = "";
                                    while (DB_Reader.Read())
                                    {
                                        for (int i = 0; i < cols.Length; i++)
                                        {
                                            item = DB_Reader[cols[i]].ToString();
                                            sb.Append(item + "@");
                                        }
                                        sb.Append(";");
                                    }
                                }
                                else
                                {
                                    String item = "";
                                    while (DB_Reader.Read())
                                    {
                                        item = DB_Reader[column].ToString();
                                        if (strResult.Length == 0)
                                            strResult = item;
                                        else
                                            strResult = strResult + "," + item;
                                    }
                                    return strResult;
                                }
                                strResult = sb.ToString();
                            }
                            catch (Exception e)
                            {
                                strResult = e.Message;
                            }
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception e)
            {
                strResult = e.Message;
            }
            return strResult;
        }

        public Dictionary<String,String> getETF()
        {
            Dictionary<String, String> etf_list = new Dictionary<String, String>();
            try
            {
                //parse etf
                String url = "https://www.twse.com.tw/zh/page/ETF/list.html";
                String xpath = "//*[@id=\"main\"]/article/section/table/tbody";
                WebClient client = new WebClient();
                //note：調整protocol
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                MemoryStream ms = new MemoryStream(client.DownloadData(url));

                // 使用預設編碼讀入 HTML 
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.UTF8);

                // 裝載第一層查詢結果 
                HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();
                // table / tbody / tr[1] / td[2]
                hdc.LoadHtml(doc.DocumentNode.SelectSingleNode(xpath).InnerHtml);

                // 取得個股標頭 
                HtmlNodeCollection nodeHeaders = hdc.DocumentNode.SelectNodes("./tr");
                string[] values = hdc.DocumentNode.SelectSingleNode("./tr[1]").InnerText.Trim().Split('\n');

                StringBuilder sb = new StringBuilder();
                

                for (int i = 0; i < nodeHeaders.Count; i++)
                //for (int i = 0; i < 2; i++)
                {
                    String stock_code = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[2]").InnerText.Trim();
                    String stock_name = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[3]").InnerText.Trim();
                    if (!etf_list.ContainsKey(stock_code))
                    {
                        etf_list.Add(stock_code, stock_name);
                    }
                }
                foreach (var item in etf_list)
                    sb.Append(String.Format("[{0}]{1}{2} ", item.Key, item.Value, Environment.NewLine));

                updateMsgNew(sb.ToString());
            }
            catch(Exception ee)
            {
                updateMsgNew("連線過度頻繁 被Ban惹! =>" + ee.Message);
            }

            return etf_list;
        }

        public Dictionary<String, String> getPublic()
        {
            Dictionary<String, String> public_list = new Dictionary<String, String>();
            try
            {
                //取得最新交易日期
                DateTime latest_date = getLatestDate();

                //parse public traded company
                String url = "http://www.twse.com.tw/exchangeReport/BWIBBU_d?response=html&date=" + latest_date.ToString("yyyyMMdd") + "&selectType=ALL";

                String xPath = "/html[1]/body[1]/div[1]/table[1]/tbody[1]";
                WebClient client = new WebClient();
                //note：調整protocol
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                MemoryStream ms = new MemoryStream(client.DownloadData(url));

                // 使用預設編碼讀入 HTML 
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.UTF8);

                // 裝載第一層查詢結果 
                HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();
                // table / tbody / tr[1] / td[2]
                hdc.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);

                // 取得個股標頭 
                HtmlNodeCollection nodeHeaders = hdc.DocumentNode.SelectNodes("./tr");
                string[] values = hdc.DocumentNode.SelectSingleNode("./tr[1]").InnerText.Trim().Split('\n');

                for (int i = 0; i < nodeHeaders.Count; i++)
                {
                    String stock_code = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[1]").InnerText.Trim();
                    String stock_name = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[2]").InnerText.Trim();
                    if (!public_list.ContainsKey(stock_code))
                    {
                        public_list.Add(stock_code, stock_name);
                    }
                    updateMsgNew(String.Format("[{0}] {1}{2}", stock_code, stock_name, Environment.NewLine));
                }
            }
            catch (Exception ee)
            {
                updateMsgNew("[錯誤] =>" + ee.Message);
            }

            return public_list;
        }


        public void getStockList()
        {
            try
            {
                stock_list = getETF();
                int etf_num = stock_list.Count;

                DateTime a = DateTime.Now;

                //String url = "http://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=html";
                //String url = "http://www.twse.com.tw/exchangeReport/BWIBBU_d?response=html&date="+a.ToString("yyyyMMdd")+"&selectType=ALL";
                String url = "http://www.twse.com.tw/exchangeReport/BWIBBU_d?response=html&date=" + a.AddDays(-2).ToString("yyyyMMdd") + "&selectType=ALL";

                WebClient client = new WebClient();
                //note：調整protocol
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                MemoryStream ms = new MemoryStream(client.DownloadData(url));

                // 使用預設編碼讀入 HTML 
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.UTF8);

                // 裝載第一層查詢結果 
                HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();
                String msg = doc.DocumentNode.SelectSingleNode("/html/body/div").InnerHtml;
                if (msg.Contains("很抱歉，沒有符合條件的資料!"))
                    MessageBox.Show("很抱歉，沒有符合條件的資料!");
                else
                {
                    hdc.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/table[1]/tbody[1]").InnerHtml);
                    //hdc.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/table[1]/tbody[1]/tr[1]/td[2]").InnerHtml);
                    // 取得個股標頭 
                    HtmlNodeCollection nodeHeaders = hdc.DocumentNode.SelectNodes("./tr");

                    //string[] values = hdc.DocumentNode.SelectSingleNode("./tr[1]").InnerText.Trim().Split('\n');
                    updateMsgNew("正在讀取網頁清單 請稍後...");
                    // 輸出資料 
                    updateMsgNew(String.Format("總計{0}筆股票", nodeHeaders.Count));
                    updateMsgNew("=========================");
                    //string[] values = hdc.DocumentNode.SelectSingleNode("./tr[1]").InnerText.Trim();
                    StringBuilder sb = new StringBuilder();
                    StringBuilder sb_sql = new StringBuilder();
                    a = DateTime.Now;
                    TimeSpan t = DateTime.Now - a;
                    updateMsgNew("等待網頁回應資料花費時間" + t.TotalSeconds + "秒");


                    for (int i = 0; i < nodeHeaders.Count; i++)
                    //for (int i = 0; i < 2; i++)
                    {
                        String stock_code = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[1]").InnerText.Trim();
                        String stock_name = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[2]").InnerText.Trim();
                        if (!stock_list.ContainsKey(stock_code))
                        {
                            stock_list.Add(stock_code, stock_name);
                        }
                    }

                    String sql = "";
                    String updateNum = "";
                    String date = DateTime.Now.ToString("yyyy/MM/dd");
                    foreach (var item in stock_list)
                    {
                        try
                        {
                            sb.Append(String.Format("[{0}]{1}{2} ", item.Key, item.Value, Environment.NewLine));
                            sql = String.Format("INSERT OR replace into stock_profile (stock_code,stock_name,update_time) values ('{0}','{1}','{2}') ", item.Key, item.Value, date);
                            updateNum = DB_SQL(sql);
                            if (updateNum.All(Char.IsDigit))
                            {
                                if (Convert.ToInt32(updateNum) > 0)
                                    updateMsgNew("[" + item.Key + "]寫入資料庫成功(" + updateNum + ")筆");
                                else
                                    updateMsgNew("[" + item.Key + "]寫入資料庫失敗(" + updateNum + ")筆");
                            }
                            else
                                if (updateNum.ToLower().Contains("unique"))
                                    updateMsgNew("資料重覆->更新資料");
                                else
                                    updateMsgNew("[錯誤]" + updateNum);
                        }
                        catch (Exception eex)
                        {
                            updateMsgNew(eex.StackTrace.ToString());
                        }
                    }
                    updateMsgNew(sb.ToString());
                    updateMsgNew(String.Format("總計{0}筆ETF {1}筆上市股票 總計{2}筆", etf_num, nodeHeaders.Count, etf_num + nodeHeaders.Count));
                    //{1}已完成下載 {2}尚未完成下載
                    //button3.Visible = true;
                }
            }
            catch (Exception ee)
            {
                if (ee.Message.Contains("連線被拒"))
                    updateMsg("連線過度頻繁 被Ban惹! =>" + ee.Message);
                else
                    updateMsg(ee.Message);
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            new Thread(() =>
            {
                getStockList();
            });
        }

        public String getStockDate(String stock_code)
        {
            String ans = "";
            String url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=html&date=10010101&stockNo=" + stock_code;
            try
            { 
                //DateTime sDate = DateTime.Now;
                WebClient client = new WebClient();
                //note：調整protocol
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                MemoryStream ms = new MemoryStream(client.DownloadData(url));

                // 使用預設編碼讀入 HTML 
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.UTF8);

                // 裝載第一層查詢結果 
                HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();
                String msg = doc.DocumentNode.SelectSingleNode("/html/body/div").InnerHtml;
                if (msg.Contains("查詢日期小於"))
                {
                    ans = stock_code+"=>"+msg.Trim().Substring(6).Split('，')[0];
                }
            }
            catch(Exception ee)
            {
                updateMsg("連線過度頻繁 被Ban惹! =>" + ee.Message);
            }
            return ans;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            new Thread(() =>
            {
                String sql = "";

                //紀錄trackBar值
                int trackBarValue = 0;
                this.Invoke(new MethodInvoker(delegate { trackBarValue = trackBar1.Value; }));


                String result = "";
                //if (tabControl1.SelectedTab == tabPage5)
                //{
                this.Invoke((MethodInvoker)delegate {
                    label3.Text = String.Format("說明：挑出10年內(1)連續{0}年發股息 (2)全部填權息 ☆(3)WADAR指標前五十名 『總排行榜』", trackBarValue);
                    label3.Visible = true;
                });

                //更新GridView
                sql = " select stock_code,stock_name,win_times,dividen_times,avg_win_rate,dividen from money_rank_new " +
                    " where dividen_times >= " + trackBarValue +
                    " order by (win_times/dividen_times) desc ,avg_win_rate desc,dividen_times desc limit 0,50";
                result = DB_SQL(sql, "stock_code,stock_name,win_times,dividen_times,avg_win_rate,dividen");

                CreateGVData(dataGridView2, sql);
                //}
                //else if (tabControl1.SelectedTab == tabPage3)
                //{
                this.Invoke((MethodInvoker)delegate {
                    label3.Text = "說明：挑出10年內(1)至少連續" + trackBarValue + "年發股息 (2)全部填權息 ☆(3)WADAR指標前五十名 (4)尚未填息 的『可買進名單』";
                    label3.Visible = true;
                });

                //更新GridView
                sql = "select m.stock_code,m.stock_name,m.win_times,m.dividen_times,m.avg_win_rate,p.close,d.before_price, " +
                           "  Round((d.before_price - p.close),2) price_diff, Round(100*(d.before_price - p.close) / before_price,3) price_percent,d.dividen_date,p.stock_date " +
                           "       from money_rank_new m left outer join latest_stock_price p " +
                           "      on m.stock_code = p.stock_code " +
                           "   left outer join latest_stock_dividen d on d.stock_code = m.stock_code " +
                           "   where m.dividen_times >=" + trackBarValue + " and p.close < d.before_price and m.win_times = m.dividen_times " +
                           "   order by(price_diff/ before_price) desc , m.avg_win_rate desc limit 0,50";
                result = DB_SQL(sql, "stock_code,stock_name,win_times,dividen_times,avg_win_rate,close,before_price,price_diff,price_percent,dividen_date,stock_date");

                CreateGVData(dataGridView1, sql);
            }).Start();
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show(String.Format("來抓除{0}權息價格", textBox2.Text),"開始發財$");

            //note：編碼 顯示中文
            //var rows = File.ReadAllLines("2018.csv.txt").Select(l => l.Split(',').ToArray()).ToArray();
            new Thread(() => {
                int idx = 0;        //往前剖析10年
                                    //for (idx = 1; idx <= 10; idx++)
                for (idx = 11; idx <=11; idx++)
                {
                    var rows = File.ReadAllLines(String.Format("{0}.txt", (2019 - idx)), Encoding.Default).Select(l => l.Split(',').ToArray()).ToArray();
                    //MessageBox.Show(String.Format("{0}年除權息總計 [{1}] rows", (2019-idx),rows.Length));
                    updateMsg("clear");
                    double before_price = 0.0;
                    double after_price = 0.0;
                    //組sql字串 
                    String sql = String.Empty;
                    String updateNum = "";
                    //for (int i = 0; i < 10; i++)
                    for (int i = 0; i < rows.Length; i++)
                    {
                        if (rows[i].Length == 17)
                            if (rows[i][6].Length > 0)
                            {
                                after_price = Convert.ToDouble(rows[i][6]);
                                before_price = after_price + Convert.ToDouble(rows[i][16]);

                                updateMsg(String.Format("[{0}]{1} 除權息日期[{2}]{6}除權後價格:{3} 股權股息:{4} 除權前價格:{5}", rows[i][1], rows[i][2], rows[i][5], rows[i][6], rows[i][16], before_price,Environment.NewLine));
                                sql = String.Format("insert into stock_dividen (stock_code,stock_name,dividen_date,before_price,after_price,dividen,compare_date) values ('{0}','{1}','{2}',{3},{4},{5},'{6}')", rows[i][1], rows[i][2], rows[i][5], before_price, rows[i][6], rows[i][16],rows[i][5].Replace("/",""));
                                //updateMsg(sql);
                                updateNum = DB_SQL(sql.ToString());
                                if (updateNum.All(Char.IsDigit))
                                {
                                    if (Convert.ToInt32(updateNum) > 0)
                                        updateMsg("寫入資料成功(" + updateNum.ToString() + ")");
                                    else
                                        updateMsg("寫入資料失敗(" + updateNum.ToString() + ")");
                                }
                                else
                                {
                                    if (updateNum.ToLower().Contains("unique"))
                                        updateMsg("資料重覆->更新資料");
                                    else
                                        updateMsg("[錯誤]" + updateNum);
                                }
                            }
                    }
                    updateMsg(String.Format("=================================={2}統計{0}年總共有{1}公司發放股息股利{2}==================================", 
                        (2019-idx),rows.Length,Environment.NewLine));
                    Thread.Sleep(1500);
                    //MessageBox.Show((2019 - idx) + "年資料剖析完畢", (2019 - idx)+"年總計"+rows.Length+"筆資料");

                }
                updateMsg("成功! 已抓取完近十年所有發股息股利資料!");
            }).Start();
            
        }

        public void LoadStock(String stock_code)
        {
            int success_cnt = 0;
            int fail_cnt = 0;

            //String stock_code = textBox1.Text;
            if (stock_code.Length <= 0)
                MessageBox.Show("未輸入代碼", "錯誤");
            else
            {
                String path = String.Format("c:\\CSV\\{0}\\{0}.csv", stock_code);
                if (File.Exists(path))
                {
                    var rows = File.ReadAllLines(path).Select(line => line.Split(',').ToArray()).ToArray();

                    updateMsg("clear");
                    updateMsg(String.Format("匯入代碼[{0}]歷史資料，總計{1}筆", stock_code, rows.Length));
                    String one_row = "";
                    bool isNull = false;
                    //NOTE i=0 is header so begin with i=1
                    for (int i = 1; i < rows.Length; i++)
                    {
                        one_row = "";
                        //check if contains null
                        isNull = false;
                        for (int j = 0; j < rows[i].Length; j++)
                        {
                            one_row = one_row + rows[i][j] + " ";
                            if (rows[i][j].ToLower().Contains("null"))
                                isNull = true;
                        }
                        //insert data to stock_price
                        if (isNull == true)
                        {
                            //MessageBox.Show("row[" + i + "] contains null");
                            fail_cnt++;
                        }
                        else
                        {
                            success_cnt++;
                            updateMsg(String.Format("[{0}]{1}", i, one_row));
                            String sql = String.Format("insert into stock_price (stock_code,stock_date,compare_date,open,high,low,close) values ('{0}','{1}','{2}',{3},{4},{5},{6})", stock_code, rows[i][0], rows[i][0].Replace("-", ""), rows[i][1], rows[i][2], rows[i][3], rows[i][4]);
                            //updateMsg(sql);
                            String updateNum = DB_SQL(sql.ToString());

                            if (updateNum.All(Char.IsDigit))
                            {
                                if (Convert.ToInt32(updateNum) > 0)
                                    updateMsg("[" + rows[i][0] + "]寫入資料成功(" + updateNum.ToString() + ")");
                                else
                                    updateMsg("寫入資料失敗(" + updateNum.ToString() + ")");
                            }
                            else
                                updateMsg("[錯誤]" + updateNum);
                        }
                    }
                    updateMsg(String.Format("統計[{0}]匯入成功{1}筆 失敗{2}筆 總計{3}筆", stock_code, success_cnt, fail_cnt, rows.Length));
                }
                else
                    updateMsg(String.Format("錯誤! [{0}]代碼歷史資料不存在",stock_code));
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new Thread(() => {
                LoadStock(textBox1.Text);
            }).Start();
        }

        //舊版寫法 待刪除
        private void button6_Click(object sender, EventArgs e)
        {
            //new Thread(() =>
            //{
            //    String sql = "";

            //    //發放股利股息總共幾年
            //    int dividen_times = 0;
            //    //股利總和
            //    double dividen = 0.0;
            //    //取得所有代碼清單
            //    String[] stock_list = DB_SQL("select stock_code from stock_profile order by stock_code ", "stock_code").Split(',');
            //    //調整成 十年都有發股息的清單
            //    //String[] stock_list = DB_SQL("select stock_code from stock_dividen  where  stock_code>='3008' group by stock_code  order by stock_code", "stock_code").Split(',');

            //    //股利列表
            //    String[] dividen_list = null;

            //    //目前年度股利股息
            //    String[] year_dividen = null;
            //    //前一年度股利股息
            //    String[] next_dividen = null;
            //    //針對每隻股票去掃描

            //    //總交易日
            //    int total_days = 0;
            //    //有填權息日數
            //    int fill_days = 0;
            //    //有填權年數
            //    int win_times = 0;
            //    //勝率
            //    double win_rate = 0.0;
            //    double avg_win_rate = 0.0;
            //    //股票名稱
            //    String stock_name = "";
            //    String updateNum = "";
            //    //區間最大最小股價
            //    double min_price = 0.0;
            //    double max_price = 0.0;
            //    double max_range = 0.0;
            //    double min_range = 0.0;
            //    //最小填權日數
            //    String min_date = "";
            //    int least_fill_days = -1;

            //    for (int i = 0; i < stock_list.Length; i++)
            //    {
            //        sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}' and compare_date>20100101", stock_list[i]);
            //        dividen_times = Convert.ToInt32(DB_SQL(sql, "cnt"));
            //        //每次重算
            //        dividen = 0.0;
            //        win_rate = 0.0;
            //        win_times = 0;
            //        avg_win_rate = 0.0;

            //        stock_name = DB_SQL("select stock_name from stock_profile where stock_code='" + stock_list[i] + "'", "stock_name");
            //        //先過濾掉無資料的股票
            //        if (Convert.ToInt32(DB_SQL("select count(*) cnt from stock_price_new where stock_code='" + stock_list[i] + "'", "cnt")) == 0)
            //            updateMsgNew(String.Format("[{0}]{1} 無股價資料", stock_list[i], stock_name));
            //        else
            //        {
            //            for (int j = 0; j < dividen_times; j++)
            //            {
            //                sql = "select before_price,after_price,dividen,dividen_date,compare_date from stock_dividen where stock_code='" + stock_list[i] + "' order by dividen_date desc";
            //                dividen_list = DB_SQL(sql, "dividen_date,dividen,before_price,after_price,compare_date").Split(';');
            //                if (j == 0) //最新一年
            //                {
            //                    year_dividen = dividen_list[j].Split('@');
            //                    //updateMsgNew(String.Format("[{0}]{1}", stock_list[i], stock_name));
            //                    updateMsgNew(String.Format("[{0}~現在]股息{1}元 填息價{2}元 除權價{3}元", year_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));
            //                    //計算總交易日
            //                    sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date > {1} ", stock_list[i], year_dividen[4]);
            //                    total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
            //                    //計算有填權息日數
            //                    sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date > {1} and close > {2}", stock_list[i], year_dividen[4], year_dividen[2]);
            //                    fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
            //                    win_rate = (double)fill_days / total_days;

            //                    //取得區間最大、最小股價
            //                    sql = String.Format("select max(close) max_price,min(close) min_price from stock_price_new where stock_code='{0}' and compare_date > {1} ", stock_list[i], year_dividen[4]);
            //                    max_price = Convert.ToDouble(DB_SQL(sql, "max_price"));
            //                    sql = String.Format("select max(close) max_price,min(close) min_price from stock_price_new where stock_code='{0}' and compare_date > {1} ", stock_list[i], year_dividen[4]);
            //                    min_price = Convert.ToDouble(DB_SQL(sql, "min_price"));
            //                    //計算最大漲跌幅
            //                    max_range = max_price / Convert.ToDouble(year_dividen[2]);
            //                    min_range = min_price / Convert.ToDouble(year_dividen[2]);

            //                    //計算是否填權息
            //                    if (fill_days > 0)
            //                    {
            //                        win_times++;
            //                        sql = String.Format("select min(stock_date) min_date from stock_price_new where stock_code='{0}' and compare_date > {1} and close > {2}", stock_list[i], year_dividen[4], year_dividen[2]);
            //                        min_date = DB_SQL(sql, "min_date");
            //                        least_fill_days = (Convert.ToDateTime(min_date).Date - Convert.ToDateTime(year_dividen[0]).Date).Days;
            //                    }

            //                    updateMsgNew(String.Format("總交易天數{0}天 超過填息價總計{1}天 => 勝率 {2} % 最高股價{3}元(幅度{4}) 最低股價{5}元(幅度{6}){7}"
            //                        , total_days, fill_days, (win_rate * 100).ToString("F1"),max_price,max_range,min_price,min_range,(least_fill_days==-1)?"":"填息最少天數"+least_fill_days+"天"));

            //                    //計算股利合
            //                    dividen = dividen + Convert.ToDouble(year_dividen[1]);
            //                    //計算勝率
            //                    avg_win_rate = avg_win_rate + win_rate;

                                
            //                }
            //                else
            //                {
            //                    year_dividen = dividen_list[j].Split('@');
            //                    next_dividen = dividen_list[j - 1].Split('@');
            //                    updateMsgNew(String.Format("[{0}~{1}]股息{2}元 填息價{3}元 除權價{4}元", year_dividen[0], next_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));
            //                    //計算總交易日
            //                    sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_list[i], year_dividen[4], next_dividen[4]);
            //                    total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
            //                    //計算有填權息日數
            //                    sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} and close > {3}", stock_list[i], year_dividen[4], next_dividen[4], year_dividen[2]);
            //                    fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
            //                    //☆☆debug
            //                    if (total_days == 0)
            //                        win_rate = 0.0;
            //                    else
            //                        win_rate = (double)fill_days / total_days;

            //                    updateMsgNew(String.Format("總交易天數{0}天 超過填息價總計{1}天 => 勝率 {2} %{3}", total_days, fill_days, (win_rate * 100).ToString("F1"), Environment.NewLine));

            //                    //計算是否填權息
            //                    if (fill_days > 0)
            //                    {
            //                        win_times++;
            //                        sql = String.Format("select min(stock_date) min_date from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} and close > {3}", stock_list[i], year_dividen[4], next_dividen[4], year_dividen[2]);
            //                        min_date = DB_SQL(sql, "min_date");
            //                        least_fill_days = (Convert.ToDateTime(min_date).Date - Convert.ToDateTime(year_dividen[0]).Date).Days;
            //                    }
            //                    //計算股利合
            //                    dividen = dividen + Convert.ToDouble(year_dividen[1]);
            //                    //計算勝率
            //                    avg_win_rate = avg_win_rate + win_rate;

            //                }
            //            }

            //            //DB_SQL()
            //            //發大財勝率
            //            String big_money_rate = ((avg_win_rate / dividen_times) * 100).ToString("F1");
            //            updateMsgNew("=====================================================");
            //            updateMsgNew(String.Format("[{0}]{1}{6}十年發 {2}次股利(總合{3}元)填權息{4}次 ☆WADAR指標={5}%", stock_list[i], stock_name, dividen_times, dividen, win_times, big_money_rate, Environment.NewLine));
            //            updateMsgNew("=====================================================");
            //            //if(stock_list[i]=="1104")
            //            //    Thread.Sleep(200000);
            //            sql = String.Format("insert into money_rank_new (stock_code,stock_name,win_times,dividen_times,dividen,avg_win_rate,cal_date) values ('{0}','{1}',{2},{3},{4},{5},'{6}')", stock_list[i], stock_name, win_times, dividen_times, dividen, big_money_rate, DateTime.Now.ToString("yyyy/MM/dd"));

            //            updateNum = DB_SQL(sql.ToString());
            //            if (updateNum.All(Char.IsDigit))
            //            {
            //                if (Convert.ToInt32(updateNum) > 0)
            //                    updateMsgNew("寫入資料成功(" + updateNum.ToString() + ")");
            //                else
            //                    updateMsgNew("寫入資料失敗(" + updateNum.ToString() + ")");
            //            }
            //            else
            //                updateMsgNew("寫入資料失敗(" + updateNum + ")");
            //        }
            //    }
            //}).Start();
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            new Thread(() => {

                //取得所有代碼清單
                String[] stock_list = DB_SQL("select stock_code from stock_profile order by stock_code ", "stock_code").Split(',');

                for (int i = 0; i < stock_list.Length; i++)
                {
                    LoadStock(stock_list[i]);
                }
                
            }).Start();
            
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            new Thread(() =>
            {
                int finish_cnt = 0;
                int fail_count = 0;


                String stock_code = textBox1.Text;
                String saveFolder = String.Format(@"c://CSV//{0}//", stock_code);

                //updateMsg("clear");
                if (!Directory.Exists(saveFolder))
                {
                    //MessageBox.Show("尚未抓取過資料->建立資料夾 [" + stock_code + "]");
                    Directory.CreateDirectory(saveFolder);
                    updateMsg("未抓取過資料 建立資料夾[" + stock_code + "]成功");
                }

                //取得目前抓到的年份清單
                String[] files = Directory.GetFiles(saveFolder);
                try
                {
                    if (files.Length > 0)
                    {
                        updateMsg(String.Format("[{0}]檔案已下載完成", stock_code));
                        finish_cnt++;
                    }
                    else
                    {

                        updateMsg(String.Format("[{0}]...開始下載{1}", stock_code, Environment.NewLine));

                        //抓資料
                        DateTime a = DateTime.Now;
                        String crawl_date = a.ToString("yyyyMMdd");
                        //manual crawl old data
                        int year = Convert.ToInt32(crawl_date.Substring(0, 4));
                        int month = Convert.ToInt32(crawl_date.Substring(5, 2));
                        String file_name = stock_code + ".csv";

                        String url = "";
                        //updateMsg("抓取資料[" + stock_code + "]");
                        DateTime current = DateTime.Now;

                        //檢查file_name是否重覆
                        DateTime t1 = DateTime.Now;

                        //if (!files.Any(s => s.Contains(file_name)))
                        using (var client = new WebClient())
                        {
                            //get crumb & cookie
                            client.Headers.Add("user-agent", "Mozilla/5.0 (X11; U; Linux i686) Gecko/20071127 Firefox/2.0.0.11");
                            Stream data = client.OpenRead("https://finance.yahoo.com/quote/%5EGSPC");
                            string cookie = client.ResponseHeaders["Set-Cookie"];
                            StreamReader reader = new StreamReader(data);
                            String crumbPattern = "\"CrumbStore\":{\"crumb\":\"(?<crumb>[^\"]+)\"}";
                            String html = reader.ReadToEnd();
                            Match mCrumb = Regex.Match(html, crumbPattern);
                            string[] strs = mCrumb.Value.Split(':');
                            string currentCrumb = strs[2].Substring(1, strs[2].Length - 3);

                            client.Headers.Add("Cookie", cookie);
                            url = String.Format(@"https://query1.finance.yahoo.com/v7/finance/download/{0}.TW?period1=1230739200&period2=1569254400&interval=1d&events=history&crumb={1}", stock_code, currentCrumb);
                            //cral url            
                            //調整protocol
                            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                            client.DownloadFile(url, saveFolder + file_name);
                        }
                        DateTime t2 = DateTime.Now;
                        updateMsg(String.Format("{0}.csv下載完畢!費時{1}秒", stock_code, (t2 - t1).TotalSeconds));
                        //Thread.Sleep(Convert.ToInt32(textBox3.Text));
                    }
                }
                catch (Exception ee)
                {
                    updateMsg("[下載失敗]" + ee.Message);
                    fail_count++;
                }

                    
                updateMsg(String.Format("[{2}]成功下載{0}筆，失敗{1}筆", finish_cnt, fail_count,stock_code));
                
            }).Start();
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage2;
            
            if (!(Regex.IsMatch(textBox1.Text, @"[\u4E00-\u9FFF]\s[a-zA-Z0-9]") || Regex.IsMatch(textBox1.Text, @"[a-zA-Z0-9]\s[\u4E00 -\u9FFF]")))
            {
                MessageBox.Show("無此代碼", "輸入代碼錯誤"); 
            }
            else
            {
                label3.Visible = true;
                String stock_code = "";
                if (Regex.IsMatch(textBox1.Text.Split(' ')[0], @"[a-zA-Z0-9]"))
                    stock_code = textBox1.Text.Split(' ')[0];
                else
                    stock_code = textBox1.Text.Split(' ')[1];

                //取得最新交易日期
                String latest_trade_date = getLatestDate().ToString("yyyy/MM/dd");

                //從資料庫找出股價資料最新資料之日期
                String latest_date = getLatestDate().AddDays(-11).ToString("yyyy/MM/dd");//DB_SQL("select max(stock_date) latest_date from stock_price_new ", "latest_date");

                //just for test
                //latest_date = "2020/02/20";
                MessageBox.Show(String.Format("股價資料將自動爬取[{0}]最近10日交易資料[{1}~{2}]", textBox1.Text,latest_date, latest_trade_date));

                String sql = String.Empty;

                new Thread(() =>
                {
                    try
                    {
                        DateTime t1 = DateTime.Now;
                        try
                        {
                            double avg_high = 0.0;
                            double avg_low = 0.0;
                            double avg_amount = 0.0;
                            double now_price = 0.0;

                            dynamic rows = null;
                            DateTime sDateNew = Convert.ToDateTime(latest_date);

                            updateMsgNew(String.Format("即時爬取股票代碼「{0}」日期區間{1}~{2}", stock_code, latest_date, latest_trade_date));
                            //sDateNew = sDateNew.AddDays(-sub_days);
                            //☆需要修改的地方
                            String url = String.Format("https://www.cnyes.com/twstock/ps_historyprice.aspx?code={0}&ctl00$ContentPlaceHolder1$startText={1}&ctl00$ContentPlaceHolder1$endText={2}", stock_code, latest_date, latest_trade_date);
                            String xPath = "//*[@id=\"main3\"]/div[5]/div[3]/table";

                            DateTime t2 = DateTime.Now;
                            WebClient client = new WebClient();
                            MemoryStream ms = new MemoryStream(client.DownloadData(url));
                            // 使用utf8編碼讀入 HTML 才能正確顯示中文 
                            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                            doc.Load(ms, Encoding.UTF8);
                            // 裝載第一層查詢結果 
                            HtmlAgilityPack.HtmlDocument tableContent = new HtmlAgilityPack.HtmlDocument();
                            tableContent.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);
                            rows = tableContent.DocumentNode.SelectNodes("./tr");

                            int i = 0;

                            foreach (HtmlNode row in tableContent.DocumentNode.SelectNodes("./tr")) //☆需要修改的地方
                            {
                                if (i == 0)
                                {
                                    updateMsgNew(String.Format("[{0}]{1}", stock_code, stock_list[stock_code]));
                                }
                                else
                                {

                                    //☆需要修改的地方
                                    HtmlNodeCollection cells = row.SelectNodes("td");

                                    updateMsgNew(String.Format("日期[{0}] 開盤[{1}] 最高[{2}] 最低[{3}] 收盤[{4}] 成交量[{5}]", cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText, cells[4].InnerText, cells[8].InnerText));
                                    avg_high = avg_high + Convert.ToDouble(cells[2].InnerText);
                                    avg_low = avg_low +Convert.ToDouble(cells[3].InnerText);
                                    avg_amount = avg_amount +Convert.ToDouble(cells[8].InnerText.Replace(",",""));
                                    //最新價格
                                    now_price = Convert.ToDouble(cells[1].InnerText);
                                }
                                i++;
                            }
                            avg_high = Math.Round(avg_high / (now_price) ,1);
                            avg_low = Math.Round(avg_low / (now_price) ,1) ;
                            avg_amount = Math.Round(avg_amount / 10,1);
                            double now_percent = Math.Round(100*(Convert.ToDouble(textBox3.Text) - now_price) / now_price ,1);
                            String msg = String.Format(@"
===========================================================
                個人風險管家智慧分析
===========================================================

違約風險評估如下：

                   近[10]天平均 最高漲幅 {0}%  

        ☆ 輸入價格 ☆            漲跌幅 {2}%

                   近[10]天平均 最大跌幅 {1}%        

                           平均交易量 {3} 張 (『 高  』交易量)", avg_high,avg_low, now_percent,avg_amount);

                            TimeSpan ts1 = DateTime.Now - t2;
                            updateMsgNew(msg);
                            //updateMsgNew(String.Format("插入.更新{0}筆資料 費時{1}秒", rows.Count, ts1.TotalSeconds));
                            doc = null;
                            tableContent = null;
                            client = null;
                            ms.Close();
                        }
                        catch (Exception eee)
                        {
                            updateMsgNew(eee.Message);
                        }
                        //using (SQLiteConnection conn = new SQLiteConnection(connStr))
                        //{
                        //    conn.Open();
                        //    using (SQLiteTransaction trans = conn.BeginTransaction())
                        //    {
                        //        foreach (var item in stock_list)
                        //        {
                        //            num++;
                        //            //CrawlHtml(item.Key, latest_date, latest_trade_date);
                        //            String stock_code = item.Key;
                        //            try
                        //            {
                        //                dynamic rows = null;
                        //                DateTime sDateNew = Convert.ToDateTime(latest_date);

                        //                updateMsgNew(String.Format("更新股票代碼「{0}」日期區間{1}~{2}", stock_code, latest_date, latest_trade_date));
                        //                //sDateNew = sDateNew.AddDays(-sub_days);
                        //                //☆需要修改的地方
                        //                String url = String.Format("https://www.cnyes.com/twstock/ps_historyprice.aspx?code={0}&ctl00$ContentPlaceHolder1$startText={1}&ctl00$ContentPlaceHolder1$endText={2}", stock_code, latest_date, latest_trade_date);
                        //                String xPath = "//*[@id=\"main3\"]/div[5]/div[3]/table";

                        //                DateTime t2 = DateTime.Now;
                        //                WebClient client = new WebClient();
                        //                MemoryStream ms = new MemoryStream(client.DownloadData(url));
                        //                // 使用utf8編碼讀入 HTML 才能正確顯示中文 
                        //                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        //                doc.Load(ms, Encoding.UTF8);
                        //                // 裝載第一層查詢結果 
                        //                HtmlAgilityPack.HtmlDocument tableContent = new HtmlAgilityPack.HtmlDocument();
                        //                tableContent.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);
                        //                rows = tableContent.DocumentNode.SelectNodes("./tr");


                        //                int i = 0;


                        //                foreach (HtmlNode row in tableContent.DocumentNode.SelectNodes("./tr")) //☆需要修改的地方
                        //                {
                        //                    if (i == 0)
                        //                    {
                        //                        updateMsgNew(String.Format("[{0}]{1}", stock_code, stock_list[stock_code]));
                        //                    }
                        //                    else
                        //                    {

                        //                        //☆需要修改的地方
                        //                        HtmlNodeCollection cells = row.SelectNodes("td");

                        //                        updateMsgNew(String.Format("date[{0}] open[{1}] high[{2}] low[{3}] close[{4}] quantity[{5}]", cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText, cells[4].InnerText, cells[8].InnerText));

                        //                        //sql = sql + String.Format("INSERT OR replace into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change,quantity) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7},{8});", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText, cells[8].InnerText.Replace(",", ""));
                        //                        //no quantity
                        //                        sql = sql + String.Format("INSERT OR replace into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7});", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText);
                        //                    }
                        //                    i++;
                        //                }
                        //                String updateNum = "";
                        //                //☆ 主要修改 只使用一個連線
                        //                //測試 每10筆才寫入一次看看效果
                        //                if (num % 10 == 0)
                        //                {
                        //                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                        //                    {
                        //                        updateNum = cmd.ExecuteNonQuery().ToString();
                        //                    }

                        //                    if (updateNum.All(Char.IsDigit))
                        //                    {
                        //                        if (Convert.ToInt32(updateNum) > 0)
                        //                            updateMsgNew("寫入資料成功(" + updateNum.ToString() + ")");
                        //                        else
                        //                            updateMsgNew("寫入資料失敗(" + updateNum.ToString() + ")");
                        //                    }
                        //                    else
                        //                        if (updateNum.ToLower().Contains("unique"))
                        //                        updateMsg("資料重覆->更新資料");
                        //                    else
                        //                        updateMsg("[錯誤]" + updateNum);

                        //                    //清空sql字句 
                        //                    sql = String.Empty;
                        //                }

                        //                TimeSpan ts1 = DateTime.Now - t2;
                        //                updateMsgNew(String.Format("插入.更新{0}筆資料 費時{1}秒", rows.Count, ts1.TotalSeconds));
                        //                doc = null;
                        //                tableContent = null;
                        //                client = null;
                        //                ms.Close();
                        //            }
                        //            catch (Exception eee)
                        //            {
                        //                updateMsgNew(eee.Message);
                        //            }

                        //        }//end of foreach loop

                        //        trans.Commit();
                        //    }
                        //    conn.Close();
                        //}
                        TimeSpan ts = DateTime.Now - t1;

                        //updateMsgNew(String.Format("插入更新{0}檔股票 費時{1}秒", stock_list.Count, ts.TotalSeconds));

                    }
                    catch (Exception ee)
                    {
                        updateMsgNew(ee.Message);
                    }
                }).Start();
                //new Thread(() =>
                //{
                //    String sql = "";

                //    String stock_code = "";
                //    if (Regex.IsMatch(textBox1.Text.Split(' ')[0], @"[a-zA-Z0-9]"))
                //        stock_code = textBox1.Text.Split(' ')[0];
                //    else
                //        stock_code = textBox1.Text.Split(' ')[1];

                //    //發放股利股息總共幾年
                //    int dividen_times = 0;
                //    //股利總和
                //    double dividen = 0.0;

                //    //股利列表
                //    String[] dividen_list = null;

                //    //目前年度股利股息
                //    String[] year_dividen = null;
                //    //前一年度股利股息
                //    String[] next_dividen = null;
                //    //針對每隻股票去掃描

                //    //總交易日
                //    int total_days = 0;
                //    //有填權息日數
                //    int fill_days = 0;
                //    //有填權年數
                //    int win_times = 0;
                //    //勝率
                //    double win_rate = 0.0;
                //    double avg_win_rate = 0.0;
                //    //股票名稱
                //    String stock_name = "";

                //    //區間最大最小股價
                //    double min_price = 0.0;
                //    double max_price = 0.0;
                //    double max_range = 0.0;
                //    double min_range = 0.0;

                //    double increase_max = 0.0;  //歷史填權最大漲幅
                //    double increase_min = 0.0;  //歷史填權最小漲幅
                //    double decrease_max = 0.0;  //歷史最大跌幅
                //    double decrease_min = 0.0;  //歷史最小跌幅

                //    //最小填權日數
                //    String min_date = "";
                //    int least_fill_days = -1;   //日曆天
                //    int least_trade_days = -1;  //交易天

                //    //計算總權重數
                //    double total_weight = 0.0;

                //    //最新一年 填息價
                //    double latest_fill_price = 0.0;
                //    //推估最新一年買入價格 & 停利價格
                //    double min_buy_price = 0.0;
                //    double min_sell_price = 0.0;

                //    //紀錄trackBar值
                //    int trackBarValue = 0;
                //    this.Invoke(new MethodInvoker(delegate { trackBarValue = trackBar1.Value; }));

                //    //計算總花費時間
                //    DateTime t1 = DateTime.Now;

                //    //最小填息天數平均
                //    double avg_fill_days = 0.0;

                //    //ai迴歸預測最低價格
                //    double ai_predict_price = 0.0;
                //    double ai_predict_range = 0.0;

                //    DateTime t2 = DateTime.Now;
                //    //sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}' and compare_date >={1}0101", stock_code, DateTime.Now.AddYears(-trackBarValue).ToString("yyyy"));
                //    sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}' ", stock_code, DateTime.Now.AddYears(-trackBarValue).ToString("yyyy"));

                //    dividen_times = Convert.ToInt32(DB_SQL(sql, "cnt"));
                //    //每次重算
                //    dividen = 0.0;
                //    win_rate = 0.0;
                //    win_times = 0;
                //    avg_win_rate = 0.0;
                //    try
                //    {
                //            //String stock_code = stock_list[i];
                //            stock_name = DB_SQL("select stock_name from stock_profile where stock_code='" + stock_code + "'", "stock_name");
                //            //先過濾掉無資料的股票
                //            if (Convert.ToInt32(DB_SQL("select count(*) cnt from stock_price_new where stock_code='" + stock_code + "'", "cnt")) == 0)
                //                updateMsg(String.Format("[{0}]{1} 無股價資料", stock_code, stock_name));
                //            else
                //            {
                //                updateMsg(String.Format("==============================================================={0}[{1}]{2}", Environment.NewLine, stock_code, stock_name));
                //                if (dividen_times > 0)
                //                {
                //                    for (int j = 0; j < dividen_times; j++)
                //                    {

                //                        sql = "select before_price,after_price,dividen,dividen_date,compare_date,least_days from stock_dividen where stock_code='" + stock_code + "' order by dividen_date desc";
                //                        dividen_list = DB_SQL(sql, "dividen_date,dividen,before_price,after_price,compare_date,least_days").Split(';');

                //                        if (j == 0) //最新一年
                //                        {
                //                            year_dividen = dividen_list[j].Split('@');
                //                            //updateMsg(String.Format("[{0}]{1}", stock_list[i], stock_name));
                //                            updateMsg(String.Format("[{0}~現在]股息{1}元 填息價{2}元 除權價{3}元", year_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));

                //                            //紀錄最新一年填息價
                //                            latest_fill_price = Convert.ToDouble(year_dividen[2]);

                //                            //計算總交易日
                //                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date > {1} ", stock_code, year_dividen[4]);
                //                            total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                //                            //計算有填權息日數
                //                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date > {1} and close > {2}", stock_code, year_dividen[4], year_dividen[2]);
                //                            fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                //                            win_rate = (double)fill_days / total_days;

                //                            //取得區間最大、最小股價
                //                            sql = String.Format("select max(close) max_price,min(close) min_price from stock_price_new where stock_code='{0}' and compare_date > {1} ", stock_code, year_dividen[4]);
                //                            max_price = Convert.ToDouble(DB_SQL(sql, "max_price"));
                //                            min_price = Convert.ToDouble(DB_SQL(sql, "min_price"));
                //                            //計算最大漲跌幅
                //                            max_range = (Math.Round((max_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                //                            min_range = (Math.Round((min_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                //                            increase_max = max_range;
                //                            increase_min = max_range;
                //                            decrease_max = min_range;
                //                            decrease_min = min_range;

                //                            //計算是否填權息
                //                            if (fill_days > 0)
                //                            {
                //                                win_times++;
                //                                sql = String.Format("select min(stock_date) min_date from stock_price_new where stock_code='{0}' and compare_date >= {1} and close >= {2}", stock_code, year_dividen[4], year_dividen[2]);
                //                                min_date = DB_SQL(sql, "min_date");
                //                                least_fill_days = (Convert.ToDateTime(min_date).Date - Convert.ToDateTime(year_dividen[0]).Date).Days;
                //                                sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} ", stock_code, year_dividen[4], min_date.Replace("/", ""));

                //                                //updateMsg(sql);

                //                                least_trade_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                //                                avg_fill_days = least_trade_days;
                //                            }
                //                            //updateMsg(String.Format("最高股價{3}元(填息後最大漲幅{4}%) 最低股價{5}元(除息後最大跌幅{6}%)"
                //                            //    , total_days, fill_days, (win_rate * 100).ToString("F1"), max_price, max_range, min_price, min_range, (fill_days > 0) ? " 填息最少交易" + least_trade_days + "天(日曆" + least_fill_days + "天)" : " 沒填息", Environment.NewLine, year_dividen[5], year_dividen[4], min_date));

                //                        updateMsg(String.Format("總交易天數{0}天 超過填息價總計{1}天{8}=> 勝率 {2} % 最高股價{3}元(填息後最大漲幅{4}%) 最低股價{5}元(除息後最大跌幅{6}%){8}    程式計算{7} [{10}~{11}]{8}vs GoodInfo計算最小天數{9}天{8}"
                //                           , total_days, fill_days, (win_rate * 100).ToString("F1"), max_price, max_range, min_price, min_range, (fill_days > 0) ? " 填息最少交易" + least_trade_days + "天(日曆" + least_fill_days + "天)" : " 沒填息", Environment.NewLine, year_dividen[5], year_dividen[4], min_date));

                //                        //計算股利合
                //                        dividen = dividen + Convert.ToDouble(year_dividen[1]);

                //                        //最新一年權重=1
                //                        total_weight = 1;
                //                        //計算勝率 最新一年權重 =  1
                //                        avg_win_rate = avg_win_rate + win_rate * 1;

                //                        //2020.6.8  AI預測
                //                        updateMsg("利用載入資料建立「預測模型引擎」");

                //                        //利用載入資料建立「預測模型引擎」
                //                        MLContext mlContext = new MLContext(seed: 0);
                //                        updateMsg("載入特徵值並轉換資料中…");
                //                        //updateMsg("載入中…");
                //                        updateMsg("載入特徵[stock_code]");
                //                        updateMsg("載入特徵[dividen]");
                //                        updateMsg("載入特徵[last_price]");
                //                        updateMsg("載入特徵[win_rate]");
                //                        updateMsg("載入特徵[fill_dividen]");
                //                        updateMsg("載入特徵[max_rise]");
                //                        updateMsg("載入特徵[max_drop]");
                //                        //updateMsg("開始載入訓練資料->迴歸模型");
                //                        var model = Train(mlContext, _trainDataPath);
                //                        var predict_test = mlContext.Model.CreatePredictionEngine<MaxRiseDrop, MaxRiseDropPrdeiction>(model);


                //                        updateMsg(String.Format("最新區間 代碼[{0}] 股利[{1}] 除息前價格[{2}] 單一區間WADAR值[{3}] 是否填息[{4}] 最大漲幅[{5}%] 最大跌幅[{6}%]",
                //                                                stock_code, year_dividen[1], year_dividen[2], win_rate,(fill_days > 0) ? "是" : "否", max_range,min_range));
                //                            //var testSample = new MaxRiseDrop()
                //                            //{
                //                            //    stock_code = "2330",
                //                            //    dividen = 0.7F,
                //                            //    last_price = 92.15F,
                //                            //    win_rate = 0.81F,
                //                            //    fill_dividen = 1F,
                //                            //    max_rise = 19.2F
                //                            //    ,
                //                            //    max_drop = 0    //用來預測真實的數值
                //                            //};

                //                            var testSample = new MaxRiseDrop()
                //                            {
                //                                stock_code = stock_code,
                //                                dividen = Convert.ToSingle(year_dividen[1]),
                //                                last_price = Convert.ToSingle(year_dividen[2]),
                //                                win_rate = Convert.ToSingle(win_rate),
                //                                fill_dividen = Convert.ToSingle((fill_days > 0)?"1":"0"),
                //                                max_rise = Convert.ToSingle(max_range),
                //                                //用來預測真實的數值
                //                                max_drop = 0    
                //                            };
                //                            //單次測試
                //                            updateMsg("開始單次預測結果");
                //                            var prediction = predict_test.Predict(testSample);
                //                            //double true_value = Convert.ToDouble(one_row[6]);
                //                            ai_predict_price = Math.Round(Convert.ToDouble(year_dividen[2]) * (100 - prediction.max_drop) / 100, 2);
                //                            updateMsg(String.Format("最新填息價{0}元 預測下跌:{1}% AI預測可買進價格{2}元  {3}{3}", year_dividen[2], prediction.max_drop, ai_predict_price,Environment.NewLine));

                //                            ai_predict_range = Math.Round(prediction.max_drop,1);
                //                    }
                //                    else
                //                    {
                //                            year_dividen = dividen_list[j].Split('@');
                //                            next_dividen = dividen_list[j - 1].Split('@');
                //                            updateMsg(String.Format("[{0}~{1}]股息{2}元 填息價{3}元 除權價{4}元", year_dividen[0], next_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));
                //                            //計算總交易日
                //                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], next_dividen[4]);
                //                            total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                //                            //計算有填權息日數
                //                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} and close >= {3}", stock_code, year_dividen[4], next_dividen[4], year_dividen[2]);
                //                            fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                //                            //☆☆debug
                //                            if (total_days == 0)
                //                                win_rate = 0.0;
                //                            else
                //                                win_rate = (double)fill_days / total_days;

                //                            //取得區間最大、最小股價
                //                            sql = String.Format("select max(close) max_price,min(close) min_price from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], next_dividen[4]);
                //                            max_price = Convert.ToDouble(DB_SQL(sql, "max_price"));
                //                            min_price = Convert.ToDouble(DB_SQL(sql, "min_price"));
                //                            //計算最大漲跌幅
                //                            max_range = (Math.Round((max_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                //                            min_range = (Math.Round((min_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                //                            if (max_range > increase_max) increase_max = max_range;
                //                            if (max_range < increase_min) increase_min = max_range;
                //                            if (min_range > decrease_min) decrease_min = min_range;
                //                            if (min_range < decrease_max) decrease_max = min_range;

                //                            //計算是否填權息
                //                            if (fill_days > 0)
                //                            {
                //                                win_times++;
                //                                sql = String.Format("select min(stock_date) min_date from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} and close >= {3}", stock_code, year_dividen[4], next_dividen[4], year_dividen[2]);
                //                                min_date = DB_SQL(sql, "min_date");
                //                                least_fill_days = (Convert.ToDateTime(min_date).Date - Convert.ToDateTime(year_dividen[0]).Date).Days;
                //                                sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], min_date.Replace("/", ""), year_dividen[2]);
                //                                //updateMsg(sql);
                //                                least_trade_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                //                                avg_fill_days = avg_fill_days + least_trade_days;
                //                            }
                //                            updateMsg(String.Format("總交易天數{0}天 超過填息價總計{1}天{8}=> 勝率 {2} % 最高股價{3}元(填息後最大漲幅{4}%) 最低股價{5}元(除息後最大跌幅{6}%){8}    程式計算{7} [{10}~{11}]{8}vs GoodInfo計算最小天數{9}天{8}"
                //                                 , total_days, fill_days, (win_rate * 100).ToString("F1"), max_price, max_range, min_price, min_range, (fill_days > 0) ? " 填息最少交易" + least_trade_days + "天(日曆" + least_fill_days + "天)" : " 沒填息", Environment.NewLine, year_dividen[5], year_dividen[4], min_date));

                //                            //updateMsg(String.Format("最高股價{3}元(填息後最大漲幅{4}%) 最低股價{5}元(除息後最大跌幅{6}%)"
                //                            //        , total_days, fill_days, (win_rate * 100).ToString("F1"), max_price, max_range, min_price, min_range, (fill_days > 0) ? " 填息最少交易" + least_trade_days + "天(日曆" + least_fill_days + "天)" : " 沒填息", Environment.NewLine, year_dividen[5], year_dividen[4], min_date));


                //                            //計算股利合
                //                            dividen = dividen + Convert.ToDouble(year_dividen[1]);
                //                            //每年的權重 依序為  0.9 ~ 0.8 ... 0.1
                //                            double year_weight = 1 - 0.1 * j;
                //                            total_weight = total_weight + year_weight;
                //                            //計算勝率 採用加權比重平均
                //                            avg_win_rate = avg_win_rate + win_rate * (year_weight);

                //                        }
                //                    }

                //                }

                //            }
                //            //計算 哇達 勝率
                //            String WADAR_win_rate = ((avg_win_rate / total_weight) * 100).ToString("F1");
                //            //計算 單一筆 花費時間
                //            TimeSpan ts2 = DateTime.Now - t2;
                //            min_buy_price = Math.Round(latest_fill_price * (100 + decrease_min) / 100, 2);
                //            min_sell_price = Math.Round(latest_fill_price * (100 + increase_min) / 100, 2);

                //            avg_fill_days = Math.Round(avg_fill_days /win_times, 2);
                //            updateMsg(String.Format("[{0}]{1}{6}十年發 {2}次股利(總合{3}元)填權息{4}次 ☆WADAR指標={5}% 最小平均填息天數{11}天 {6}填息後最大漲幅介於{7}%~{8}% 除息後最大跌幅介於{9}%~{10}%",
                //                      stock_code, stock_name, dividen_times, dividen, win_times, WADAR_win_rate, Environment.NewLine, increase_min, increase_max, decrease_min, decrease_max, avg_fill_days));
                //        //updateMsg(String.Format("★ 最新填息價 {0} 元,推估至少低於{1}元({2}%)買入(買進點) 高於至少{3}元({4}%)賣出(停利點)", latest_fill_price, min_buy_price, decrease_min, min_sell_price, increase_min));
                //        updateMsg(String.Format("★ 大數據推測 最小獲利價格至少低於{1}元({2}%)買入 最新填息價 {0} 元 ", latest_fill_price, min_buy_price, decrease_min, min_sell_price, increase_min));

                //        //double ai_price = Math.Round(latest_fill_price * (100 + decrease_min) / 200, 2);
                //        updateMsg(String.Format("★ AI迴歸預測 波段最低價格至少低於{0}元(跌{1}%)", ai_predict_price, ai_predict_range));
                //        //updateMsg(String.Format("花費時間{0}秒", Math.Round(ts2.TotalSeconds,2)));
                //        updateMsg("===============================================================");
                //        }
                //        catch (Exception eex)
                //        {
                //            updateMsg(eex.Message);
                //        }

                //}).Start();
            }
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            new Thread(() =>
            {
                updateMsg("clear");
                String sql = "";

                //發放股利股息總共幾年
                int dividen_times = 0;
                //股利總和
                double dividen = 0.0;
                //取得所有代碼清單
                //String[] stock_list = DB_SQL("select stock_code from stock_profile order by stock_code ", "stock_code").Split(',');
                //調整成 十年都有發股息的清單
                String[] db_data = DB_SQL("select stock_code,stock_name from stock_dividen group by stock_code having COUNT(*) >=10 order by stock_code", "stock_code,stock_name").Split(';');

                //var rows = File.ReadAllLines(path).Select(line => line.Split(',').ToArray()).ToArray();
                var stock_list = db_data.Select(line => line.ToString().Split('@').ToArray()).ToArray();
                
                //股利列表
                String[] dividen_list = null;

                //目前年度股利股息
                String[] year_dividen = null;
                //前一年度股利股息
                String[] next_dividen = null;
                //針對每隻股票去掃描

                //總交易日
                int total_days = 0;
                //有填權息日數
                int fill_days = 0;
                //有填權年數
                int win_times = 0;
                //勝率
                double win_rate = 0.0;
                double avg_win_rate = 0.0;
                //股票名稱
                String stock_name = "";
                String updateNum = "";

                //除權前價格
                double before_price = 0.0;

                //檢查是否在列表裡
                int isInclude = 0;
                updateMsgNew("十年皆有發股利名單");
                for (int i = 0; i < stock_list.Length-1; i++)
                {
                    updateMsgNew(String.Format("({0})[{1}]{2}",(i+1),stock_list[i][0],stock_list[i][1]));
                    if (textBox1.Text == stock_list[i][0])
                    {
                        isInclude = 1;
                        stock_name = stock_list[i][1];
                    }
                }
                if (isInclude == 0)
                    MessageBox.Show(textBox1.Text + " 不在名單內 無資料");
                else
                {
                    String stock_code = textBox1.Text;

                    //for (int i = 0; i < stock_list.Length; i++)
                    //{
                        sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}' order by dividen_date desc", stock_code);
                        dividen_times = Convert.ToInt32(DB_SQL(sql, "cnt"));
                        //每次重算
                        dividen = 0.0;
                        win_rate = 0.0;
                        win_times = 0;
                        avg_win_rate = 0.0;

                        //stock_name = DB_SQL("select stock_name from stock_profile where stock_code='" + stock_list[i] + "'", "stock_name");
                        //先過濾掉無資料的股票
                        if (Convert.ToInt32(DB_SQL("select count(*) cnt from stock_price where stock_code='" + stock_code + "'", "cnt")) == 0)
                            updateMsg(String.Format("[{0}]{1} 無股價資料", stock_code, stock_name));
                        else
                        {
                            for (int j = 0; j < dividen_times; j++)
                            {
                                sql = "select before_price,after_price,dividen,dividen_date,compare_date from stock_dividen where stock_code='" + stock_code + "' order by dividen_date desc";
                                dividen_list = DB_SQL(sql, "dividen_date,dividen,before_price,after_price,compare_date").Split(';');
                                if (j == 0) //最新一年
                                {
                                    year_dividen = dividen_list[j].Split('@');
                                    //updateMsg(String.Format("[{0}]{1}", stock_list[i], stock_name));
                                    updateMsg(String.Format("[{0}~現在]股息{1}元 填息價{2}元 除權價{3}元", year_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));
                                    //計算總交易日
                                    sql = String.Format("select count(*) cnt from stock_price where stock_code='{0}' and compare_date > {1} ", stock_code, year_dividen[4]);
                                    total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                    //計算有填權息日數
                                    sql = String.Format("select count(*) cnt from stock_price where stock_code='{0}' and compare_date > {1} and close > {2}", stock_code, year_dividen[4], year_dividen[2]);
                                    fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                    win_rate = (double)fill_days / total_days;

                                    //====================================================================================================================
                                    //2019.10.5 檢核資料細項
                                    sql = String.Format("select stock_date,close from stock_price where stock_code='{0}' and compare_date > {1} ", stock_code, year_dividen[4]);
                                    String[] detail_row = DB_SQL(sql, "stock_date,close").Split(';');
                                    var details = detail_row.Select(line => line.Split('@').ToArray()).ToArray();

                                    before_price = Convert.ToDouble(year_dividen[2]);
                                    double day_price = 0.0;
                                    int win_days = 0;
                                    for (int k = 0; k < details.Length-1; k++)
                                    {
                                        day_price = Convert.ToDouble(details[k][1]);
                                        if (before_price <= day_price)
                                        {
                                            win_days++;
                                            updateMsg(String.Format("[{0}]收盤價:{1} 高於填權價 {2} (累積{3}天)",details[k][0],details[k][1],before_price,win_days));
                                        }
                                        else
                                        {
                                            updateMsg(String.Format("[{0}]收盤價:{1} 低於填權價 {2} ", details[k][0], details[k][1], before_price));
                                        }
                                    }
                                    //====================================================================================================================

                                    updateMsg(String.Format("總交易天數{0}天 超過填息價總計{1}天 => 勝率 {2} %", total_days, fill_days, (win_rate * 100).ToString("F1")));

                                    //計算是否填權息
                                    if (fill_days > 0)
                                        win_times++;
                                    //計算股利合
                                    dividen = dividen + Convert.ToDouble(year_dividen[1]);
                                    //計算勝率
                                    avg_win_rate = avg_win_rate + win_rate;

                                    //sql = String.Format("insert into fa_big_money (stock_code,stock_name,fill_days,total_days,win_rate,fill_dividen,start_date,end_date,dividen) values ('{0}','{1}',{2},{3},{4},{5},'{6}','{7}',{8})",
                                    //     stock_code, stock_name, fill_days, total_days, win_rate, (fill_days > 0) ? "1" : "0", year_dividen[0], DateTime.Now.ToString("yyyy/MM/dd"), dividen);

                                    //updateNum = DB_SQL(sql.ToString());
                                    //if (updateNum.All(Char.IsDigit))
                                    //{
                                    //    if (Convert.ToInt32(updateNum) > 0)
                                    //        updateMsg("寫入資料成功(" + updateNum.ToString() + ")");
                                    //    else
                                    //        updateMsg("寫入資料失敗(" + updateNum.ToString() + ")");
                                    //}
                                    //else
                                    //    updateMsg("寫入資料失敗(" + updateNum + ")");
                                }
                                else
                                {
                                    year_dividen = dividen_list[j].Split('@');
                                    next_dividen = dividen_list[j - 1].Split('@');
                                    updateMsg(String.Format("[{0}~{1}]股息{2}元 填息價{3}元 除權價{4}元", year_dividen[0], next_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));
                                    //計算總交易日
                                    sql = String.Format("select count(*) cnt from stock_price where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], next_dividen[4]);
                                    total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                    //計算有填權息日數
                                    sql = String.Format("select count(*) cnt from stock_price where stock_code='{0}' and compare_date between {1} and {2} and close > {3}", stock_code, year_dividen[4], next_dividen[4], year_dividen[2]);
                                    fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                    //☆☆debug
                                    if (total_days == 0)
                                        win_rate = 0.0;
                                    else
                                        win_rate = (double)fill_days / total_days;

                                    //====================================================================================================================
                                    //2019.10.5 檢核資料細項
                                    sql = String.Format("select stock_date,close from stock_price where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], next_dividen[4]);
                                    String[] detail_row = DB_SQL(sql, "stock_date,close").Split(';');
                                    var details = detail_row.Select(line => line.Split('@').ToArray()).ToArray();

                                    before_price = Convert.ToDouble(year_dividen[2]);
                                    double day_price = 0.0;
                                    int win_days = 0;
                                    for (int k = 0; k < details.Length-1; k++)
                                    {
                                        day_price = Convert.ToDouble(details[k][1]);
                                        if (before_price <= day_price)
                                        {
                                            win_days++;
                                            updateMsg(String.Format("[{0}]收盤價:{1} 高於填權價 {2} (累積{3}天)", details[k][0], details[k][1], before_price, win_days));
                                        }
                                        else
                                        {
                                            updateMsg(String.Format("[{0}]收盤價:{1} 低於填權價 {2} ", details[k][0], details[k][1], before_price));
                                        }
                                    }
                                    //====================================================================================================================


                                updateMsg(String.Format("總交易天數{0}天 超過填息價總計{1}天 => 勝率 {2} %{3}", total_days, fill_days, (win_rate * 100).ToString("F1"), Environment.NewLine));

                                    //計算是否填權息
                                    if (fill_days > 0)
                                        win_times++;
                                    //計算股利合
                                    dividen = dividen + Convert.ToDouble(year_dividen[1]);
                                    //計算勝率
                                    avg_win_rate = avg_win_rate + win_rate;
                                    //Thread.Sleep(10000);
                                    //sql = String.Format("insert into fa_big_money (stock_code,stock_name,fill_days,total_days,win_rate,fill_dividen,start_date,end_date,dividen) values ('{0}','{1}',{2},{3},{4},{5},'{6}','{7}',{8})",
                                    //                                               stock_code, stock_name, fill_days, total_days, win_rate, (fill_days > 0) ? "1" : "0", year_dividen[0], next_dividen[0], dividen);
                                    //updateNum = DB_SQL(sql.ToString());
                                    //if (updateNum.All(Char.IsDigit))
                                    //{
                                    //    if (Convert.ToInt32(updateNum) > 0)
                                    //        updateMsg("寫入資料成功(" + updateNum.ToString() + ")");
                                    //    else
                                    //        updateMsg("寫入資料失敗(" + updateNum.ToString() + ")");
                                    //}
                                    //else
                                    //    updateMsg("寫入資料失敗(" + updateNum + ")");
                                }
                            }

                            //發大財勝率WADAR指標
                            String big_money_rate = ((avg_win_rate / dividen_times) * 100).ToString("F1");
                            updateMsg("=====================================================");
                            updateMsg(String.Format("[{0}]{1}{6}十年發 {2}次股利(總合{3}元)填權息{4}次 ☆WADAR指標={5}%", stock_code, stock_name, dividen_times, dividen, win_times, big_money_rate, Environment.NewLine));
                            updateMsg("=====================================================");

                            sql = String.Format("insert into money_rank (stock_code,stock_name,win_times,dividen_times,dividen,avg_win_rate,cal_date) values ('{0}','{1}',{2},{3},{4},{5},'{6}')", stock_code, stock_name, win_times, dividen_times, dividen, big_money_rate, DateTime.Now.ToString("yyyy/MM/dd"));

                            updateNum = DB_SQL(sql.ToString());
                            if (updateNum.All(Char.IsDigit))
                            {
                                if (Convert.ToInt32(updateNum) > 0)
                                    updateMsg("寫入資料成功(" + updateNum.ToString() + ")");
                                else
                                    updateMsg("寫入資料失敗(" + updateNum.ToString() + ")");
                            }
                            else
                                updateMsg("寫入資料失敗(" + updateNum + ")");
                        }
                    //}
                
                }
            }).Start();
        }
        //爬取資料->爬取歷史股價
        private void 爬取歷史股價ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage1;
            label3.Visible = false;
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;
            //取得最新日期
            String now_date = DateTime.Now.ToString("yyyy/MM/dd");
            String sql = String.Empty;

            DialogResult result = MessageBox.Show("此功能將花費極大量時間(超過6小時)是否確定要下載?","大量下載警告",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                MessageBox.Show("選擇取消下載", "取消下載");
            }
            else
            {
                //updateMsgNew(String.Format("更新股票代碼「{0}」日期區間{1}~{2}", "0050", "2008/01/01", now_date));
                //CrawlHtml("0050", "2008/01/01", now_date);
                String stock_code = "";
                String latest_date = "2008/01/01";
                int num = 0;
                String updateNum = "";
                new Thread(() =>
                {
                    try
                    {
                        DateTime t1 = DateTime.Now;

                        using (SQLiteConnection conn = new SQLiteConnection(connStr))
                        {
                            conn.Open();
                            using (SQLiteTransaction trans = conn.BeginTransaction())
                            {
                                foreach (var item in stock_list)
                                {
                                    num++;
                                    //CrawlHtml(item.Key, latest_date, latest_trade_date);
                                    stock_code = item.Key;

                                    if(stock_code.Length ==4)
                                    if(Convert.ToInt32(stock_code)<=2630)
                                    try
                                    {
                                        dynamic rows = null;
                                        DateTime sDateNew = Convert.ToDateTime(latest_date);

                                        updateMsgNew(String.Format("更新股票代碼「{0}」日期區間{1}~{2}", stock_code, latest_date, now_date));
                                        //sDateNew = sDateNew.AddDays(-sub_days);
                                        //☆需要修改的地方
                                        String url = String.Format("https://www.cnyes.com/twstock/ps_historyprice.aspx?code={0}&ctl00$ContentPlaceHolder1$startText={1}&ctl00$ContentPlaceHolder1$endText={2}", stock_code, latest_date, now_date);
                                        String xPath = "//*[@id=\"main3\"]/div[5]/div[3]/table";

                                        DateTime t2 = DateTime.Now;
                                        WebClient client = new WebClient();
                                        MemoryStream ms = new MemoryStream(client.DownloadData(url));
                                        // 使用utf8編碼讀入 HTML 才能正確顯示中文 
                                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                        doc.Load(ms, Encoding.UTF8);
                                        // 裝載第一層查詢結果 
                                        HtmlAgilityPack.HtmlDocument tableContent = new HtmlAgilityPack.HtmlDocument();
                                        tableContent.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);
                                        rows = tableContent.DocumentNode.SelectNodes("./tr");


                                        int i = 0;


                                        foreach (HtmlNode row in tableContent.DocumentNode.SelectNodes("./tr")) //☆需要修改的地方
                                        {
                                            if (i == 0)
                                            {
                                                updateMsgNew(String.Format("[{0}]{1}", stock_code, stock_list[stock_code]));
                                            }
                                            else
                                            {

                                                //☆需要修改的地方
                                                HtmlNodeCollection cells = row.SelectNodes("td");

                                                updateMsgNew(String.Format("date[{0}] open[{1}] high[{2}] low[{3}] close[{4}] quantity[{5}]", cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText, cells[4].InnerText, cells[8].InnerText));
                                                        
                                                sql = sql + String.Format("INSERT OR replace into stock_price_test (stock_code,stock_date,compare_date,open,high,low,close,price_change,quantity) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7},{8});", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText, cells[8].InnerText.Replace(",", ""));
                                                        //updateMsgNew(sql);
                                                        //using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                                                        //{
                                                        //    updateNum = cmd.ExecuteNonQuery().ToString();
                                                        //}

                                                        //if (updateNum.All(Char.IsDigit))
                                                        //{
                                                        //    if (Convert.ToInt32(updateNum) > 0)
                                                        //        updateMsgNew("寫入資料成功(" + updateNum.ToString() + ")");
                                                        //    else
                                                        //        updateMsgNew("寫入資料失敗(" + updateNum.ToString() + ")");
                                                        //}
                                                        //else
                                                        //    if (updateNum.ToLower().Contains("unique"))
                                                        //    updateMsg("資料重覆->更新資料");
                                                        //else
                                                        //    updateMsg("[錯誤]" + updateNum);

                                                        ////清空sql字句 
                                                        //sql = String.Empty;
                                                        //Thread.Sleep(100000);
                                                        //no quantity
                                                        //sql = sql + String.Format("INSERT OR replace into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7});", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText);
                                                    }
                                            i++;
                                        }

                                                //☆ 主要修改 只使用一個連線
                                                //測試 每10筆才寫入一次看看效果
                                                //if (num % 10 == 0)
                                                //{
                                                using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                                                {
                                                    updateNum = cmd.ExecuteNonQuery().ToString();
                                                }

                                                if (updateNum.All(Char.IsDigit))
                                                {
                                                    if (Convert.ToInt32(updateNum) > 0)
                                                        updateMsgNew("寫入資料成功(" + updateNum.ToString() + ")");
                                                    else
                                                        updateMsgNew("寫入資料失敗(" + updateNum.ToString() + ")");
                                                }
                                                else
                                                    if (updateNum.ToLower().Contains("unique"))
                                                    updateMsg("資料重覆->更新資料");
                                                else
                                                    updateMsg("[錯誤]" + updateNum);

                                                //清空sql字句 
                                                sql = String.Empty;
                                                //}

                                                TimeSpan ts1 = DateTime.Now - t2;
                                        updateMsgNew(String.Format("API下載{0}筆資料 費時{1}秒", rows.Count, ts1.TotalSeconds));
                                        doc = null;
                                        tableContent = null;
                                        client = null;
                                        ms.Close();
                                    }
                                    catch (Exception eee)
                                    {
                                        updateMsgNew(eee.Message);
                                    }

                                }//end of foreach loop

                                trans.Commit();
                            }
                            conn.Close();
                        }
                        TimeSpan ts = DateTime.Now - t1;
                        updateMsgNew(String.Format("插入更新{0}檔股票 費時{1}秒", stock_list.Count, ts.TotalSeconds));

                    }
                    catch (Exception ee)
                    {
                        updateMsgNew(ee.Message);
                    }
                }).Start();
            }
        }

        //爬取資料->爬取歷史股利
        private void 爬取歷史股利ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage1;
            label3.Visible = false;
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            MessageBox.Show(String.Format("來抓除{0}權息價格", textBox2.Text), "開始發財$");

            //note：編碼 顯示中文
            //var rows = File.ReadAllLines("2018.csv.txt").Select(l => l.Split(',').ToArray()).ToArray();
            new Thread(() => {
                int idx = 0;        //往前剖析10年
                //for (idx = 1; idx <= 10; idx++)
                for (idx = 0; idx < 10; idx++)
                {
                    var rows = File.ReadAllLines(String.Format("{0}.txt", (2019 - idx)), Encoding.Default).Select(l => l.Split(',').ToArray()).ToArray();
                    //MessageBox.Show(String.Format("{0}年除權息總計 [{1}] rows", (2019-idx),rows.Length));
                    updateMsgNew("clear");
                    double before_price = 0.0;
                    double after_price = 0.0;
                    //組sql字串 
                    String sql = String.Empty;
                    String updateNum = "";
                    //for (int i = 0; i < 10; i++)
                    for (int i = 0; i < rows.Length; i++)
                    {
                        if (rows[i].Length == 17)
                            if (rows[i][6].Length > 0)
                            {
                                after_price = Convert.ToDouble(rows[i][6]);
                                before_price = after_price + Convert.ToDouble(rows[i][16]);

                                updateMsgNew(String.Format("[{0}]{1} 除權息日期[{2}]{6}除權後價格:{3} 股權股息:{4} 除權前價格:{5}", rows[i][1], rows[i][2], rows[i][5], rows[i][6], rows[i][16], before_price, Environment.NewLine));
                                sql = String.Format("INSERT OR replace into stock_dividen (stock_code,stock_name,dividen_date,before_price,after_price,dividen,compare_date) values ('{0}','{1}','{2}',{3},{4},{5},'{6}')", rows[i][1], rows[i][2], rows[i][5], before_price, rows[i][6], rows[i][16], rows[i][5].Replace("/", ""));
                                //updateMsgNew(sql);
                                updateNum = DB_SQL(sql.ToString());
                                if (updateNum.All(Char.IsDigit))
                                {
                                    if (Convert.ToInt32(updateNum) > 0)
                                        updateMsgNew("寫入資料成功(" + updateNum.ToString() + ")");
                                    else
                                        updateMsgNew("寫入資料失敗(" + updateNum.ToString() + ")");
                                }
                                else
                                    if (updateNum.ToLower().Contains("unique"))
                                    updateMsgNew("資料重覆->更新資料");
                                else
                                    updateMsgNew("[錯誤]" + updateNum);
                            }
                    }
                    updateMsgNew(String.Format("=================================={2}統計{0}年總共有{1}公司發放股息股利{2}==================================",
                        (2019 - idx), rows.Length, Environment.NewLine));
                    Thread.Sleep(1500);
                    MessageBox.Show((2019 - idx) + "年資料剖析完畢", (2019 - idx) + "年總計" + rows.Length + "筆資料");

                }
                updateMsgNew("成功! 已抓取完近十年所有發股息股利資料!");
            }).Start();
        }

        //爬取資料->更新股票列表
        private void UpdateList_Click(object sender, EventArgs e)
        {
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            tabControl1.SelectedTab = tabPage1;
            label3.Visible = false;

            String sql = String.Empty;
            String updateNum = String.Empty;
            Dictionary<String, String> pulic_company = null;
            if (stock_list.Count == 0)
            {
                //預設先建立股票清單
                pulic_company = getPublic();
                stock_list = getETF().Union(pulic_company).ToDictionary(item => item.Key, item => item.Value);
            }
            new Thread(() =>
            {
                foreach (var item in stock_list)
                {
                    try
                    {

                        sql = String.Format("INSERT OR replace into stock_profile (stock_code,stock_name,update_time) values ('{0}','{1}','{2}') ", item.Key, item.Value, DateTime.Now.ToString("yyyy/MM/dd"));

                        updateNum = DB_SQL(sql);
                        if (updateNum.All(Char.IsDigit))
                        {
                            if (Convert.ToInt32(updateNum) > 0)
                                updateMsgNew(String.Format("[{0}]{1}{2}更新資料庫 成功{3}筆", item.Key, item.Value, Environment.NewLine, updateNum));
                            else
                                updateMsgNew(String.Format("[{0}]{1}{2}更新資料庫 失敗({3})", item.Key, item.Value, Environment.NewLine, updateNum));
                        }
                        else
                            if (updateNum.ToLower().Contains("unique"))
                            updateMsgNew(String.Format("[{0}]{1}資料重覆->略過資料", item.Key, item.Value));
                        else
                            updateMsgNew("[錯誤]" + updateNum);
                    }
                    catch (Exception eex)
                    {
                        updateMsgNew(eex.StackTrace.ToString());
                    }
                }
                updateMsgNew(String.Format("總計更新{0}隻股票名稱&代碼 ", stock_list.Count));
            }).Start();
        }

        //爬取資料->更新最新股利
        private void 更新最新資料ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ////hide track bar
            //trackBar1.Visible = false;
            //label4.Visible = false;
            //label5.Visible = false;
            //button3.Visible = false;

            //tabControl1.SelectedTab = tabPage1;
            ////MessageBox.Show("還沒做唷 ♥","等我有空❤");
            //label3.Visible = false;

            //String now_year = DateTime.Now.ToString("yyyy");

            //String stock_code = "";
            //String stock_name = "";
            //String dividen_date = "";
            //String after_price = "";
            //String before_price = "";
            //String cash_dividen = "";
            //String stock_dividen = "";
            //String dividen = "";
            //String compare_date = "";

            //MessageBox.Show("最新交易年份：" + now_year, "開始更新最新股利");
            //try
            //{
            //    //預設撈上市公司 (含ETF)
            //    String url = "https://goodinfo.tw/StockInfo/StockDividendScheduleList.asp?MARKET_CAT=%E4%B8%8A%E5%B8%82&INDUSTRY_CAT=%E5%85%A8%E9%83%A8&YEAR=" + now_year;
            //    String xPath = "//*[@id=\"divDetail\"]/table";
            //    using (WebClient client = new WebClient())
            //    {
            //        //新增user-agent避免被認為是機械人
            //        client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36");
            //        //note：調整protocol
            //        ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            //        MemoryStream ms = new MemoryStream(client.DownloadData(url));
            //        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            //        // 使用預設編碼讀入 HTML  
            //        doc.Load(ms, Encoding.UTF8);

            //        // 裝載第一層查詢結果 
            //        HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();

            //        hdc.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);

            //        // 取得個股標頭 
            //        HtmlNodeCollection rows = hdc.DocumentNode.SelectNodes("./tr");
            //        //MessageBox.Show("tr num = "+tr.Count);

            //        //使用同一個連線
            //        using (SQLiteConnection conn = new SQLiteConnection(connStr))
            //        {
            //            conn.Open();
            //            using (SQLiteTransaction trans = conn.BeginTransaction())
            //            {
            //                for (int i = 0; i < rows.Count; i++)
            //                {
            //                    try
            //                    {
            //                        stock_code = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[2]").InnerText.Trim().Replace("&nbsp;", "");
            //                        stock_name = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[3]").InnerText.Trim().Replace("&nbsp;", "");
            //                        //排除只有發權不發息的狀況
            //                        //if(stock_code == "4566")
            //                        if( hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[6]").InnerText.Trim().Replace("&nbsp;", "").Length ==0)
            //                            dividen_date = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[8]").InnerText.Trim().Replace("&nbsp;", "");
            //                        else//只有發息的日期
            //                            dividen_date = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[6]").InnerText.Trim().Replace("&nbsp;", "");

            //                        //排除只有發權不發息的狀況
            //                        if (hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[7]").InnerText.Trim().Replace("&nbsp;", "").Length == 0)
            //                            after_price = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[9]").InnerText.Trim().Replace("&nbsp;", "");
            //                        else
            //                            after_price = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[7]").InnerText.Trim().Replace("&nbsp;", "");

            //                        if (!dividen_date.Contains("(") && after_price.Length>0)    //不是即將除權息的資料才寫入
            //                        {
            //                            //before_price =  hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[4]").InnerText.Trim();


            //                            cash_dividen = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[13]").InnerText.Trim().Replace("&nbsp;", "");
            //                            stock_dividen = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[16]").InnerText.Trim().Replace("&nbsp;", "");
            //                            dividen = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[17]").InnerText.Trim().Replace("&nbsp;", "");
            //                            //MessageBox.Show(String.Format("[{0}] {1} dividen_date:{2} before_price:{3} after_price:{4} cash_dividen:{5} stock_dividen:{6} dividen:{7} {8}"
            //                            //    , stock_code, stock_name, dividen_date, before_price, after_price, cash_dividen, stock_dividen, dividen, Environment.NewLine));
            //                            updateMsgNew(String.Format("[{0}] {1} dividen_date:{2} before_price:{3} after_price:{4} cash_dividen:{5} stock_dividen:{6} dividen:{7} {8}"
            //                                , stock_code, stock_name, dividen_date, before_price, after_price, cash_dividen, stock_dividen, dividen, Environment.NewLine));
            //                            compare_date = dividen_date.Replace("/", "");
            //                            before_price = (Convert.ToDouble(after_price) + Convert.ToDouble(dividen)).ToString();
            //                            String sql = String.Format("replace into stock_dividen (stock_code,stock_name,cash_dividen,stock_dividen,dividen,dividen_date,compare_date,before_price,after_price) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')",
            //                                stock_code, stock_name, cash_dividen, stock_dividen, dividen, dividen_date, compare_date, before_price, after_price);

            //                            updateMsgNew(sql);
            //                            String updateNum = "";
            //                            //☆ 主要修改 只使用一個連線
            //                            using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
            //                            {
            //                                updateNum = cmd.ExecuteNonQuery().ToString();


            //                                if (updateNum.All(Char.IsDigit))
            //                                {
            //                                    if (Convert.ToInt32(updateNum) > 0)
            //                                        updateMsgNew("[" + stock_code + "]更新股利成功(" + updateNum.ToString() + ")");
            //                                    else
            //                                        updateMsgNew("更新股利失敗(" + updateNum.ToString() + ")");
            //                                }
            //                                else
            //                                    updateMsgNew("[錯誤]" + updateNum);
            //                            }

            //                        }
            //                    }
            //                    catch(Exception eee)
            //                    {
            //                        //MessageBox.Show("[錯誤]" + eee.Message + ","+ dividen_date, "更新股利錯誤");
            //                        MessageBox.Show("[錯誤]" + eee.Message, "更新股利錯誤");
            //                    }


            //                }//end of for loop

            //                trans.Commit();
            //            }
            //            conn.Close();
            //        }

            //    }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("[錯誤]" + ex.Message, "更新股利錯誤");
            //}


            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            tabControl1.SelectedTab = tabPage1;
            //MessageBox.Show("還沒做唷 ♥","等我有空❤");
            label3.Visible = false;

            String now_year = DateTime.Now.ToString("yyyy");

            String stock_code = "";
            String stock_name = "";
            String dividen_date = "";
            String after_price = "";
            String before_price = "";
            String cash_dividen = "";
            String stock_dividen = "";
            String dividen = "";
            String compare_date = "";
            String least_days = "";
            int success_num = 0;        //成功更新股利筆數
            MessageBox.Show("最新交易年份：" + now_year, "開始更新最新股利");
            new Thread(() =>
            {
                try
                {
                    //預設撈上市公司 (含ETF)
                    String url = "https://goodinfo.tw/StockInfo/StockDividendScheduleList.asp?MARKET_CAT=%E4%B8%8A%E5%B8%82&INDUSTRY_CAT=%E5%85%A8%E9%83%A8&YEAR=" + now_year;
                    String xPath = "//*[@id=\"divDetail\"]/table";
                    //測速
                    DateTime t1 = DateTime.Now;

                    //使用同一個連線
                    using (SQLiteConnection conn = new SQLiteConnection(connStr))
                    {
                        conn.Open();
                        using (SQLiteTransaction trans = conn.BeginTransaction())
                        {
                            for (int k = 2020; k <= 2020; k++)
                            {
                                using (WebClient client = new WebClient())
                                {
                                    //新增user-agent避免被認為是機械人
                                    client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36");
                                    //note：調整protocol
                                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                                    url = "https://goodinfo.tw/StockInfo/StockDividendScheduleList.asp?MARKET_CAT=%E4%B8%8A%E5%B8%82&INDUSTRY_CAT=%E5%85%A8%E9%83%A8&YEAR=" + k.ToString();

                                    MemoryStream ms = new MemoryStream(client.DownloadData(url));
                                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                    // 使用預設編碼讀入 HTML  
                                    doc.Load(ms, Encoding.UTF8);

                                    // 裝載第一層查詢結果 
                                    HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();

                                    hdc.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);

                                    // 取得個股標頭 
                                    HtmlNodeCollection rows = hdc.DocumentNode.SelectNodes("./tr");
                                    //MessageBox.Show("tr num = "+tr.Count);

                                    for (int i = 0; i < rows.Count; i++)
                                    {
                                        try
                                        {
                                            stock_code = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[2]").InnerText.Trim().Replace("&nbsp;", "");
                                            stock_name = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[3]").InnerText.Trim().Replace("&nbsp;", "");
                                            //排除只有發權不發息的狀況
                                            //if(stock_code == "4566")
                                            if (hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[5]").InnerText.Trim().Replace("&nbsp;", "").Length == 0)
                                                dividen_date = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[10]").InnerText.Trim().Replace("&nbsp;", "").Replace("即將除息", "");
                                            else//只有發息的日期
                                                dividen_date = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[5]").InnerText.Trim().Replace("&nbsp;", "").Replace("即將除息", "");
                                            //配合調整日期格式
                                            dividen_date = String.Format("20{0}", dividen_date.Replace("'", "/"));

                                            //排除只有發權不發息的狀況
                                            if (hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[6]").InnerText.Trim().Replace("&nbsp;", "").Length == 0)
                                                after_price = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[11]").InnerText.Trim().Replace("&nbsp;", "");
                                            else
                                                after_price = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[6]").InnerText.Trim().Replace("&nbsp;", "");

                                            if (!dividen_date.Contains("(") && after_price.Length > 0)    //不是即將除權息的資料才寫入
                                            {
                                                //before_price =  hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[4]").InnerText.Trim();


                                                cash_dividen = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[16]").InnerText.Trim().Replace("&nbsp;", "");
                                                stock_dividen = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[19]").InnerText.Trim().Replace("&nbsp;", "");
                                                dividen = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[20]").InnerText.Trim().Replace("&nbsp;", "");
                                                //MessageBox.Show(String.Format("[{0}] {1} dividen_date:{2} before_price:{3} after_price:{4} cash_dividen:{5} stock_dividen:{6} dividen:{7} {8}"
                                                //    , stock_code, stock_name, dividen_date, before_price, after_price, cash_dividen, stock_dividen, dividen, Environment.NewLine));
                                                updateMsgNew(String.Format("[{0}] {1} dividen_date:{2} before_price:{3} after_price:{4} cash_dividen:{5} stock_dividen:{6} dividen:{7} {8}"
                                                    , stock_code, stock_name, dividen_date, before_price, after_price, cash_dividen, stock_dividen, dividen, Environment.NewLine));
                                                compare_date = dividen_date.Replace("/", "");
                                                before_price = (Convert.ToDouble(after_price) + Convert.ToDouble(dividen)).ToString();
                                                least_days = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[8]").InnerText.Trim().Replace("&nbsp;", "");
                                                String sql = String.Format("INSERT OR replace into stock_dividen (stock_code,stock_name,cash_dividen,stock_dividen,dividen,dividen_date,compare_date,before_price,after_price,least_days) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')",
                                                    stock_code, stock_name, cash_dividen, stock_dividen, dividen, dividen_date, compare_date, before_price, after_price, least_days);

                                                updateMsgNew(sql);
                                                String updateNum = "";
                                                //☆ 主要修改 只使用一個連線
                                                using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                                                {
                                                    updateNum = cmd.ExecuteNonQuery().ToString();


                                                    if (updateNum.All(Char.IsDigit))
                                                    {
                                                        if (Convert.ToInt32(updateNum) > 0)
                                                        {
                                                            updateMsgNew(k + "[" + stock_code + "]更新股利成功(" + updateNum.ToString() + ")");
                                                            success_num++;
                                                        }
                                                        else
                                                            updateMsgNew("更新股利失敗(" + updateNum.ToString() + ")");
                                                    }
                                                    else
                                                        updateMsgNew("[錯誤]" + updateNum);
                                                }

                                            }
                                        }
                                        catch (Exception eee)
                                        {
                                            //MessageBox.Show("[錯誤]" + eee.Message + ","+ dividen_date, "更新股利錯誤");
                                            MessageBox.Show("[錯誤]" + eee.Message, "更新股利錯誤");
                                        }
                                    }//end of for loop
                                }//end of using webclient
                            }//end of for loop (int k = 2008; k <= 2019; k++)
                            trans.Commit();
                        }
                        conn.Close();
                    }
                    TimeSpan t = (DateTime.Now - t1);
                    updateMsgNew(String.Format("最新年份[{0}] 總計更新{1}筆股利資料 費時{2}秒", now_year, success_num, t.TotalSeconds));

                    //移除所有重覆資料
                    String sq1l = "delete   from stock_dividen where rowid not in ( " +
                        " select  min(rowid) " +
                        "  from stock_dividen " +
                        "  group by " +
                        "          stock_code " +
                        "  ,       dividen_date " +
                        "  ) ";
                    DB_SQL(sq1l);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("[錯誤]" + ex.Message, "更新股利錯誤");
                }
            }).Start();
        }



        private void 爬取股票列表ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //hide track bar
                trackBar1.Visible = false;
                label4.Visible = false;
                label5.Visible = false;
                button3.Visible = false;

                tabControl1.SelectedTab = tabPage1;
                label3.Visible = false;

                Dictionary<String, String> stock_list = new Dictionary<string, String>();

                stock_list = getETF();
                int etf_num = stock_list.Count;

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();

                //取得最新交易日期
                DateTime latest_date = getLatestDate();

                hdc.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/table[1]/tbody[1]").InnerHtml);
                // 取得個股標頭 
                HtmlNodeCollection nodeHeaders = hdc.DocumentNode.SelectNodes("./tr");

                updateMsgNew("正在讀取網頁清單 請稍後...");
                // 輸出資料 
                updateMsgNew(String.Format("總計{0}筆股票", nodeHeaders.Count));
                updateMsgNew("=========================");
                //string[] values = hdc.DocumentNode.SelectSingleNode("./tr[1]").InnerText.Trim();
                StringBuilder sb = new StringBuilder();
                StringBuilder sb_sql = new StringBuilder();
                //a = DateTime.Now;
                //TimeSpan t = DateTime.Now - a;
                //updateMsgNew("等待網頁回應資料花費時間" + t.TotalSeconds + "秒");

                for (int i = 0; i < nodeHeaders.Count; i++)
                //for (int i = 0; i < 2; i++)
                {
                    String stock_code = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[1]").InnerText.Trim();
                    String stock_name = hdc.DocumentNode.SelectSingleNode("./tr[" + (i + 1) + "]/td[2]").InnerText.Trim();
                    if (!stock_list.ContainsKey(stock_code))
                    {
                        stock_list.Add(stock_code, stock_name);
                    }
                }

                String sql = "";
                String updateNum = "";
                String date = DateTime.Now.ToString("yyyy/MM/dd");
                foreach (var item in stock_list)
                {
                    try
                    {
                        sb.Append(String.Format("[{0}]{1}{2} ", item.Key, item.Value, Environment.NewLine));
                        sql = String.Format("INSERT OR replace into stock_profile (stock_code,stock_name,update_time) values ('{0}','{1}','{2}') ", item.Key, item.Value, date);

                        updateNum = DB_SQL(sql);
                        if (updateNum.All(Char.IsDigit))
                        {
                            if (Convert.ToInt32(updateNum) > 0)
                                updateMsgNew(String.Format("[{0}]{1}{2}更新資料庫 成功{3}筆", item.Key,item.Value,Environment.NewLine,updateNum));
                            else
                                updateMsgNew(String.Format("[{0}]{1}{2}更新資料庫 失敗({3})", item.Key,item.Value,Environment.NewLine,updateNum));
                        }
                        else
                            if (updateNum.ToLower().Contains("unique"))
                                updateMsgNew("資料重覆->更新資料");
                            else
                                updateMsgNew("[錯誤]" + updateNum);
                    }
                    catch (Exception eex)
                    {
                        updateMsgNew(eex.StackTrace.ToString());
                    }
                }
                updateMsgNew(sb.ToString());
                updateMsgNew(String.Format("總計{0}筆ETF {1}筆上市股票 總計{2}筆", etf_num, nodeHeaders.Count, etf_num + nodeHeaders.Count));
            }
            catch (Exception ee)
            {
                if (ee.Message.Contains("連線被拒"))
                    updateMsg("連線過度頻繁 被Ban惹! =>" + ee.Message);
                else
                    updateMsg(ee.Message);
            }
        }

        private void 單一分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage2;
            label1.Visible = true;
            //button1.Visible = true;
            button9.Visible = true;
            textBox1.Visible = true;

            label2.Visible = false;
            textBox3.Visible = false;
            label3.Visible = false;
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

        }

        private void 關於ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            //ProcessStartInfo Info = new ProcessStartInfo();
            //Info.FileName = "intro.vbs";//執行的檔案名稱
            //Info.WorkingDirectory = @"E:\NVR";//檔案所在的目錄
            //Process.Start(Info);// RUN bat

            MessageBox.Show("爬蟲分析智慧選股排名","發大$  By Kuroboy");
            //預設先建立股票清單
            //Dictionary<String,String> pulic_company = getPublic();
            //stock_list = getETF().Union(pulic_company).ToDictionary(item => item.Key, item => item.Value);
            //foreach (var item in stock_list)
            //    updateMsgNew(String.Format("[{0}]{1}", item.Key, item.Value));
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                MessageBox.Show("未輸入代碼", "錯誤");
            else
            {
                String stock_code = textBox1.Text;
                //檢查資料夾是否存在 (1)不存在->建立  (2)存在->寫檔案
                String saveFolder = String.Format(@"c://CSV//{0}//", textBox1.Text);
                if (!Directory.Exists(saveFolder))
                {
                    MessageBox.Show("尚未抓取過資料->建立資料夾 [" + textBox1.Text + "]");
                    Directory.CreateDirectory(saveFolder);
                    updateMsg("建立資料夾[" + textBox1.Text + "]成功");
                }

                new Thread(() =>
                {
                    CrawlHtml(textBox1.Text);
                }).Start();

            }
        }

        private void CrawlHtml(String stock_code)
        {
            try
            {
                DateTime latest_date = getLatestDate();
                //☆需要修改的地方
                String url = String.Format("https://www.cnyes.com/twstock/ps_historyprice.aspx?code={0}&ctl00$ContentPlaceHolder1$startText=2008/01/01&ctl00$ContentPlaceHolder1$endText={1}", stock_code, latest_date.ToString("yyyy/MM/dd"));
                String xPath = "//*[@id=\"main3\"]/div[5]/div[3]/table";

                WebClient client = new WebClient();
                MemoryStream ms = new MemoryStream(client.DownloadData(url));
                // 使用utf8編碼讀入 HTML 才能正確顯示中文 
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.UTF8);
                // 裝載第一層查詢結果 
                HtmlAgilityPack.HtmlDocument tableContent = new HtmlAgilityPack.HtmlDocument();
                tableContent.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);


                DateTime t1 = DateTime.Now;
                int i = 0;

                //版本1：使用函數 分次插入db
                //foreach (HtmlNode row in tableContent.DocumentNode.SelectNodes("./tr")) //☆需要修改的地方
                //{
                //    if (i == 0)
                //    {
                //        //HtmlNodeCollection title = row.SelectNodes("th");
                //        //updateMsg(String.Format("標頭 {0} {1} {2} {3} {4}", title[0].InnerText, title[1].InnerText, title[2].InnerText, title[3].InnerText, title[4].InnerText));
                //        updateMsg(String.Format("[{0}]", stock_code));
                //    }
                //    else
                //    {
                //        //☆需要修改的地方
                //        HtmlNodeCollection cells = row.SelectNodes("td");
                //        updateMsg(String.Format("date[{0}] open[{1}] high[{2}] low[{3}] close[{4}]", cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText, cells[4].InnerText));
                //        String sql = String.Format("insert into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7})", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText, cells[2].InnerText, cells[3].InnerText, cells[4].InnerText, cells[5].InnerText);
                //        //updateMsg(sql);
                //        String updateNum = DB_SQL(sql.ToString());
                //        if (updateNum.All(Char.IsDigit))
                //        {
                //            if (Convert.ToInt32(updateNum) > 0)
                //                updateMsg("[" + cells[0].InnerText + "]寫入資料成功(" + updateNum.ToString() + ")");
                //            else
                //                updateMsg("寫入資料失敗(" + updateNum.ToString() + ")");
                //        }
                //        else
                //            updateMsg("[錯誤]" + updateNum);
                //    }
                //    i++;
                //}

                //版本2：使用同一個連線 持續寫入db
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (SQLiteTransaction trans = conn.BeginTransaction())
                    {
                        foreach (HtmlNode row in tableContent.DocumentNode.SelectNodes("./tr")) //☆需要修改的地方
                        {
                            if (i == 0)
                            {
                                updateMsg(String.Format("[{0}]", stock_code));
                            }
                            else
                            {
                                try
                                {
                                    //☆需要修改的地方
                                    HtmlNodeCollection cells = row.SelectNodes("td");

                                    updateMsg(String.Format("date[{0}] open[{1}] high[{2}] low[{3}] close[{4}]", cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText, cells[4].InnerText));

                                    String sql = String.Format("INSERT OR replace into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7})", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText);

                                    //針對破千元股票特別處理
                                    //if (stock_code == "3008")
                                    //     sql = String.Format("insert into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7})", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText + cells[2].InnerText, cells[3].InnerText + cells[4].InnerText, cells[5].InnerText + cells[2].InnerText, cells[7].InnerText + cells[8].InnerText, cells[9].InnerText);

                                    //updateMsg(sql);

                                    String updateNum = "";
                                    //☆ 主要修改 只使用一個連線
                                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                                    {
                                        updateNum = cmd.ExecuteNonQuery().ToString();
                                    }

                                    if (updateNum.All(Char.IsDigit))
                                    {
                                        if (Convert.ToInt32(updateNum) > 0)
                                            updateMsg("[" + cells[0].InnerText + "]寫入資料成功(" + updateNum.ToString() + ")");
                                        else
                                            updateMsg("寫入資料失敗(" + updateNum.ToString() + ")");
                                    }
                                    else
                                        updateMsg("[錯誤]" + updateNum);
                                }
                                catch (Exception e)
                                {
                                    updateMsg("[錯誤]" + e.Message);
                                }

                            }
                            i++;
                        }
                        trans.Commit();
                    }
                    conn.Close();
                }

                TimeSpan ts = DateTime.Now - t1;
                updateMsg(String.Format("總共插入{0}筆資料 費時{1}秒", i, ts.TotalSeconds));
                doc = null;
                tableContent = null;
                client = null;
                ms.Close();
            }
            catch (Exception e)
            {
                updateMsg(e.Message);
            }
        }

        private void updateDividen(String stock_code)
        {
            try
            {
                //☆需要修改的地方
                //String url = String.Format("https://goodinfo.tw/StockInfo/StockDividendScheduleList.asp?MARKET_CAT=全部&INDUSTRY_CAT=全部&YEAR=2019");
                //String xPath = "//*[@id=\"divDetail\"]/table";

                String url = String.Format("https://www.cnyes.com/twstock/ps_historyprice.aspx?code={0}&ctl00$ContentPlaceHolder1$startText=2008/01/01&ctl00$ContentPlaceHolder1$endText=2019/10/06", stock_code);
                String xPath = "//*[@id=\"main3\"]/div[5]/div[3]/table";

                //MemoryStream ms = new MemoryStream(client.DownloadData("http://tw.stock.yahoo.com/q/q?s=" + stock_code));
                //xPath = "/html[1]/body[1]/center[1]/table[2]/tr[1]/td[1]/table[1]";


                WebClient client = new WebClient();
                MemoryStream ms = new MemoryStream(client.DownloadData(url));
                // 使用utf8編碼讀入 HTML 才能正確顯示中文 
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.UTF8);
                // 裝載第一層查詢結果 
                HtmlAgilityPack.HtmlDocument tableContent = new HtmlAgilityPack.HtmlDocument();
                tableContent.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);
                int i = 0;

                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (SQLiteTransaction trans = conn.BeginTransaction())
                    {
                        foreach (HtmlNode row in tableContent.DocumentNode.SelectNodes("./tr")) //☆需要修改的地方
                        {
                            if (i == 0)
                            {
                                updateMsg(String.Format("[{0}]", stock_code));
                            }
                            else
                            {
                                try
                                {
                                    //☆需要修改的地方
                                    HtmlNodeCollection cells = row.SelectNodes("td");

                                    updateMsg(String.Format("date[{0}] open[{1}] high[{2}] low[{3}] close[{4}]", cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText, cells[4].InnerText));

                                    String sql = String.Format("INSERT OR replace into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7})", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText);


                                    String updateNum = "";
                                    //☆ 主要修改 只使用一個連線
                                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                                    {
                                        updateNum = cmd.ExecuteNonQuery().ToString();
                                    }

                                    if (updateNum.All(Char.IsDigit))
                                    {
                                        if (Convert.ToInt32(updateNum) > 0)
                                            updateMsg("[" + cells[0].InnerText + "]寫入資料成功(" + updateNum.ToString() + ")");
                                        else
                                            updateMsg("寫入資料失敗(" + updateNum.ToString() + ")");
                                    }
                                    else
                                        updateMsg("[錯誤]" + updateNum);
                                }
                                catch (Exception e)
                                {
                                    updateMsg("[錯誤]" + e.Message);
                                }

                            }
                            i++;
                        }
                        trans.Commit();
                    }
                    conn.Close();
                }

                //TimeSpan ts = DateTime.Now - t1;
                //updateMsg(String.Format("總共插入{0}筆資料 費時{1}秒", i, ts.TotalSeconds));
                doc = null;
                tableContent = null;
                client = null;
                ms.Close();
            }
            catch (Exception e)
            {
                updateMsg(e.Message);
            }
        }
        //function overloading
        private void CrawlHtml(String stock_code,String sDate,String eDate)
        {
            try
            {
                dynamic rows = null;
                DateTime sDateNew = Convert.ToDateTime(sDate);

                updateMsgNew(String.Format("更新股票代碼「{0}」日期區間{1}~{2}",stock_code,sDate,eDate));
                //sDateNew = sDateNew.AddDays(-sub_days);
                //☆需要修改的地方
                String url = String.Format("https://www.cnyes.com/twstock/ps_historyprice.aspx?code={0}&ctl00$ContentPlaceHolder1$startText={1}&ctl00$ContentPlaceHolder1$endText={2}", stock_code, sDate, eDate) ;
                String xPath = "//*[@id=\"main3\"]/div[5]/div[3]/table";

                DateTime t1 = DateTime.Now;
                WebClient client = new WebClient();
                MemoryStream ms = new MemoryStream(client.DownloadData(url));
                // 使用utf8編碼讀入 HTML 才能正確顯示中文 
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.UTF8);
                // 裝載第一層查詢結果 
                HtmlAgilityPack.HtmlDocument tableContent = new HtmlAgilityPack.HtmlDocument();
                tableContent.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);
                rows =  tableContent.DocumentNode.SelectNodes("./tr");

                
                int i = 0;
                String sql = String.Empty;
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (SQLiteTransaction trans = conn.BeginTransaction())
                    {
                        foreach (HtmlNode row in tableContent.DocumentNode.SelectNodes("./tr")) //☆需要修改的地方
                        {
                            if (i == 0)
                            {
                                updateMsgNew(String.Format("[{0}]{1}", stock_code,stock_list[stock_code]));
                            }
                            else
                            {
                                try
                                {
                                    //☆需要修改的地方
                                    HtmlNodeCollection cells = row.SelectNodes("td");

                                    updateMsgNew(String.Format("date[{0}] open[{1}] high[{2}] low[{3}] close[{4}] quantity[{5}]", cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText, cells[4].InnerText, cells[8].InnerText));

                                    //sql = sql + String.Format("INSERT OR replace into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change,quantity) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7},{8});", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText, cells[8].InnerText.Replace(",", ""));
                                    //no quantity
                                    sql = sql + String.Format("INSERT OR replace into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7});", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText);


                                }
                                catch (Exception e)
                                {
                                    updateMsgNew("[錯誤]" + e.Message);
                                }

                            }
                            i++;
                        }
                        String updateNum = "";
                        //☆ 主要修改 只使用一個連線
                        using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                        {
                            updateNum = cmd.ExecuteNonQuery().ToString();
                        }

                        if (updateNum.All(Char.IsDigit))
                        {
                            if (Convert.ToInt32(updateNum) > 0)
                                updateMsgNew("寫入資料成功(" + updateNum.ToString() + ")");
                            else
                                updateMsgNew("寫入資料失敗(" + updateNum.ToString() + ")");
                        }
                        else
                            if (updateNum.ToLower().Contains("unique"))
                            updateMsg("資料重覆->更新資料");
                        else
                            updateMsg("[錯誤]" + updateNum);
                        trans.Commit();
                    }
                    conn.Close();
                }

                TimeSpan ts = DateTime.Now - t1;
                updateMsgNew(String.Format("插入.更新{0}筆資料 費時{1}秒", rows.Count, ts.TotalSeconds));
                doc = null;
                tableContent = null;
                client = null;
                ms.Close();
            }
            catch (Exception e)
            {
                updateMsgNew(e.Message);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            new Thread(() =>
            {
                DialogResult result = MessageBox.Show("此功能將花費極大量時間(超過6小時)是否確定要下載?","大量下載警告",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                if (result == DialogResult.No)
                {
                    MessageBox.Show("選擇取消下載", "取消下載");
                }
                else
                {
                    String stock_code = "";
                    if (stock_list.Count == 0)
                        getStockList();
                    foreach (var item in stock_list)
                    {
                        stock_code = item.Key;
                        CrawlHtml(stock_code);
                    }
                }
            }).Start();
        }

        private void 每日精選ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void 發財排行榜ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //show track bar
            trackBar1.Visible = true;
            label4.Visible = true;
            label5.Visible = true;
            button3.Visible = true;

            ProcessStartInfo Info = new ProcessStartInfo();
            Info.FileName = "fada.vbs";//執行的檔案名稱
            Process.Start(Info);// RUN bat

            tabControl1.SelectedTab = tabPage5;
            label3.Text = String.Format("說明：挑出10年內(1)連續{0}年發股息 (2)全部填權息 ☆(3)WADAR指標前五十名 『總排行榜』", trackBar1.Value);
            label3.Visible = true;
            int trackBarValue = 0;
            this.Invoke(new MethodInvoker(delegate { trackBarValue = trackBar1.Value; }));
            new Thread(() =>
            {
                String sql = " select stock_code,stock_name,win_times,dividen_times,avg_win_rate,dividen from money_rank_new " +
                             " where dividen_times >= "+ trackBarValue.ToString() +
                             " order by (win_times/dividen_times) desc ,dividen_times desc,avg_win_rate desc limit 0,50";
                String result = DB_SQL(sql, "stock_code,stock_name,win_times,dividen_times,avg_win_rate,dividen");

                CreateGVData(dataGridView2, sql);
            }).Start();
        }

        private void CreateGVData(DataGridView gv, String sql)
        {
            this.Invoke((MethodInvoker)delegate
            {

                DataTable newTable = new DataTable();
                //table = newTable;
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                    {
                        using (SQLiteDataReader dr = cmd.ExecuteReader())
                        {
                            newTable.Load(dr);
                        }
                    }
                }

                gv.DataSource = newTable;
                gv.AllowUserToAddRows = false;

                if (gv == dataGridView1)
                {
                    dataGridView1.Columns[0].HeaderText = "代碼";
                    dataGridView1.Columns[1].HeaderText = "名稱";
                    dataGridView1.Columns[2].HeaderText = "填息次數";
                    dataGridView1.Columns[3].HeaderText = "發息次數";
                    dataGridView1.Columns[4].HeaderText = "☆WADAR指標";
                    dataGridView1.Columns[5].HeaderText = "最新收盤價";
                    dataGridView1.Columns[6].HeaderText = "填權息價格";
                    dataGridView1.Columns[7].HeaderText = "填權息價差";
                    dataGridView1.Columns[8].HeaderText = "填權息價差%";
                    dataGridView1.Columns[9].HeaderText = "最新除權日期";
                    dataGridView1.Columns[10].HeaderText = "最新收盤日期";

                    if (dataGridView1.Columns.Count == 12)
                        dataGridView1.Columns[11].HeaderText = "最小填息天數";
                    //datagridview前面的空白部分去除
                    dataGridView1.RowHeadersVisible = false; 
                }
                if (gv == dataGridView2)
                {
                    dataGridView2.Columns[0].HeaderText = "代碼";
                    dataGridView2.Columns[1].HeaderText = "名稱";
                    dataGridView2.Columns[2].HeaderText = "填息次數";
                    dataGridView2.Columns[3].HeaderText = "發息次數";
                    dataGridView2.Columns[4].HeaderText = "☆WADAR指標";
                    dataGridView2.Columns[5].HeaderText = "股利息總合";

                    //datagridview前面的空白部分去除
                    dataGridView2.RowHeadersVisible = false;
                }

                if (gv == dataGridView3 && dataGridView3.Columns.Count == 10)
                {
                    dataGridView3.Columns[0].HeaderText = "代碼";
                    dataGridView3.Columns[1].HeaderText = "名稱";
                    dataGridView3.Columns[2].HeaderText = "追踨價格";
                    dataGridView3.Columns[3].HeaderText = "追踨日期";
                    dataGridView3.Columns[4].HeaderText = "最新價格";
                    dataGridView3.Columns[5].HeaderText = "最新日期";
                    dataGridView3.Columns[6].HeaderText = "損益價差";
                    dataGridView3.Columns[7].HeaderText = "損益率(%)";
                    dataGridView3.Columns[8].HeaderText = "填權息價格";
                    dataGridView3.Columns[9].HeaderText = "最新除權日期";
                    //datagridview前面的空白部分去除
                    dataGridView3.RowHeadersVisible = false;

                    if (dataGridView3.Columns.Count == 10)
                    { 
                        DataGridViewButtonColumn btnDel = new DataGridViewButtonColumn();
                        btnDel.Name = "btnDel";
                        btnDel.Text = "刪除";

                        btnDel.HeaderText = "動作";
                        btnDel.UseColumnTextForButtonValue = true;
                        dataGridView3.Columns.Add(btnDel);
                        //DataGridViewButtonColumn btnMod = new DataGridViewButtonColumn();
                        //btnDel.Name = "btnMod";
                        //btnDel.Text = "修改";
                    }
                }
            });
        }

        public void updateTable(DataGridView gv, String result)
        {
            //一行資料
            String[] row = null;
            this.Invoke((MethodInvoker)delegate
            {
                //所有行資料
                String[] rows = result.Split(';');
                for(int i=0;i<rows.Length-1;i++)
                {
                    row = rows[i].Split('@');
                    for (int j = 0; j < row.Length-1; j++)
                        gv.Rows[i].Cells[j].Value = row[j];
                }

            }); 
        }

        //取得最新交易日期
        public DateTime getLatestDate()
        {
            //=============================================================================================================================================
            //取得最新交易日期
            //=============================================================================================================================================
            DateTime latest_date = DateTime.Now;
            DateTime end_date = DateTime.Now;
            int sub_days = 0;
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            //選擇回第一個分頁
            tabControl1.SelectedTab = tabPage1;

            HtmlAgilityPack.HtmlDocument tableContent = new HtmlAgilityPack.HtmlDocument();
            dynamic rows = null;
            do
            {
                latest_date = latest_date.AddDays(-sub_days);
                //☆需要修改的地方
                String url = String.Format("https://www.cnyes.com/twstock/ps_historyprice.aspx?code={0}&ctl00$ContentPlaceHolder1$startText={1}&ctl00$ContentPlaceHolder1$endText={2}", "0050", latest_date.ToString("yyyy/MM/dd"), end_date.ToString("yyyy/MM/dd"));
                String xPath = "//*[@id=\"main3\"]/div[5]/div[3]/table";
                WebClient client = new WebClient();
                MemoryStream ms = new MemoryStream(client.DownloadData(url));
                // 使用utf8編碼讀入 HTML 才能正確顯示中文 
                doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.UTF8);
                // 裝載第一層查詢結果 
                tableContent.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);
                rows = tableContent.DocumentNode.SelectNodes("./tr");
                //updateMsgNew(String.Format("檢查日期[{0}]", latest_date.ToString("yyyy/MM/dd")));
                sub_days++;
            } while (rows.Count == 1);
            //MessageBox.Show("最新交易日：" + latest_date.ToString("yyyy-MM-dd"), "最新交易日期");
            //=============================================================================================================================================
            return latest_date;
        }

        private void 更新最新股價ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            tabControl1.SelectedTab = tabPage1;
            label3.Visible = false;

            //取得最新交易日期
            String latest_trade_date = getLatestDate().ToString("yyyy/MM/dd");

            //從資料庫找出股價資料最新資料之日期
            String latest_date = DB_SQL("select max(stock_date) latest_date from stock_price_new ", "latest_date");

            //just for test
            //latest_date = "2020/02/20";
            MessageBox.Show(String.Format("股價資料將從 資料庫中最舊日期[{0}]更新至 最新交易日期[{1}]",latest_date, latest_trade_date));

            //計算目前下載到第幾隻股票
            int num = 0;
            String sql = String.Empty;

            new Thread(() =>
            {
                try
                {
                    DateTime t1 = DateTime.Now;

                    using (SQLiteConnection conn = new SQLiteConnection(connStr))
                    {
                        conn.Open();
                        using (SQLiteTransaction trans = conn.BeginTransaction())
                        {
                            foreach (var item in stock_list)
                            {
                                num++;
                                //CrawlHtml(item.Key, latest_date, latest_trade_date);
                                String stock_code = item.Key;
                                try
                                {
                                    dynamic rows = null;
                                    DateTime sDateNew = Convert.ToDateTime(latest_date);

                                    updateMsgNew(String.Format("更新股票代碼「{0}」日期區間{1}~{2}", stock_code, latest_date, latest_trade_date));
                                    //sDateNew = sDateNew.AddDays(-sub_days);
                                    //☆需要修改的地方
                                    String url = String.Format("https://www.cnyes.com/twstock/ps_historyprice.aspx?code={0}&ctl00$ContentPlaceHolder1$startText={1}&ctl00$ContentPlaceHolder1$endText={2}", stock_code, latest_date, latest_trade_date);
                                    String xPath = "//*[@id=\"main3\"]/div[5]/div[3]/table";

                                    DateTime t2 = DateTime.Now;
                                    WebClient client = new WebClient();
                                    MemoryStream ms = new MemoryStream(client.DownloadData(url));
                                    // 使用utf8編碼讀入 HTML 才能正確顯示中文 
                                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                    doc.Load(ms, Encoding.UTF8);
                                    // 裝載第一層查詢結果 
                                    HtmlAgilityPack.HtmlDocument tableContent = new HtmlAgilityPack.HtmlDocument();
                                    tableContent.LoadHtml(doc.DocumentNode.SelectSingleNode(xPath).InnerHtml);
                                    rows = tableContent.DocumentNode.SelectNodes("./tr");


                                    int i = 0;
                                    
                            
                                    foreach (HtmlNode row in tableContent.DocumentNode.SelectNodes("./tr")) //☆需要修改的地方
                                    {
                                        if (i == 0)
                                        {
                                            updateMsgNew(String.Format("[{0}]{1}", stock_code, stock_list[stock_code]));
                                        }
                                        else
                                        {

                                                //☆需要修改的地方
                                                HtmlNodeCollection cells = row.SelectNodes("td");

                                                updateMsgNew(String.Format("date[{0}] open[{1}] high[{2}] low[{3}] close[{4}] quantity[{5}]", cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText, cells[4].InnerText, cells[8].InnerText));

                                                //sql = sql + String.Format("INSERT OR replace into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change,quantity) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7},{8});", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText, cells[8].InnerText.Replace(",", ""));
                                                //no quantity
                                                sql = sql + String.Format("INSERT OR replace into stock_price_new (stock_code,stock_date,compare_date,open,high,low,close,price_change) values ('{0}','{1}','{2}',{3},{4},{5},{6},{7});", stock_code, cells[0].InnerText, cells[0].InnerText.Replace("/", ""), cells[1].InnerText.Replace(",", ""), cells[2].InnerText.Replace(",", ""), cells[3].InnerText.Replace(",", ""), cells[4].InnerText.Replace(",", ""), cells[5].InnerText);
                                        }
                                        i++;
                                    }
                                    String updateNum = "";
                                    //☆ 主要修改 只使用一個連線
                                    //測試 每10筆才寫入一次看看效果
                                    if (num % 10 == 0)
                                    {
                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                                        {
                                            updateNum = cmd.ExecuteNonQuery().ToString();
                                        }

                                        if (updateNum.All(Char.IsDigit))
                                        {
                                            if (Convert.ToInt32(updateNum) > 0)
                                                updateMsgNew("寫入資料成功(" + updateNum.ToString() + ")");
                                            else
                                                updateMsgNew("寫入資料失敗(" + updateNum.ToString() + ")");
                                        }
                                        else
                                            if (updateNum.ToLower().Contains("unique"))
                                            updateMsg("資料重覆->更新資料");
                                        else
                                            updateMsg("[錯誤]" + updateNum);

                                        //清空sql字句 
                                        sql = String.Empty;
                                    }

                                    TimeSpan ts1 = DateTime.Now - t2;
                                    updateMsgNew(String.Format("插入.更新{0}筆資料 費時{1}秒", rows.Count, ts1.TotalSeconds));
                                    doc = null;
                                    tableContent = null;
                                    client = null;
                                    ms.Close();
                                }
                                catch (Exception eee)
                                {
                                    updateMsgNew(eee.Message);
                                }

                            }//end of foreach loop
                    
                            trans.Commit();
                        }
                        conn.Close();
                    }
                    TimeSpan ts = DateTime.Now - t1;
                    updateMsgNew(String.Format("插入更新{0}檔股票 費時{1}秒", stock_list.Count, ts.TotalSeconds));

                }
                catch (Exception ee)
                {
                    updateMsgNew(ee.Message);
                }
            }).Start();

        }

        private void 績效成果ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //show track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            label2.Visible = true;
            textBox3.Visible = true;

            button9.Visible = false;
            label1.Visible = true;
            textBox1.Visible = true;
            label3.Text = "說明：將代碼及股價「加入追踨」後 將每日計算 (1)是否填權息 (2)獲利差價 (3)獲利率";
            label3.Visible = true;

            //選擇tabPage4
            tabControl1.SelectedTab = tabPage4;

            String sql = "select t.stock_code,t.stock_name,t.track_price,t.track_date,l.close,l.stock_date,(close-track_price) price_diff, Round((close-track_price)/track_price*100,2) diff_percent,d.before_price,d.dividen_date from track_list t left outer join latest_stock_price l on t.stock_code = l.stock_code left outer join latest_stock_dividen d on t.stock_code = d.stock_code";
            CreateGVData(dataGridView3, sql);

            if (dataGridView3.Columns.Count == 11 && dataGridView3.Rows.Count >0)
            {
                DataGridViewTextBoxColumn ifWin = new DataGridViewTextBoxColumn();
                ifWin.HeaderText = "勝敗";
                //ifWin.CellType = ;
                dataGridView3.Columns.Add(ifWin);
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    if (Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) >= Convert.ToDouble(dataGridView3.Rows[i].Cells[8].Value))
                    {
                        dataGridView3.Rows[i].Cells[11].Value = " ☆ 勝";
                        for (int j = 0; j < dataGridView3.Rows[i].Cells.Count; j++)
                            dataGridView3.Rows[i].Cells[j].Style.BackColor = Color.LightPink;
                    }
                    else
                    {
                        dataGridView3.Rows[i].Cells[11].Value = " 敗";
                        for (int j = 0; j < dataGridView3.Rows[i].Cells.Count; j++)
                            dataGridView3.Rows[i].Cells[j].Style.BackColor = Color.Green;
                    }
                }
            }
            //DataGridViewStyle(dataGridView3);
        }
        //績效追縱
        private void Button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                //檢查stock code
                String stock_code = "";
                if (Regex.IsMatch(textBox1.Text.Split(' ')[0], @"[a-zA-Z0-9]"))
                    stock_code = textBox1.Text.Split(' ')[0];
                else
                    stock_code = textBox1.Text.Split(' ')[1];

                String sql = String.Empty;
                if (stock_list.ContainsKey(stock_code))
                    MessageBox.Show("錯誤!查無此代碼 請重新輸入", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else if (Regex.IsMatch(textBox3.Text, @"^(/d*)[.](/d*)$") || textBox3.Text.Length == 0)
                    MessageBox.Show("錯誤!輸入價格未輸入 or 非數字 請重新輸入", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    MessageBox.Show(String.Format("代碼[{0}] 價格[{1}] 日期[{2}] 格式正確", textBox1.Text, textBox3.Text, DateTime.Now.ToString("yyyy/MM/dd")));
                    sql = String.Format("INSERT OR replace into track_list (stock_code,stock_name,track_price,track_date) values ('{0}','{1}','{2}','{3}')", textBox1.Text, stock_list[textBox1.Text], textBox3.Text, DateTime.Now.ToString("yyyy/MM/dd"));

                    String updateNum = DB_SQL(sql);
                    if (updateNum.All(Char.IsDigit))
                    {
                        if (Convert.ToInt32(updateNum) > 0)
                            MessageBox.Show(String.Format("追踨名單{0} 價格{1}寫入資料庫成功{2}筆", textBox1.Text, textBox3.Text, updateNum));
                        else
                            MessageBox.Show(String.Format("追踨名單{0} 價格{1}寫入資料庫失敗{2}筆", textBox1.Text, textBox3.Text, updateNum));

                    }
                    else
                    {
                        if (updateNum.ToLower().Contains("unique"))
                            MessageBox.Show("資料重覆->更新資料");
                        else
                            MessageBox.Show("[錯誤]" + updateNum);
                    }
                }

                sql = "select t.stock_code,t.stock_name,t.track_price,t.track_date,l.close,l.stock_date,(close-track_price) price_diff, Round((close-track_price)/track_price,2) diff_percent,d.before_price,d.dividen_date from track_list t left outer join latest_stock_price l on t.stock_code = l.stock_code left outer join latest_stock_dividen d on t.stock_code = d.stock_code";
                CreateGVData(dataGridView3, sql);

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    dataGridView3.Rows[i].Cells[11].Value = (Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) >= Convert.ToDouble(dataGridView3.Rows[i].Cells[8].Value)) ? " ☆ 勝" : "敗";
                }

            }
            catch (Exception eee)
            {
                MessageBox.Show(eee.Message);
            }
        }

        private void DataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView3.Columns["btnDel"].Index && e.RowIndex >= 0)
            {
                if (MessageBox.Show(String.Format("是否確定刪除[{0}-{1}]{2}資料?", dataGridView3.Rows[e.RowIndex].Cells[0].Value, dataGridView3.Rows[e.RowIndex].Cells[1].Value, dataGridView3.Rows[e.RowIndex].Cells[3].Value), "刪除追踨資料", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    String sql = String.Format("delete from track_list where stock_code='{0}' and track_date='{1}' ",dataGridView3.Rows[e.RowIndex].Cells[0].Value,dataGridView3.Rows[e.RowIndex].Cells[3].Value);
                    String updateNum = DB_SQL(sql);


                    sql = "select t.stock_code,t.stock_name,t.track_price,t.track_date,l.close,l.stock_date,(close-track_price) price_diff, Round((close-track_price)/track_price,2) diff_percent,d.before_price,d.dividen_date from track_list t left outer join latest_stock_price l on t.stock_code = l.stock_code left outer join latest_stock_dividen d on t.stock_code = d.stock_code";
                    CreateGVData(dataGridView3, sql);
                    if (dataGridView3.Columns.Count == 11 && dataGridView3.Rows.Count > 0)
                    {
                        DataGridViewTextBoxColumn ifWin = new DataGridViewTextBoxColumn();
                        ifWin.HeaderText = "勝敗";
                        //ifWin.CellType = ;
                        dataGridView3.Columns.Add(ifWin);
                        for (int i = 0; i < dataGridView3.Rows.Count; i++)
                        {
                            dataGridView3.Rows[i].Cells[11].Value = (Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) >= Convert.ToDouble(dataGridView3.Rows[i].Cells[8].Value)) ? " ☆ 勝" : "敗";
                        }
                    }
                }
                else
                    MessageBox.Show("未刪除資料! 選擇取消刪除!","取消刪除",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
        }

        private void TrackBar1_Scroll(object sender, EventArgs e)
        {
            //MessageBox.Show(String.Format("目前門檻{0}年",trackBar1.Value));
        }

        private void 大量智慧剖析ToolStripMenuItem_Click(object sender, EventArgs e)  //1135
        {
            tabControl1.SelectedTab = tabPage1;

            label1.Visible = true;
            //button1.Visible = true;
            button9.Visible = true;
            textBox1.Visible = true;

            label2.Visible = false;

            textBox3.Visible = false;
            label3.Visible = false;
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            double increase_max = 0.0;  //歷史填權最大漲幅
            double increase_min = 0.0;  //歷史填權最小漲幅
            double decrease_max = 0.0;  //歷史最大跌幅
            double decrease_min = 0.0;  //勘史最小跌幅

            DialogResult select = MessageBox.Show("此功能將花費較多時間是否確定要重新剖析?", "大量智慧剖析", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (select == DialogResult.No)
            {
                MessageBox.Show("選擇取消重新剖析", "取消重新剖析");
            }
            else
            {
                new Thread(() =>
                {
                    String sql = "";

                    //發放股利股息總共幾年
                    int dividen_times = 0;
                    //股利總和
                    double dividen = 0.0;
                    //取得所有代碼清單
                    String[] stock_list = DB_SQL("select stock_code from stock_profile order by stock_code ", "stock_code").Split(',');
                    //調整成 十年都有發股息的清單
                    //String[] stock_list = DB_SQL("select stock_code from stock_dividen  where  stock_code>='3008' group by stock_code  order by stock_code", "stock_code").Split(',');

                    //股利列表
                    String[] dividen_list = null;

                    //目前年度股利股息
                    String[] year_dividen = null;
                    //前一年度股利股息
                    String[] next_dividen = null;
                    //針對每隻股票去掃描

                    //總交易日
                    int total_days = 0;
                    //有填權息日數
                    int fill_days = 0;
                    //有填權年數
                    int win_times = 0;
                    //勝率
                    double win_rate = 0.0;
                    double avg_win_rate = 0.0;
                    //股票名稱
                    String stock_name = "";
                    String updateNum = "";
                    //區間最大最小股價
                    double min_price = 0.0;
                    double max_price = 0.0;
                    double max_range = 0.0;
                    double min_range = 0.0;


                    //最小填權日數
                    String min_date = "";
                    int least_fill_days = -1;   //日曆天
                    int least_trade_days = -1;  //交易天

                    //計算總權重數
                    double total_weight = 0.0;
                    //先清空所有排名
                    updateMsgNew("清空資料庫所有排名");
                    DB_SQL("delete from money_rank_new");
                    DB_SQL("delete from conclude_data");

                    //最新一年 填息價
                    double latest_fill_price = 0.0;
                    //推估最新一年買入價格 & 停利價格
                    double min_buy_price = 0.0;
                    double min_sell_price = 0.0;

                    //平均最小填息『交易日』數
                    double avg_trade_days = 0.0;

                    //紀錄trackBar值
                    int trackBarValue = 0;
                    this.Invoke(new MethodInvoker(delegate { trackBarValue = trackBar1.Value; }));
                    String result = "";

                    //最小平均獲利
                    double avg_min_profit = 0.0;
                    double avg_max_profit = 0.0;
                    //有計算最小獲利的總數
                    int total_num = 0;

                    //計算總花費時間
                    DateTime t1 = DateTime.Now;
                    for (int i = 0; i < stock_list.Length; i++)
                    {
                        DateTime t2 = DateTime.Now;
                        //sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}' and compare_date >={1}0101", stock_list[i], DateTime.Now.AddYears(-trackBarValue).ToString("yyyy"));
                        
                        //回溯測試
                        //sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}'  and compare_date <20190101", stock_list[i]);
                        
                        //計算12年資料
                        sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}'", stock_list[i]);
                        String stock_code = stock_list[i];

                        dividen_times = Convert.ToInt32(DB_SQL(sql, "cnt"));
                        //每次重算
                        dividen = 0.0;
                        win_rate = 0.0;
                        win_times = 0;
                        avg_win_rate = 0.0;

                        //☆☆for debug
                        //if(stock_code=="2330")
                        try
                        {
                            //String stock_code = stock_list[i];
                            stock_name = DB_SQL("select stock_name from stock_profile where stock_code='" + stock_code + "'", "stock_name");
                            //先過濾掉無資料的股票
                            if (Convert.ToInt32(DB_SQL("select count(*) cnt from stock_price_new where stock_code='" + stock_code + "'", "cnt")) == 0)
                                updateMsgNew(String.Format("[{0}]{1} 無股價資料", stock_code, stock_name));
                            else
                            {
                                updateMsgNew(String.Format("==============================================================={0}[{1}]{2}",Environment.NewLine,stock_code,stock_name));
                                if (dividen_times > 0)
                                {
                                    total_num++;
                                    for (int j = 0; j < dividen_times; j++)
                                    {
                                        
                                        sql = "select before_price,after_price,dividen,dividen_date,compare_date,least_days from stock_dividen where stock_code='" + stock_code + "' order by dividen_date desc";
                                        dividen_list = DB_SQL(sql, "dividen_date,dividen,before_price,after_price,compare_date,least_days").Split(';');

                                        if (j == 0) //最新一年
                                        {
                                            year_dividen = dividen_list[j].Split('@');
                                            //updateMsg(String.Format("[{0}]{1}", stock_list[i], stock_name));
                                            updateMsgNew(String.Format("[{0}~現在]股息{1}元 填息價{2}元 除權價{3}元", year_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));

                                            //紀錄最新一年填息價
                                            latest_fill_price = Convert.ToDouble(year_dividen[2]);

                                            //計算總交易日
                                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date >= {1} ", stock_code, year_dividen[4]);
                                            total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                            //計算有填權息日數
                                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date >= {1} and close >= {2}", stock_code, year_dividen[4], year_dividen[2]);
                                            fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));


                                            win_rate = (double)fill_days / total_days;

                                            //取得區間最大、最小股價
                                            sql = String.Format("select max(close) max_price,min(close) min_price from stock_price_new where stock_code='{0}' and compare_date >= {1} ", stock_code, year_dividen[4]);
                                            max_price = Convert.ToDouble(DB_SQL(sql, "max_price"));
                                            min_price = Convert.ToDouble(DB_SQL(sql, "min_price"));
                                            //計算最大漲跌幅
                                            max_range = (Math.Round((max_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                                            min_range = (Math.Round((min_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                                            increase_max = max_range;
                                            increase_min = max_range;
                                            decrease_max = min_range;
                                            decrease_min = min_range;

                                            //計算是否填權息
                                            if (fill_days > 0)
                                            {
                                                win_times++;
                                                sql = String.Format("select min(stock_date) min_date from stock_price_new where stock_code='{0}' and compare_date >= {1} and close >= {2}", stock_code, year_dividen[4], year_dividen[2]);
                                                min_date = DB_SQL(sql, "min_date");
                                                least_fill_days = (Convert.ToDateTime(min_date).Date - Convert.ToDateTime(year_dividen[0]).Date).Days;
                                                sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} ", stock_code, year_dividen[4], min_date.Replace("/", ""));

                                                //updateMsgNew(sql);

                                                least_trade_days = Convert.ToInt32(DB_SQL(sql,"cnt"));
                                                avg_trade_days = least_trade_days;
                                            }
                                            updateMsgNew(String.Format("總交易天數{0}天 超過填息價總計{1}天{8}=> 勝率 {2} % 最高股價{3}元(填息後最大漲幅{4}%) 最低股價{5}元(除息後最大跌幅{6}%){8}    程式計算{7} [{10}~{11}]{8}vs GoodInfo計算最小天數{9}天{8}"
                                                , total_days, fill_days, (win_rate * 100).ToString("F1"), max_price, max_range, min_price, min_range, (fill_days >0) ?  " 填息最少交易"+least_trade_days+"天(日曆" + least_fill_days + "天)":" 沒填息" , Environment.NewLine, year_dividen[5], year_dividen[4], min_date));

                                            //計算股利合
                                            dividen = dividen + Convert.ToDouble(year_dividen[1]);
                                            
                                            //最新一年權重=1
                                            total_weight = 1;
                                            //計算勝率 最新一年權重 =  1
                                            avg_win_rate = avg_win_rate + win_rate*1;
                                        }
                                        else
                                        {
                                            year_dividen = dividen_list[j].Split('@');
                                            next_dividen = dividen_list[j - 1].Split('@');
                                            updateMsgNew(String.Format("[{0}~{1}]股息{2}元 填息價{3}元 除權價{4}元", year_dividen[0], next_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));
                                            //計算總交易日
                                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], next_dividen[4]);
                                            total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                            //計算有填權息日數
                                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} and close >= {3}", stock_code, year_dividen[4], next_dividen[4], year_dividen[2]);
                                            fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                            //☆☆debug
                                            if (total_days == 0)
                                                win_rate = 0.0;
                                            else
                                                win_rate = (double)fill_days / total_days;

                                            //取得區間最大、最小股價
                                            sql = String.Format("select max(close) max_price,min(close) min_price from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], next_dividen[4]);
                                            max_price = Convert.ToDouble(DB_SQL(sql, "max_price"));
                                            min_price = Convert.ToDouble(DB_SQL(sql, "min_price"));

                                            //計算最大漲跌幅
                                            max_range = (Math.Round((max_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                                            min_range = (Math.Round((min_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                                            if (max_range > increase_max) increase_max = max_range;
                                            if (max_range < increase_min) increase_min = max_range;
                                            if (min_range > decrease_min) decrease_min = min_range;
                                            if (min_range < decrease_max) decrease_max = min_range;

                                            //計算是否填權息
                                            if (fill_days > 0)
                                            {
                                                win_times++;
                                                //找出收盤價>填息價格的「日期」
                                                sql = String.Format("select min(stock_date) min_date from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} and close >= {3}", stock_code, year_dividen[4], next_dividen[4], year_dividen[2]);
                                                min_date = DB_SQL(sql, "min_date");

                                                //計算兩區間之間 填息最小交易日數(分子)
                                                least_fill_days = (Convert.ToDateTime(min_date).Date - Convert.ToDateTime(year_dividen[0]).Date).Days;
                                                sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], min_date.Replace("/", ""), year_dividen[2]);
                                                updateMsgNew(sql);
                                                //計算兩區間之間 總交易日數 (分母)
                                                least_trade_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                                avg_trade_days = avg_trade_days + least_trade_days;
                                            }
                                            //else //未填息的情況下 
                                            //{
                                            //    //最小填息交易日數設成-1
                                            //    least_fill_days = -1;
                                            //    min_date = "";
                                            //}
                                            updateMsgNew(String.Format("總交易天數{0}天 超過填息價總計{1}天{8}=> 勝率 {2} % 最高股價{3}元(填息後最大漲幅{4}%) 最低股價{5}元(除息後最大跌幅{6}%){8}    程式計算{7} [{10}~{11}]{8}vs GoodInfo計算最小天數{9}天{8}"
                                                 , total_days, fill_days, (win_rate * 100).ToString("F1"), max_price, max_range, min_price, min_range, (fill_days>0) ? " 填息最少交易" + least_trade_days + "天(日曆" + least_fill_days + "天)": " 沒填息" , Environment.NewLine, year_dividen[5], year_dividen[4], min_date));

                                            //計算股利合
                                            dividen = dividen + Convert.ToDouble(year_dividen[1]);
                                            //每年的權重 依序為  0.9 ~ 0.8 ... 0.1
                                            double year_weight = 1 - 0.1 * j;
                                            total_weight = total_weight + year_weight;
                                            //計算勝率 採用加權比重平均
                                            avg_win_rate = avg_win_rate + win_rate * (year_weight);
                                            String last_price = DB_SQL(String.Format("select close from stock_price_new where stock_code = '{0}' and stock_date<'{1}' order by stock_date desc limit 0,1",stock_code, next_dividen[0]),"close");
                                            
                                            //☆☆ 寫入conclude_data 作為AI迴歸測試用資料
                                            double win_rate_percent = Math.Round(win_rate * 100, 2,MidpointRounding.AwayFromZero);
                                            double WADAR = Math.Round(((avg_win_rate / total_weight) * 100),2, MidpointRounding.AwayFromZero);

                                            updateMsgNew(String.Format("區間勝率{0} 目前權重{1} WADAR{2}", win_rate, total_weight,WADAR));


                                            sql = String.Format("insert into conclude_data (stock_code,stock_name,fill_days,total_days,win_rate         ,fill_dividen        , start_date     ,end_date        ,dividen          ,max_rise ,max_drop ,ai_max_rise,ai_max_drop,last_price,wadar) values ('{0}','{1}',{2},{3},{4},{5},'{6}','{7}',{8},{9},{10},{11},{12},{13},{14})",
                                                                                            stock_code,stock_name,fill_days,total_days, win_rate_percent, (fill_days > 0)?1:0, year_dividen[0], next_dividen[0], year_dividen[1], max_range,min_range,-1         ,-1         ,last_price,WADAR);
                                            updateNum = DB_SQL(sql.ToString());
                                            if (updateNum.All(Char.IsDigit))
                                            {
                                                if (Convert.ToInt32(updateNum) > 0)
                                                    updateMsgNew("寫入AI測試資料成功(" + updateNum.ToString() + ")");
                                                else
                                                    updateMsgNew("寫入AI測試資料失敗(" + updateNum.ToString() + ")");
                                            }
                                            else
                                                updateMsgNew("寫入AI測試資料失敗(" + sql + ")");
                                        }
                                    }

                                }
                                
                            }//end of else
                            //計算 哇達 勝率
                            String WADAR_win_rate = ((avg_win_rate / total_weight) * 100).ToString("F1");

                            //計算 填息平均最小天數
                            if (win_times > 0)//填息次數至少要大於1次
                                avg_trade_days = Math.Round(avg_trade_days / (double)win_times, 2);
                            else//從來沒有填息過 設成-1
                                avg_trade_days = -1;
                            //計算 單一筆 花費時間
                            TimeSpan ts2 = DateTime.Now - t2;
                            min_buy_price = Math.Round(latest_fill_price * (100+decrease_min)/100,2);
                            min_sell_price = Math.Round(latest_fill_price * (100+increase_min)/100,2);
                            
                            updateMsgNew(String.Format("[{0}]{1}{6}十年發 {2}次股利(總合{3}元)填權息{4}次 ☆WADAR指標={5}% {6}填息後最大漲幅介於{7}%~{8}% 除息後最大跌幅介於{9}%~{10}%",
                                      stock_code, stock_name, dividen_times, dividen, win_times, WADAR_win_rate, Environment.NewLine, increase_min, increase_max, decrease_min, decrease_max));
                            updateMsgNew(String.Format("★ 最新填息價 {0} 元,推估至少低於{1}元({2}%)買入(買進點) 高於至少{3}元({4}%)賣出(停利點)", latest_fill_price, min_buy_price,decrease_min,min_sell_price,increase_min));
                            updateMsgNew(String.Format("填息最小交易日數平均:{0}天 計算總計花費時間{1}秒", avg_trade_days,ts2.TotalSeconds));
                            updateMsgNew("===============================================================");
                            //    Thread.Sleep(200000);
                            //計算最小獲利
                            avg_min_profit = avg_min_profit + decrease_min;
                            avg_max_profit = avg_max_profit + decrease_max;

                            sql = String.Format("insert into money_rank_new (stock_code,stock_name,win_times,dividen_times,dividen,avg_win_rate,cal_date,least_days) values ('{0}','{1}',{2},{3},{4},{5},'{6}',{7})", stock_list[i], stock_name, win_times, dividen_times, dividen, WADAR_win_rate, DateTime.Now.ToString("yyyy/MM/dd"), avg_trade_days);

                            updateNum = DB_SQL(sql.ToString());
                            if (updateNum.All(Char.IsDigit))
                            {
                                if (Convert.ToInt32(updateNum) > 0)
                                    updateMsgNew("寫入排名資料成功(" + updateNum.ToString() + ")");
                                else
                                    updateMsgNew("寫入排名資料失敗(" + updateNum.ToString() + ")");
                            }
                            else
                                updateMsgNew("寫入排名資料失敗(" + updateNum + ")");

                        }
                        catch (Exception exx)
                        {
                            updateMsgNew(exx.Message);
                        }
                    }
                    TimeSpan ts1 = DateTime.Now - t1;
                    updateMsgNew(String.Format("統計：計算{0}隻股票 總花費時間{1}秒", stock_list.Length,ts1.TotalSeconds));

                    avg_min_profit = avg_min_profit / total_num;
                    avg_max_profit = avg_max_profit / total_num;
                    updateMsgNew(String.Format("平均最小獲利={0},num={1},平均最大獲利={2}", avg_min_profit, total_num, avg_max_profit));

                    this.Invoke((MethodInvoker)delegate
                    {
                        label3.Text = String.Format("說明：挑出10年內(1)連續{0}年發股息 (2)全部填權息 ☆(3)WADAR指標前五十名 『總排行榜』", trackBarValue);
                        label3.Visible = true;
                    });

                    //更新GridView
                    sql = " select stock_code,stock_name,win_times,dividen_times,avg_win_rate,dividen from money_rank_new " +
                        " where dividen_times >= " + trackBarValue +
                        " order by (win_times/dividen_times) desc ,dividen_times desc,avg_win_rate desc limit 0,50";
                    result = DB_SQL(sql, "stock_code,stock_name,win_times,dividen_times,avg_win_rate,dividen");

                    CreateGVData(dataGridView2, sql);

                    this.Invoke((MethodInvoker)delegate
                    {
                        label3.Text = "說明：挑出10年內(1)至少連續" + trackBarValue + "年發股息 (2)全部填權息 ☆(3)WADAR指標前五十名 (4)尚未填息 的『可買進名單』";
                        label3.Visible = true;
                    });

                    //更新GridView
                    sql = "select m.stock_code,m.stock_name,m.win_times,m.dividen_times,m.avg_win_rate,p.close,d.before_price, " +
                               "  Round((d.before_price - p.close),2) price_diff, Round(100*(d.before_price - p.close) / before_price,3) price_percent,d.dividen_date,p.stock_date " +
                               "       from money_rank_new m left outer join latest_stock_price p " +
                               "      on m.stock_code = p.stock_code " +
                               "   left outer join latest_stock_dividen d on d.stock_code = m.stock_code " +
                               "   where m.dividen_times >=" + trackBarValue + " and p.close < d.before_price and m.win_times = m.dividen_times " +
                               "   order by(price_diff/ before_price) desc , m.avg_win_rate desc limit 0,50";
                    result = DB_SQL(sql, "stock_code,stock_name,win_times,dividen_times,avg_win_rate,close,before_price,price_diff,price_percent,dividen_date,stock_date");

                    CreateGVData(dataGridView1, sql);
                }).Start();
            }
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void 價差排行榜ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //show track bar
            trackBar1.Visible = true;
            label4.Visible = true;
            label5.Visible = true;
            button3.Visible = true;

            ProcessStartInfo Info = new ProcessStartInfo();
            Info.FileName = "dayfa.vbs";//執行的檔案名稱
            //Info.WorkingDirectory = @"E:\NVR";//檔案所在的目錄
            Process.Start(Info);// RUN bat

            MessageBox.Show("必須先更新最新股價資訊", "發財需知");
            tabControl1.SelectedTab = tabPage3;
            label3.Text = "說明：挑出10年內(1)連續10年發股息 (2)填息次數比例最高 ☆(3)WADAR指標前二十名 (4)尚未填息 的『可買進名單』";
            label3.Visible = true;
            new Thread(() =>
            {
                /*
                    dataGridView1.Columns[0].HeaderText = "代碼";
                    dataGridView1.Columns[1].HeaderText = "名稱";
                    dataGridView1.Columns[2].HeaderText = "填息次數";
                    dataGridView1.Columns[3].HeaderText = "發息次數";
                    dataGridView1.Columns[4].HeaderText = "☆WADAR指標";
                    dataGridView1.Columns[5].HeaderText = "最新收盤價";
                    dataGridView1.Columns[6].HeaderText = "填權息價格";
                    dataGridView1.Columns[7].HeaderText = "填權息價差";
                    dataGridView1.Columns[8].HeaderText = "填權息價差%";
                    dataGridView1.Columns[9].HeaderText = "最新除權日期";
                    dataGridView1.Columns[10].HeaderText = "最新收盤日期";
                 */
                String sql = @"
                            select distinct m.stock_code,m.stock_name,m.win_times,m.dividen_times,m.avg_win_rate,p.close,d.before_price 
                            ,Round((d.before_price - p.close),2) price_diff, Round(100*(d.before_price - p.close) / before_price,2) price_percent,d.dividen_date,p.stock_date 
                            from (select distinct stock_code,stock_name,win_times,dividen_times,avg_win_rate
                            from money_rank_new m where dividen_times >=12
                            order by (win_times/dividen_times) desc,dividen_times desc,avg_win_rate desc
                            limit 0,20) m 
                            left outer join latest_stock_price p on m.stock_code = p.stock_code
                            left outer join latest_stock_dividen d on d.stock_code = m.stock_code
                            where  p.close < d.before_price and m.win_times = m.dividen_times
                            order by(price_diff/ before_price) desc , m.avg_win_rate
                             ";
                String result = DB_SQL(sql, "stock_code,stock_name,win_times,dividen_times,avg_win_rate,close,before_price,price_diff,price_percent,dividen_date,stock_date");

                CreateGVData(dataGridView1, sql);

            }).Start();
        }

        private void 最小天數排行榜ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //show track bar
            trackBar1.Visible = true;
            label4.Visible = true;
            label5.Visible = true;
            button3.Visible = true;
            tabControl1.SelectedTab = tabPage3;

            ProcessStartInfo Info = new ProcessStartInfo();
            Info.FileName = "dayfa.vbs";//執行的檔案名稱
            //Info.WorkingDirectory = @"E:\NVR";//檔案所在的目錄
            Process.Start(Info);// RUN bat

            MessageBox.Show("必須先更新最新股價資訊", "發財需知");
            tabControl1.SelectedTab = tabPage3;
            label3.Text = "說明：挑出12年內(1)連續12年發股息 (2)填息次數比例最高 ☆(3)WADAR指標前二十名 (4)尚未填息 的『可買進名單』";
            label3.Visible = true;
            new Thread(() =>
            {
                /*
                    dataGridView1.Columns[0].HeaderText = "代碼";
                    dataGridView1.Columns[1].HeaderText = "名稱";
                    dataGridView1.Columns[2].HeaderText = "填息次數";
                    dataGridView1.Columns[3].HeaderText = "發息次數";
                    dataGridView1.Columns[4].HeaderText = "☆WADAR指標";
                    dataGridView1.Columns[5].HeaderText = "最新收盤價";
                    dataGridView1.Columns[6].HeaderText = "填權息價格";
                    dataGridView1.Columns[7].HeaderText = "填權息價差";
                    dataGridView1.Columns[8].HeaderText = "填權息價差%";
                    dataGridView1.Columns[9].HeaderText = "最新除權日期";
                    dataGridView1.Columns[10].HeaderText = "最新收盤日期";
                 */
                String sql = @"
                                select distinct m.stock_code,m.stock_name,m.win_times,m.dividen_times,m.avg_win_rate,p.close,d.before_price 
                                ,Round((d.before_price - p.close),2) price_diff, Round(100*(d.before_price - p.close) / before_price,2) price_percent,d.dividen_date,p.stock_date,m.least_days
                                from (select distinct stock_code,stock_name,win_times,dividen_times,avg_win_rate,least_days
                                from money_rank_new where dividen_times >=12
                                order by (win_times/dividen_times) desc,dividen_times desc,avg_win_rate desc
                                limit 0,20) m 
                                left outer join latest_stock_price p on m.stock_code = p.stock_code
                                left outer join latest_stock_dividen d on d.stock_code = m.stock_code
                                where  p.close < d.before_price and m.win_times = m.dividen_times 
                                order by least_days
                            ";
                String result = DB_SQL(sql, "stock_code,stock_name,win_times,dividen_times,avg_win_rate,close,before_price,price_diff,price_percent,dividen_date,stock_date");

                CreateGVData(dataGridView1, sql);

            }).Start();
        }

        //static readonly string _trainDataPath = Path.Combine(Environment.CurrentDirectory, "Data", "taxi-fare-train.csv");
        static readonly string _trainDataPath = Path.Combine(Environment.CurrentDirectory, "Data", "train.csv");

        static readonly string _modelPath = Path.Combine(Environment.CurrentDirectory, "Data", "Model.zip");

        private void aI迴歸預測分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage2;
            label1.Visible = true;
            //button1.Visible = true;
            button9.Visible = true;
            textBox1.Visible = true;

            label2.Visible = false;
            textBox3.Visible = false;
            label3.Visible = false;
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            updateMsg("AI迴歸分析預測");

            //Console.WriteLine("從資料庫匯出訓練資料(Training Data) CSV檔");
            //String sql = "select * from conclude_data ";
            //String raw_data = DB_SQL(sql, "stock_code,dividen,last_price,win_rate,fill_dividen,max_rise,max_drop");
            ////write data to csv
            //using (StreamWriter out_file = new StreamWriter("./Data/train.csv", false))
            //{
            //    out_file.WriteLine("stock_code,dividen,last_price,win_rate,fill_dividen,max_rise,max_drop");
            //    raw_data = raw_data.Replace(";", Environment.NewLine).Replace("@", ",").Replace("-", "");
            //    out_file.Write(raw_data);
            //}
            MLContext mlContext = new MLContext(seed: 0);
            updateMsg("載入特徵值並轉換資料中…");
            //updateMsg("載入中…");
            updateMsg("載入特徵[stock_code]");
            updateMsg("載入特徵[dividen]");
            updateMsg("載入特徵[last_price]");
            updateMsg("載入特徵[win_rate]");
            updateMsg("載入特徵[fill_dividen]");
            updateMsg("載入特徵[max_rise]");
            updateMsg("載入特徵[max_drop]");
            DateTime t1 = DateTime.Now;
            updateMsg("開始載入訓練資料->迴歸模型");
            var model = Train(mlContext, _trainDataPath);
            TimeSpan ts = DateTime.Now - t1;
            updateMsg(String.Format("建立迴歸模型完成 費時{0}秒", Math.Round(ts.TotalSeconds)));
            
            updateMsg("利用載入資料建立「預測模型引擎」");
            t1 = DateTime.Now;
            //利用載入資料建立「預測模型引擎」
            var predict_test = mlContext.Model.CreatePredictionEngine<MaxRiseDrop, MaxRiseDropPrdeiction>(model);

            //大量測試 以10年台積電資料來測試
            String sql = "select * from conclude_data where stock_code='2330' limit 0,10";
            String raw_data = DB_SQL(sql, "stock_code,dividen,last_price,win_rate,fill_dividen,max_rise,max_drop");
            //每一行的資料代表 1個時間區間內的測試資料
            String[] rows = raw_data.Split(';');

            double sum_errors = 0.0;
            double error = 0.0;
            for (int i = 0; i < rows.Length-1; i++)
            {
                String[] one_row = rows[i].Split('@');
                updateMsg(String.Format("測資代碼[{0}] 股利[{1}] 除息前價格[{2}] 單一區間WADAR值[{3}] 是否填息[{4}] 最大漲幅[{5}] 最大跌幅[{6}]",
                                        one_row[0], one_row[1], one_row[2], one_row[3], one_row[4], one_row[5], one_row[6]));
                //var testSample = new MaxRiseDrop()
                //{
                //    stock_code = "2330",
                //    dividen = 0.7F,
                //    last_price = 92.15F,
                //    win_rate = 0.81F,
                //    fill_dividen = 1F,
                //    max_rise = 19.2F
                //    ,
                //    max_drop = 0    //用來預測真實的數值
                //};
                var testSample = new MaxRiseDrop()
                {
                    stock_code = one_row[0],
                    dividen = Convert.ToSingle(one_row[1]),
                    last_price = Convert.ToSingle(one_row[2]),
                    win_rate = Convert.ToSingle(one_row[3]),
                    fill_dividen = Convert.ToSingle(one_row[4]),
                    max_rise = Convert.ToSingle(one_row[5])
                    ,max_drop = 0    //用來預測真實的數值
                };
                //單次測試
                updateMsg("開始單次預測結果");
                var prediction = predict_test.Predict(testSample);
                double true_value = Convert.ToDouble(one_row[6]);
                ts = DateTime.Now - t1;
                //計算單次錯誤率
                error = Math.Round(100*(Math.Abs(Math.Abs(true_value) - Math.Abs(prediction.max_drop)) / Math.Abs(true_value)),2);
                sum_errors = sum_errors + error;
                updateMsg(String.Format("預測下跌:{0}%   vs 真實數值:{1}%  錯誤率:{2}%", prediction.max_drop, true_value, error));
                updateMsg(String.Format("測試迴歸模型1次 費時{0}秒", Math.Round(ts.TotalSeconds, 2)));
            }
            sum_errors = Math.Round(sum_errors / rows.Length - 1, 2);
            updateMsg(String.Format("總計測試{0}次 錯誤率:{1}%", rows.Length - 1, sum_errors));
        }
        //☆遇到測試資料不是數字-> 需透過 Categorical.OneHotEncoding 將本來欄位轉換成 『數值』並取一個新的名字再放入訓練

        public static ITransformer Train(MLContext mlContext, string dataPath)
        {
            IDataView dataView = mlContext.Data.LoadFromTextFile<MaxRiseDrop>(dataPath, hasHeader: true, separatorChar: ',');
            var pipeline = mlContext.Transforms.CopyColumns(outputColumnName: "Label", inputColumnName: "max_drop")
                    .Append(mlContext.Transforms.Categorical.OneHotEncoding(outputColumnName: "stock_code_encode", inputColumnName: "stock_code"))
                    //選擇輸入的欄位 
                    //.Append(mlContext.Transforms.Concatenate("Features", "stock_code_encode", "dividen", "last_price"))
                    //.Append(mlContext.Transforms.Concatenate("Features", "stock_code_encode", "dividen", "last_price", "win_rate"))
                    //.Append(mlContext.Transforms.Concatenate("Features", "stock_code_encode", "dividen", "last_price", "win_rate", "fill_dividen"))
                    //.Append(mlContext.Transforms.Concatenate("Features", "stock_code_encode", "dividen", "last_price", "win_rate", "fill_dividen", "max_rise"))
                    .Append(mlContext.Transforms.Concatenate("Features", "stock_code_encode", "dividen", "last_price", "win_rate", "fill_dividen", "max_rise","max_drop"))

                    //選擇演算法
                    //.Append(mlContext.Regression.Trainers.FastForest());
                    //.Append(mlContext.Regression.Trainers.FastTree());
                    //.Append(mlContext.Regression.Trainers.FastTreeTweedie());
                    .Append(mlContext.Regression.Trainers.LbfgsPoissonRegression());
                    //.Append(mlContext.Regression.Trainers.OnlineGradientDescent());

            var model = pipeline.Fit(dataView);
            return model;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            int n;
            if (int.TryParse(textBox5.Text, out n))
            {
                
                String stock_code = label7.Text.Split(' ')[0];
                //MessageBox.Show(stock_code, textBox1.Text.Split(' ')[1]);

                //錯誤寫法 ： String sql = "select max(close) max_price,min(close) min_price from stock_price_new where stock_code = '" + stock_code + "' order by compare_date desc limit 0,90";
                String sql = "select max(close) max_price, min(close) min_price from(select* from stock_price_new where stock_code = '" + stock_code + "'  limit 0,"+textBox5.Text+")";
                //取得最大價格、最低價格
                String max_min = DB_SQL(sql, "max_price,min_price");
                //MessageBox.Show(max_min);
                //設定最大值及最小值
                chart1.ChartAreas["ChartArea1"].AxisY.Maximum = Convert.ToDouble(max_min.Split('@')[0]) * 1.03;
                chart1.ChartAreas["ChartArea1"].AxisY.Minimum = Convert.ToDouble(max_min.Split('@')[1]) * 0.97;

                clickedNode = "";

                tabControl1.SelectedTab = tabPage6;
                chart1.ChartAreas["ChartArea1"].AxisX.MajorGrid.LineWidth = 1;
                chart1.ChartAreas["ChartArea1"].AxisY.MajorGrid.LineWidth = 1;

                //chart1.ChartAreas["ChartArea1"].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.WordWrap;
                //chart1.ChartAreas["ChartArea1"].AxisX.IsLabelAutoFit = true;
                //chart1.ChartAreas["ChartArea1"].AxisX.LabelStyle.Enabled = true;

                chart1.Series[0].XValueMember = "stock_date";
                chart1.Series[0].YValueMembers = "high,low,open,close";//順序不可以改!!!!!
                chart1.Series[0].XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Date;
                chart1.Series[0].CustomProperties = "PriceDownColor=Green,PriceUpColor=Red";

                chart1.Series[0]["OpenCloseStyle"] = "Triangle";
                chart1.Series[0]["ShowOpenClose"] = "Both";
                chart1.DataManipulator.IsStartFromFirst = true;

                //取得最近九十個交易日的資料
                sql = "select * from stock_price_new where stock_code='" + stock_code + "' order by compare_date desc limit 0,2000 ";
                //灌入資料時發現時間日期順序不對 -> 調整sql語法
                sql = "select * from (select * from stock_price_new where stock_code='" + stock_code + "' order by compare_date desc limit 0,90 ) order by compare_date asc";

                DataTable newTable = new DataTable();
                //table = newTable;
                using (SQLiteConnection conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, conn))
                    {
                        using (SQLiteDataReader dr = cmd.ExecuteReader())
                        {
                            newTable.Load(dr);
                        }
                    }
                }

                chart1.DataSource = newTable;
                chart1.DataBind();
            }
            else
                MessageBox.Show("[錯誤] 輸入天數非整數!");
        }

        private void 智慧歷史回溯ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage1;

            label1.Visible = true;
            //button1.Visible = true;
            button9.Visible = true;
            textBox1.Visible = true;

            label2.Visible = false;

            textBox3.Visible = false;
            label3.Visible = false;
            //hide track bar
            trackBar1.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            button3.Visible = false;

            double increase_max = 0.0;  //歷史填權最大漲幅
            double increase_min = 0.0;  //歷史填權最小漲幅
            double decrease_max = 0.0;  //歷史最大跌幅
            double decrease_min = 0.0;  //勘史最小跌幅

            DialogResult select = MessageBox.Show("此功能將花費較多時間是否確定要重新剖析?", "大量智慧剖析", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (select == DialogResult.No)
            {
                MessageBox.Show("選擇取消重新剖析", "取消重新剖析");
            }
            else
            {
                new Thread(() =>
                {
                    String sql = "";

                    //發放股利股息總共幾年
                    int dividen_times = 0;
                    //股利總和
                    double dividen = 0.0;
                    //取得所有代碼清單
                    String[] stock_list = DB_SQL("select stock_code from stock_profile order by stock_code ", "stock_code").Split(',');
                    //調整成 十年都有發股息的清單
                    //String[] stock_list = DB_SQL("select stock_code from stock_dividen  where  stock_code>='3008' group by stock_code  order by stock_code", "stock_code").Split(',');

                    //股利列表
                    String[] dividen_list = null;

                    //目前年度股利股息
                    String[] year_dividen = null;
                    //前一年度股利股息
                    String[] next_dividen = null;
                    //針對每隻股票去掃描

                    //總交易日
                    int total_days = 0;
                    //有填權息日數
                    int fill_days = 0;
                    //有填權年數
                    int win_times = 0;
                    //勝率
                    double win_rate = 0.0;
                    double avg_win_rate = 0.0;
                    //股票名稱
                    String stock_name = "";
                    String updateNum = "";
                    //區間最大最小股價
                    double min_price = 0.0;
                    double max_price = 0.0;
                    double max_range = 0.0;
                    double min_range = 0.0;


                    //最小填權日數
                    String min_date = "";
                    int least_fill_days = -1;   //日曆天
                    int least_trade_days = -1;  //交易天

                    //計算總權重數
                    double total_weight = 0.0;
                    //先清空所有排名
                    updateMsgNew("清空資料庫所有排名");
                    //DB_SQL("delete from money_rank");
                    DB_SQL("delete from conclude_data");

                    //最新一年 填息價
                    double latest_fill_price = 0.0;
                    //推估最新一年買入價格 & 停利價格
                    double min_buy_price = 0.0;
                    double min_sell_price = 0.0;

                    //平均最小填息『交易日』數
                    double avg_trade_days = 0.0;

                    //計算總花費時間
                    DateTime t1 = DateTime.Now;
                    for (int i = 0; i < stock_list.Length; i++)
                    {
                        DateTime t2 = DateTime.Now;
                        //sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}' and compare_date >={1}0101", stock_list[i], DateTime.Now.AddYears(-trackBarValue).ToString("yyyy"));

                        //回溯測試
                        //sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}'  and compare_date <20190101", stock_list[i]);

                        //計算12年資料
                        sql = String.Format("select count(*) cnt from stock_dividen where stock_code='{0}'", stock_list[i]);
                        String stock_code = stock_list[i];

                        dividen_times = Convert.ToInt32(DB_SQL(sql, "cnt"));
                        //每次重算
                        dividen = 0.0;
                        win_rate = 0.0;
                        win_times = 0;
                        avg_win_rate = 0.0;
                        total_weight = 0.0;

                        //☆☆for debug
                        //if(stock_code=="2330")
                        try
                        {
                            //String stock_code = stock_list[i];
                            stock_name = DB_SQL("select stock_name from stock_profile where stock_code='" + stock_code + "'", "stock_name");
                            //先過濾掉無資料的股票
                            if (Convert.ToInt32(DB_SQL("select count(*) cnt from stock_price_new where stock_code='" + stock_code + "'", "cnt")) == 0)
                                updateMsgNew(String.Format("[{0}]{1} 無股價資料", stock_code, stock_name));
                            else
                            {
                                updateMsgNew(String.Format("==============================================================={0}[{1}]{2}", Environment.NewLine, stock_code, stock_name));
                                if (dividen_times > 0)
                                {
                                    for (int j = 0; j < dividen_times; j++)
                                    {

                                        //sql = "select before_price,after_price,dividen,dividen_date,compare_date,least_days from stock_dividen where stock_code='" + stock_code + "' order by dividen_date desc";
                                        //☆逆序排列
                                        sql = "select before_price,after_price,dividen,dividen_date,compare_date,least_days from stock_dividen where stock_code='" + stock_code + "' order by dividen_date asc";

                                        dividen_list = DB_SQL(sql, "dividen_date,dividen,before_price,after_price,compare_date,least_days").Split(';');

                                        if (j == dividen_times-1) //最新一年
                                        {
                                            year_dividen = dividen_list[j].Split('@');
                                            //updateMsg(String.Format("[{0}]{1}", stock_list[i], stock_name));
                                            updateMsgNew(String.Format("[{0}~現在]股息{1}元 填息價{2}元 除權價{3}元", year_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));

                                            //紀錄最新一年填息價
                                            latest_fill_price = Convert.ToDouble(year_dividen[2]);

                                            //計算總交易日
                                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date >= {1} ", stock_code, year_dividen[4]);
                                            total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                            //計算有填權息日數
                                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date >= {1} and close >= {2}", stock_code, year_dividen[4], year_dividen[2]);
                                            fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));


                                            win_rate = (double)fill_days / total_days;

                                            //取得區間最大、最小股價
                                            sql = String.Format("select max(close) max_price,min(close) min_price from stock_price_new where stock_code='{0}' and compare_date >= {1} ", stock_code, year_dividen[4]);
                                            max_price = Convert.ToDouble(DB_SQL(sql, "max_price"));
                                            min_price = Convert.ToDouble(DB_SQL(sql, "min_price"));
                                            //計算最大漲跌幅
                                            max_range = (Math.Round((max_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                                            min_range = (Math.Round((min_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;

                                            //increase_max = max_range;
                                            //increase_min = max_range;
                                            //decrease_max = min_range;
                                            //decrease_min = min_range;
                                            if (max_range > increase_max) increase_max = max_range;
                                            if (max_range < increase_min) increase_min = max_range;
                                            if (min_range > decrease_min) decrease_min = min_range;
                                            if (min_range < decrease_max) decrease_max = min_range;

                                            //計算是否填權息
                                            if (fill_days > 0)
                                            {
                                                win_times++;
                                                sql = String.Format("select min(stock_date) min_date from stock_price_new where stock_code='{0}' and compare_date >= {1} and close >= {2}", stock_code, year_dividen[4], year_dividen[2]);
                                                min_date = DB_SQL(sql, "min_date");
                                                least_fill_days = (Convert.ToDateTime(min_date).Date - Convert.ToDateTime(year_dividen[0]).Date).Days;
                                                sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} ", stock_code, year_dividen[4], min_date.Replace("/", ""));

                                                //updateMsgNew(sql);

                                                least_trade_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                                //avg_trade_days = least_trade_days;
                                                //☆逆序調整
                                                avg_trade_days = avg_trade_days + least_trade_days;

                                                }
                                                updateMsgNew(String.Format("總交易天數{0}天 超過填息價總計{1}天{8}=> 勝率 {2} % 最高股價{3}元(填息後最大漲幅{4}%) 最低股價{5}元(除息後最大跌幅{6}%){8}    程式計算{7} [{10}~{11}]{8}vs GoodInfo計算最小天數{9}天{8}"
                                                , total_days, fill_days, (win_rate * 100).ToString("F1"), max_price, max_range, min_price, min_range, (fill_days > 0) ? " 填息最少交易" + least_trade_days + "天(日曆" + least_fill_days + "天)" : " 沒填息", Environment.NewLine, year_dividen[5], year_dividen[4], min_date));

                                            //計算股利合
                                            dividen = dividen + Convert.ToDouble(year_dividen[1]);

                                            ////最新一年權重=1
                                            //total_weight = 1;
                                            ////計算勝率 最新一年權重 =  1
                                            //avg_win_rate = avg_win_rate + win_rate * 1;
                                            //☆調整 逆序排列
                                            double year_weight = 1 - 0.05 * (dividen_times ) + (j+1) * 0.05;

                                            total_weight = total_weight + year_weight;

                                            double win_rate_percent = Math.Round(win_rate * 100, 2, MidpointRounding.AwayFromZero);
                                            double WADAR = Math.Round(((avg_win_rate / total_weight) * 100), 2, MidpointRounding.AwayFromZero);

                                            updateMsgNew(String.Format("最新 區間勝率{0}% 目前權重{1} WADAR={2}%{3}", Math.Round(win_rate * 100, 2), total_weight, WADAR, Environment.NewLine));

                                        }
                                        else
                                        {
                                            year_dividen = dividen_list[j].Split('@');
                                            //next_dividen = dividen_list[j - 1].Split('@');
                                            //☆逆序排列 改成+1
                                            next_dividen = dividen_list[j + 1].Split('@');

                                             updateMsgNew(String.Format("[{0}~{1}]股息{2}元 填息價{3}元 除權價{4}元", year_dividen[0], next_dividen[0], year_dividen[1], year_dividen[2], year_dividen[3]));
                                            //計算總交易日
                                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], next_dividen[4]);
                                            total_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                            //計算有填權息日數
                                            sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} and close >= {3}", stock_code, year_dividen[4], next_dividen[4], year_dividen[2]);
                                            fill_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                            //☆☆debug
                                            if (total_days == 0)
                                                win_rate = 0.0;
                                            else
                                                win_rate = (double)fill_days / total_days;

                                            //取得區間最大、最小股價
                                            sql = String.Format("select max(close) max_price,min(close) min_price from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], next_dividen[4]);
                                            max_price = Convert.ToDouble(DB_SQL(sql, "max_price"));
                                            min_price = Convert.ToDouble(DB_SQL(sql, "min_price"));

                                            //計算最大漲跌幅
                                            max_range = (Math.Round((max_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                                            min_range = (Math.Round((min_price / Convert.ToDouble(year_dividen[2]) - 1), 3)) * 100;
                                            if (j == 0)
                                            {
                                                increase_max = max_range;
                                                increase_min = max_range;
                                                decrease_max = min_range;
                                                decrease_min = min_range;
                                            }
                                            else
                                            {
                                                if (max_range > increase_max) increase_max = max_range;
                                                if (max_range < increase_min) increase_min = max_range;
                                                if (min_range > decrease_min) decrease_min = min_range;
                                                if (min_range < decrease_max) decrease_max = min_range;
                                            }

                                            //計算是否填權息
                                            if (fill_days > 0)
                                            {
                                                win_times++;
                                                //找出收盤價>填息價格的「日期」
                                                sql = String.Format("select min(stock_date) min_date from stock_price_new where stock_code='{0}' and compare_date between {1} and {2} and close >= {3}", stock_code, year_dividen[4], next_dividen[4], year_dividen[2]);
                                                min_date = DB_SQL(sql, "min_date");

                                                //計算兩區間之間 填息最小交易日數(分子)
                                                least_fill_days = (Convert.ToDateTime(min_date).Date - Convert.ToDateTime(year_dividen[0]).Date).Days;
                                                sql = String.Format("select count(*) cnt from stock_price_new where stock_code='{0}' and compare_date between {1} and {2}", stock_code, year_dividen[4], min_date.Replace("/", ""), year_dividen[2]);
                                                //updateMsgNew(sql);
                                                //計算兩區間之間 總交易日數 (分母)
                                                least_trade_days = Convert.ToInt32(DB_SQL(sql, "cnt"));
                                                avg_trade_days = avg_trade_days + least_trade_days;
                                            }
                                            //else //未填息的情況下 
                                            //{
                                            //    //最小填息交易日數設成-1
                                            //    least_fill_days = -1;
                                            //    min_date = "";
                                            //}
                                            updateMsgNew(String.Format("總交易天數{0}天 超過填息價總計{1}天{8}=> 勝率 {2} % 最高股價{3}元(填息後最大漲幅{4}%) 最低股價{5}元(除息後最大跌幅{6}%){8}    程式計算{7} [{10}~{11}]{8}vs GoodInfo計算最小天數{9}天"
                                                 , total_days, fill_days, (win_rate * 100).ToString("F1"), max_price, max_range, min_price, min_range, (fill_days > 0) ? " 填息最少交易" + least_trade_days + "天(日曆" + least_fill_days + "天)" : " 沒填息", Environment.NewLine, year_dividen[5], year_dividen[4], min_date));

                                            //計算股利合
                                            dividen = dividen + Convert.ToDouble(year_dividen[1]);
                                            //每年的權重 依序為  0.9 ~ 0.8 ... 0.1
                                            //double year_weight = 1 - 0.1 * j;
                                            //☆逆序推算年度權重 先扣再反加回來
                                            double year_weight = 1 - 0.05 * (dividen_times) + (j+1)*0.05;

                                            total_weight = total_weight + year_weight;
                                            //計算勝率 採用加權比重平均
                                            avg_win_rate = avg_win_rate + win_rate * (year_weight);
                                            String last_price = DB_SQL(String.Format("select close from stock_price_new where stock_code = '{0}' and stock_date<'{1}' order by stock_date desc limit 0,1", stock_code, next_dividen[0]), "close");

                                            //☆☆ 寫入conclude_data 作為AI迴歸測試用資料
                                            double win_rate_percent = Math.Round(win_rate * 100, 2, MidpointRounding.AwayFromZero);
                                            double WADAR = Math.Round(((avg_win_rate / total_weight) * 100), 2, MidpointRounding.AwayFromZero);

                                            updateMsgNew(String.Format("區間勝率{0}% 目前權重{1} WADAR={2}%{3}", Math.Round(win_rate*100,2), total_weight, WADAR , Environment.NewLine));


                                            sql = String.Format("insert into conclude_data (stock_code,stock_name,fill_days,total_days,win_rate         ,fill_dividen        , start_date     ,end_date        ,dividen          ,max_rise ,max_drop ,ai_max_rise,ai_max_drop,last_price,wadar) values ('{0}','{1}',{2},{3},{4},{5},'{6}','{7}',{8},{9},{10},{11},{12},{13},{14})",
                                                                                            stock_code, stock_name, fill_days, total_days, win_rate_percent, (fill_days > 0) ? 1 : 0, year_dividen[0], next_dividen[0], year_dividen[1], max_range, min_range, -1, -1, last_price, WADAR);
                                            updateNum = DB_SQL(sql.ToString());
                                            if (updateNum.All(Char.IsDigit))
                                            {
                                                if (Convert.ToInt32(updateNum) > 0)
                                                    updateMsgNew("寫入AI測試資料成功(" + updateNum.ToString() + ")");
                                                else
                                                    updateMsgNew("寫入AI測試資料失敗(" + updateNum.ToString() + ")");
                                            }
                                            else
                                                updateMsgNew("寫入AI測試資料失敗(" + sql + ")");
                                        }
                                    }

                                }

                            }//end of else
                            //計算 哇達 勝率
                            String WADAR_win_rate = ((avg_win_rate / total_weight) * 100).ToString("F1");

                            //計算 填息平均最小天數
                            if (win_times > 0)//填息次數至少要大於1次
                                avg_trade_days = Math.Round(avg_trade_days / (double)win_times, 2);
                            else//從來沒有填息過 設成-1
                                avg_trade_days = -1;
                            //計算 單一筆 花費時間
                            TimeSpan ts2 = DateTime.Now - t2;
                            min_buy_price = Math.Round(latest_fill_price * (100 + decrease_min) / 100, 2);
                            min_sell_price = Math.Round(latest_fill_price * (100 + increase_min) / 100, 2);

                            updateMsgNew(String.Format("[{0}]{1}{6}十年發 {2}次股利(總合{3}元)填權息{4}次 ☆WADAR指標={5}% {6}填息後最大漲幅介於{7}%~{8}% 除息後最大跌幅介於{9}%~{10}%",
                                      stock_code, stock_name, dividen_times, dividen, win_times, WADAR_win_rate, Environment.NewLine, increase_min, increase_max, decrease_min, decrease_max));
                            updateMsgNew(String.Format("★ 最新填息價 {0} 元,推估至少低於{1}元({2}%)買入(買進點) 高於至少{3}元({4}%)賣出(停利點)", latest_fill_price, min_buy_price, decrease_min, min_sell_price, increase_min));
                            updateMsgNew(String.Format("填息最小交易日數平均:{0}天 計算總計花費時間{1}秒", avg_trade_days, ts2.TotalSeconds));
                            updateMsgNew("===============================================================");
                            //    Thread.Sleep(200000);


                            //sql = String.Format("insert into money_rank_new (stock_code,stock_name,win_times,dividen_times,dividen,avg_win_rate,cal_date,least_days) values ('{0}','{1}',{2},{3},{4},{5},'{6}',{7})", stock_list[i], stock_name, win_times, dividen_times, dividen, WADAR_win_rate, DateTime.Now.ToString("yyyy/MM/dd"), avg_trade_days);

                            //updateNum = DB_SQL(sql.ToString());
                            //if (updateNum.All(Char.IsDigit))
                            //{
                            //    if (Convert.ToInt32(updateNum) > 0)
                            //        updateMsgNew("寫入排名資料成功(" + updateNum.ToString() + ")");
                            //    else
                            //        updateMsgNew("寫入排名資料失敗(" + updateNum.ToString() + ")");
                            //}
                            //else
                            //    updateMsgNew("寫入排名資料失敗(" + updateNum + ")");

                        }
                        catch (Exception exx)
                        {
                            updateMsgNew(exx.Message);
                        }
                    }
                    TimeSpan ts1 = DateTime.Now - t1;
                    updateMsgNew(String.Format("統計：計算{0}隻股票 總花費時間{1}秒", stock_list.Length, ts1.TotalSeconds));

                    //this.Invoke((MethodInvoker)delegate
                    //{
                    //    label3.Text = String.Format("說明：挑出10年內(1)連續{0}年發股息 (2)全部填權息 ☆(3)WADAR指標前五十名 『總排行榜』", trackBarValue);
                    //    label3.Visible = true;
                    //});

                    ////更新GridView
                    //sql = " select stock_code,stock_name,win_times,dividen_times,avg_win_rate,dividen from money_rank_new " +
                    //    " where dividen_times >= " + trackBarValue +
                    //    " order by (win_times/dividen_times) desc ,dividen_times desc,avg_win_rate desc limit 0,50";
                    //result = DB_SQL(sql, "stock_code,stock_name,win_times,dividen_times,avg_win_rate,dividen");

                    //CreateGVData(dataGridView2, sql);

                    //this.Invoke((MethodInvoker)delegate
                    //{
                    //    label3.Text = "說明：挑出10年內(1)至少連續" + trackBarValue + "年發股息 (2)全部填權息 ☆(3)WADAR指標前五十名 (4)尚未填息 的『可買進名單』";
                    //    label3.Visible = true;
                    //});

                    ////更新GridView
                    //sql = "select m.stock_code,m.stock_name,m.win_times,m.dividen_times,m.avg_win_rate,p.close,d.before_price, " +
                    //           "  Round((d.before_price - p.close),2) price_diff, Round(100*(d.before_price - p.close) / before_price,3) price_percent,d.dividen_date,p.stock_date " +
                    //           "       from money_rank_new m left outer join latest_stock_price p " +
                    //           "      on m.stock_code = p.stock_code " +
                    //           "   left outer join latest_stock_dividen d on d.stock_code = m.stock_code " +
                    //           "   where m.dividen_times >=" + trackBarValue + " and p.close < d.before_price and m.win_times = m.dividen_times " +
                    //           "   order by(price_diff/ before_price) desc , m.avg_win_rate desc limit 0,50";
                    //result = DB_SQL(sql, "stock_code,stock_name,win_times,dividen_times,avg_win_rate,close,before_price,price_diff,price_percent,dividen_date,stock_date");

                    //CreateGVData(dataGridView1, sql);
                }).Start();
            }
        }
    }
}
