using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CarRepaireReminder
{

    public partial class Form1 : Form
    {

        private bool isExit;

        public static Dictionary<string, int> username_2_id = new Dictionary<string, int>();
        public Dictionary<string, ExpireReminderItem> carno_2_items = new Dictionary<string, ExpireReminderItem>();
        public Dictionary<String, List<ExpireReminderItem2>> expireList = new Dictionary<string, List<ExpireReminderItem2>>();
        public static String connectionString;

        public int output_format = 1;
        public int remind_days_before = 30;

        public void loadOperator()
        {
            username_2_id.Clear();
            this.Cursor = Cursors.WaitCursor;

            
            output_format = Convert.ToInt32(ConfigurationManager.AppSettings["output_format"]);
            remind_days_before = Convert.ToInt32(ConfigurationManager.AppSettings["remind_days_before"]);
            connectionString = GetConnectionStringsConfig("netmis");
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = con.CreateCommand())
                {
                    con.Open();


                    String sql = String.Format("select UserId, UserName from UserTable");
                    cmd.CommandText = sql;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {

                        while (reader.Read())
                        {
                            int user_id = Convert.ToInt32(reader["UserId"]);
                            string username = reader["UserName"].ToString().Trim();
                            username_2_id[username] = user_id;
                        }
                    }
                }
                con.Close();
            }
            this.Cursor = Cursors.Arrow;
        }

        public Form1()
        {
            InitializeComponent();
            loadOperator();
            String icon_file = "armis.ico";
            if (File.Exists(icon_file))
            {
                notifyIcon1.Icon = new Icon(icon_file);

                MenuItem[] trayMenu = new MenuItem[3];

                MenuItem startItem = new MenuItem();
                startItem.Text = "打开";
                startItem.Click += new EventHandler(startItem_Click);

                MenuItem autoStartItem = new MenuItem();
                autoStartItem.Text = "自动启动";
                autoStartItem.Checked = autoStartSet();
                autoStartItem.Click += new EventHandler(autoStart_click);

                MenuItem exitItem = new MenuItem();
                exitItem.Text = "退出";
                exitItem.Click += new EventHandler(exitItem_Click);

                trayMenu[0] = startItem;
                 trayMenu[1] = autoStartItem;
                trayMenu[2] = exitItem;

                notifyIcon1.ContextMenu = new ContextMenu(trayMenu);

                notifyIcon1.Click += new EventHandler(notifyIcon_Click);
                //   notifyIcon1.DoubleClick += new EventHandler(notifyIcon_Click);
            }

            // 程序默认启动时隐藏窗体
            //   windowDisplay(false);

            expireList["保险"] = new List<ExpireReminderItem2>();
            expireList["保养"] = new List<ExpireReminderItem2>();
            expireList["年审"] = new List<ExpireReminderItem2>();

            isExit = false;
        }

        /// <summary>
        /// 自定义方法：窗体的隐藏与显示
        /// </summary>
        /// <param name="display"></param>
        private void windowDisplay(bool display)
        {
            if (display)
            {
                this.WindowState = FormWindowState.Normal; // 窗口常规化
                this.ShowInTaskbar = true; // 显示在任务栏
            }
            else
            {
                this.WindowState = FormWindowState.Minimized; // 窗口最小化
                this.ShowInTaskbar = false; // 不显示在任务栏
            }

        }
        void notifyIcon_Click(object sender, EventArgs e)
        {
            windowDisplay(this.WindowState == FormWindowState.Minimized);

        }

        void startItem_Click(object sender, EventArgs e)
        {
            windowDisplay(true);
        }

        public Boolean autoStartSet()
        {
            string path = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "CarExpireReminder", "");
            return !path.Equals("");
        }

        void  autoStart_click(object sender, EventArgs e)
        {
            if (autoStartSet()) //设置开机自启动  
            {
               // MessageBox.Show("设置开机自启动，需要修改注册表", "提示");
                string path = Application.ExecutablePath;
                RegistryKey rk = Registry.LocalMachine;
                RegistryKey rk2 = rk.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
                rk2.DeleteValue("CarExpireReminder", false);
                rk2.Close();
                rk.Close();
                notifyIcon1.ContextMenu.MenuItems[1].Checked = false;
            }
            else //取消开机自启动  
            {
                //  MessageBox.Show("取消开机自启动，需要修改注册表", "提示");
                string path = Application.ExecutablePath;
                RegistryKey rk = Registry.LocalMachine;
                RegistryKey rk2 = rk.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
                rk2.SetValue("CarExpireReminder", path);
                rk2.Close();
                rk.Close();
                notifyIcon1.ContextMenu.MenuItems[1].Checked = true;
            }
        }

        void exitItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            fillData();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            windowDisplay(false);
            e.Cancel = true;
        }

        public string GetConnectionStringsConfig(string connectionName)
        {
            //指定config文件读取
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            connectionString =
                config.ConnectionStrings.ConnectionStrings[connectionName].ConnectionString.ToString();
            return connectionString;
        }

        public void fillData()
        {
            loadOperator();
            this.dataGridView1.Rows.Clear();
            this.dataGridView2.Rows.Clear();
            this.dataGridView3.Rows.Clear();

            //清空展示列表数据结构
            foreach (var item in expireList)
            {
                item.Value.Clear();
            }
            
            this.carno_2_items.Clear();

            this.Cursor = Cursors.WaitCursor;
            string connStr = GetConnectionStringsConfig("netmis");
            using (SqlConnection con = new SqlConnection(connStr))
            {
                using (SqlCommand cmd = con.CreateCommand())
                {
                    con.Open();

                    String sql = String.Format("select car_no, next_baoyang_time, next_nianshen_time,next_baoxian_time, a2.username from  dt_expire_reminder a1 join UserTable a2 on a1.operator_id=a2.userid where next_baoyang_time<'{0}' or next_nianshen_time<'{0}' or next_baoxian_time<'{0}'", DateTime.Now.AddDays(remind_days_before).ToString("yyyy-MM-dd 00:00:00"));
                    cmd.CommandText = sql;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {

                        while (reader.Read())
                        {
                            string car_no = reader["car_no"].ToString().Trim();
                            string today = DateTime.Now.AddDays(remind_days_before).ToString("yyyy-MM-dd");

                            string next_baoyang_time = Convert.ToDateTime(reader["next_baoyang_time"]).ToString("yyyy-MM-dd");                            
                            string next_nianshen_time = Convert.ToDateTime(reader["next_nianshen_time"]).ToString("yyyy-MM-dd");
                            string next_baoxian_time = Convert.ToDateTime(reader["next_baoxian_time"]).ToString("yyyy-MM-dd");
                            string username = reader["username"].ToString().Trim();
                            carno_2_items[car_no] = new ExpireReminderItem(car_no, next_baoyang_time, next_nianshen_time, next_baoxian_time, username);
                            //expireList["保养"] = new List<ExpireReminderItem2>();
                            if (next_baoyang_time.CompareTo(today) < 0)
                            {
                                expireList["保养"].Add(new ExpireReminderItem2(car_no, next_baoyang_time));
                            }
                            if (next_nianshen_time.CompareTo(today) < 0)
                            {
                                expireList["年审"].Add(new ExpireReminderItem2(car_no, next_nianshen_time));
                            }
                            if (next_baoxian_time.CompareTo(today) < 0)
                            {
                                expireList["保险"].Add(new ExpireReminderItem2(car_no, next_baoxian_time));
                            }

                        }
                    }
                }
                con.Close();
            }
            this.Cursor = Cursors.Arrow;

            fillDataInternal();
        }

        private void fillDataInternal2(List<ExpireReminderItem2> expireItems, DataGridView dataGridView)
        {
            for (int idx=0; idx<expireItems.Count; idx++)
            {
                int index = dataGridView.Rows.Add();
                dataGridView.Rows[index].Cells[0].Value = expireItems[idx].car_no;
                dataGridView.Rows[index].Cells[1].Value = expireItems[idx].expire_time;
                if (expireItems[index].need_mark)
                {
                    dataGridView.Rows[index].Cells[0].Style.BackColor = Color.Red;
                    dataGridView.Rows[index].Cells[1].Style.BackColor = Color.Red;
                    dataGridView.Rows[index].Cells[2].Style.BackColor = Color.Red;
                }
            }
        }
        private void fillDataInternal()
        {
            var nianshenList = expireList["年审"].OrderByDescending(item => item.expire_time).ToList();
            var baoyangList = expireList["保养"].OrderByDescending(item => item.expire_time).ToList();
            var baoxianList = expireList["保险"].OrderByDescending(item => item.expire_time).ToList();

            fillDataInternal2(nianshenList, dataGridView1);
            fillDataInternal2(baoyangList, dataGridView2);
            fillDataInternal2(baoxianList, dataGridView3);


            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Rows[0].Selected = false;
            }
            if (dataGridView2.Rows.Count > 0)
            {
                dataGridView2.Rows[0].Selected = false;
            }
            if (dataGridView3.Rows.Count > 0)
            {
                dataGridView3.Rows[0].Selected = false;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            fillData();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            EditForm editForm = new EditForm();
            editForm.StartPosition = FormStartPosition.CenterParent;
            editForm.setForm1(this);
            editForm.ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows.Count < 1)
            {
                return;
            }
            if (dataGridView1.Columns[e.ColumnIndex].Name == "EditButton" && e.RowIndex >= 0)
            {
                EditForm editForm = new EditForm();
                editForm.StartPosition = FormStartPosition.CenterParent;
                editForm.setForm1(this);
                String car_no = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                ExpireReminderItem expireReminderItem = this.carno_2_items[car_no];
                editForm.setExpireReminderItem(expireReminderItem);
                editForm.ShowDialog();
            }

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.Rows.Count < 1)
            {
                return;
            }
            if (dataGridView2.Columns[e.ColumnIndex].Name == "EditButton2" && e.RowIndex >= 0)
            {
                EditForm editForm = new EditForm();
                editForm.StartPosition = FormStartPosition.CenterParent;
                editForm.setForm1(this);
                String car_no = this.dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                ExpireReminderItem expireReminderItem = this.carno_2_items[car_no];
                editForm.setExpireReminderItem(expireReminderItem);
                editForm.ShowDialog();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.DefaultExt = "xlsx";
            sfd.Filter = "Excel File(*.xlsx)|*.xlsx";
            sfd.RestoreDirectory = true;



            //if (carno_2_items.Count == 0)
            //{
            //    MessageBox.Show("没有纪录要导出", "提示", MessageBoxButtons.OK);
            //    return;
            //}

            sfd.FileName = String.Format("年审保养保险({0})", DateTime.Now.ToString("yyyy-MM-dd") );

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                string localFilePath = sfd.FileName.ToString(); //获得文件路径   
                DataTableToExcel2(localFilePath, false);
            }
        }

        public List<ExpireReminderItem> queryData()
        {
            List<ExpireReminderItem> datas = new List<ExpireReminderItem>();
            string connStr = GetConnectionStringsConfig("netmis");
            using (SqlConnection con = new SqlConnection(connStr))
            {
                using (SqlCommand cmd = con.CreateCommand())
                {
                    con.Open();


                    String sql = String.Format("select car_no, next_baoyang_time, next_nianshen_time,next_baoxian_time, a2.username from  dt_expire_reminder a1 join UserTable a2 on a1.operator_id=a2.userid");
                    cmd.CommandText = sql;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {

                        while (reader.Read())
                        {
                            string car_no = reader["car_no"].ToString().Trim();
                            string one_month_ago = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-dd");

                            string next_baoyang_time = Convert.ToDateTime(reader["next_baoyang_time"]).ToString("yyyy-MM-dd");
                            string next_nianshen_time = Convert.ToDateTime(reader["next_nianshen_time"]).ToString("yyyy-MM-dd");
                            string next_baoxian_time = Convert.ToDateTime(reader["next_baoxian_time"]).ToString("yyyy-MM-dd");
                            string username = reader["username"].ToString().Trim();
                            ExpireReminderItem item = new ExpireReminderItem(car_no, next_baoyang_time, next_nianshen_time, next_baoxian_time, username);
                            datas.Add(item);                          
                        }
                    }
                }
                con.Close();
            }

            return datas;
        }

        public bool DataTableToExcel2(string filePath, bool isShowExcle)
        {
            this.Cursor = Cursors.WaitCursor;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
            excel.Visible = isShowExcle;

            List<ExpireReminderItem> datas = queryData();
            object[,] objData = null;
            int total_row_numbers = datas.Count;
            if (output_format == 1)
            {
                excel.Cells[1, 1] = "车牌号码";
                excel.Cells[1, 2] = "到期名称";
                excel.Cells[1, 3] = "到期时间";
                excel.Cells[1, 4] = "操作员";

                total_row_numbers = datas.Count * 3;

                objData = new object[total_row_numbers, 4];
                for (int index = 0; index < datas.Count; index++)
                {
                    objData[index, 0] = datas[index].car_no;
                    objData[index, 1] = "年审到期";
                    objData[index, 2] = datas[index].next_nianshen_time;
                    objData[index, 3] = datas[index].username;
                }


                for (int index = 0; index < datas.Count; index++)
                {
                    objData[index + datas.Count, 0] = datas[index].car_no;
                    objData[index + datas.Count, 1] = "保养到期";
                    objData[index + datas.Count, 2] = datas[index].next_baoyang_time;
                    objData[index + datas.Count, 3] = datas[index].username;
                }

                for (int index = 0; index < datas.Count; index++)
                {
                    objData[index + datas.Count * 2, 0] = datas[index].car_no;
                    objData[index + datas.Count * 2, 1] = "保险到期";
                    objData[index + datas.Count * 2, 2] = datas[index].next_baoxian_time;
                    objData[index + datas.Count * 2, 3] = datas[index].username;
                }
            }
            else if (output_format==2)
            {
                excel.Cells[1, 1] = "车牌号码";
                excel.Cells[1, 2] = "年审到期";
                excel.Cells[1, 3] = "保养到期";
                excel.Cells[1, 4] = "保险到期";
                excel.Cells[1, 5] = "操作员";

                total_row_numbers = datas.Count;

                objData = new object[total_row_numbers, 5];
                for (int index = 0; index < datas.Count; index++)
                {
                    objData[index, 0] = datas[index].car_no;
                    objData[index, 1] = datas[index].next_nianshen_time;
                    objData[index, 2] = datas[index].next_baoyang_time;
                    objData[index, 3] = datas[index].next_baoxian_time;
                    objData[index, 4] = datas[index].username;
                }
            }


            Microsoft.Office.Interop.Excel.Range range = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[total_row_numbers + 1, 5]];
            range.Value = objData;

            range.EntireColumn.AutoFit();
            worksheet.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excel.Quit();
            //excel.Quit();
            this.Cursor = Cursors.Arrow;
            return true;
        }

        public bool DataTableToExcel(string filePath, bool isShowExcle)
        {
            this.Cursor = Cursors.WaitCursor;
            if (carno_2_items.Count == 0)
            {
                return false;
            }

            //System.Data.DataTable dataTable = dataSet.Tables[0];  
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
            excel.Visible = isShowExcle;
            excel.Cells[1, 1] = "车牌号码";
            excel.Cells[1, 2] = "到期名称";
            excel.Cells[1, 3] = "到期时间";
            excel.Cells[1, 4] = "操作员";

            int total_row_numbers = this.dataGridView1.Rows.Count + this.dataGridView2.Rows.Count;

            object[,] objData = new object[total_row_numbers, 4];

            for (int index = 0; index < dataGridView1.Rows.Count; index++)
            {
                objData[index, 0] = dataGridView1.Rows[index].Cells[0].Value.ToString();
                objData[index, 1] = "年审到期";
                objData[index, 2] = dataGridView1.Rows[index].Cells[1].Value.ToString();
                objData[index, 3] = carno_2_items[dataGridView1.Rows[index].Cells[0].Value.ToString()].username;
            }
            for (int index = 0; index < dataGridView2.Rows.Count; index++)
            {
                objData[index + dataGridView1.Rows.Count, 0] = dataGridView2.Rows[index].Cells[0].Value.ToString();
                objData[index + dataGridView1.Rows.Count, 1] = "保养到期";
                objData[index + dataGridView1.Rows.Count, 2] = dataGridView2.Rows[index].Cells[1].Value.ToString();
                objData[index + dataGridView1.Rows.Count, 3] = carno_2_items[dataGridView2.Rows[index].Cells[0].Value.ToString()].username;
            }

            Microsoft.Office.Interop.Excel.Range range = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[total_row_numbers + 1, 4]];
            range.Value = objData;

            range.EntireColumn.AutoFit();
            worksheet.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excel.Quit();
            //excel.Quit();
            this.Cursor = Cursors.Arrow;
            return true;
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int index=0; index<dataGridView2.Rows.Count; index++)
            {
                if (dataGridView2.Rows[index].Selected)
                {
                    dataGridView2.Rows[index].Selected = false;
                }
            }
            for (int index = 0; index < dataGridView3.Rows.Count; index++)
            {
                if (dataGridView3.Rows[index].Selected)
                {
                    dataGridView3.Rows[index].Selected = false;
                }
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int index = 0; index < dataGridView1.Rows.Count; index++)
            {
                if (dataGridView1.Rows[index].Selected)
                {
                    dataGridView1.Rows[index].Selected = false;
                }
            }
            for (int index = 0; index < dataGridView3.Rows.Count; index++)
            {
                if (dataGridView3.Rows[index].Selected)
                {
                    dataGridView3.Rows[index].Selected = false;
                }
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int index = 0; index < dataGridView1.Rows.Count; index++)
            {
                if (dataGridView1.Rows[index].Selected)
                {
                    dataGridView1.Rows[index].Selected = false;
                }
            }
            for (int index = 0; index < dataGridView2.Rows.Count; index++)
            {
                if (dataGridView2.Rows[index].Selected)
                {
                    dataGridView2.Rows[index].Selected = false;
                }
            }
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView3.Rows.Count < 1)
            {
                return;
            }
            if (dataGridView3.Columns[e.ColumnIndex].Name == "EditButton3" && e.RowIndex >= 0)
            {
                EditForm editForm = new EditForm();
                editForm.StartPosition = FormStartPosition.CenterParent;
                editForm.setForm1(this);
                String car_no = this.dataGridView3.Rows[e.RowIndex].Cells[0].Value.ToString();
                ExpireReminderItem expireReminderItem = this.carno_2_items[car_no];
                editForm.setExpireReminderItem(expireReminderItem);
                editForm.ShowDialog();
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            QueryForm2 queryForm2
                = new QueryForm2();
            queryForm2.StartPosition = FormStartPosition.CenterParent;
            queryForm2.setForm1(this);
            queryForm2.ShowDialog();
        }
    }


}
