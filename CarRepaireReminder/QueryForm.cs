using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CarRepaireReminder
{
    public partial class QueryForm : Form
    {
        private Form1 form1;
        public void setForm1(Form1 form1)
        {
            this.form1 = form1;
        }

        public QueryForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string car_no = this.textBox1.Text.Trim();
            //1. 判断车牌号是否为空
            if (car_no.Length==0)
            {
                MessageBox.Show("车牌号为空或者长度不对", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.textBox1.Focus();
                return;
            }

            this.listView1.Items.Clear();
            this.Cursor = Cursors.WaitCursor;
            string connStr = form1.GetConnectionStringsConfig("netmis");
            using (System.Data.SqlClient.SqlConnection con = new SqlConnection(connStr))
            {
                using (SqlCommand cmd = con.CreateCommand())
                {
                    con.Open();


                    String sql = String.Format("select car_no, next_baoyang_time, next_nianshen_time,next_baoxian_time, a2.username from  dt_expire_reminder a1 join UserTable a2 on a1.operator_id=a2.userid and a1.car_no like '%{0}%'", car_no);
                    cmd.CommandText = sql;
                    Boolean exist = false;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                
                        while (reader.Read())
                        {
                            exist = true;
                            car_no = reader["car_no"].ToString().Trim();
                            string next_baoyang_time = Convert.ToDateTime(reader["next_baoyang_time"]).ToString("yyyy-MM-dd");
                            string next_nianshen_time = Convert.ToDateTime(reader["next_nianshen_time"]).ToString("yyyy-MM-dd");
                            string next_baoxian_time = Convert.ToDateTime(reader["next_baoxian_time"]).ToString("yyyy-MM-dd");
                            string username = reader["username"].ToString().Trim();
                            ListViewItem  lvi = this.listView1.Items.Add("");
                            lvi.SubItems.Add(car_no);
                            lvi.SubItems.Add(next_nianshen_time);
                            lvi.SubItems.Add(next_baoyang_time);
                            lvi.SubItems.Add(next_baoxian_time);
                            lvi.SubItems.Add(username);
                        }
                    }

                    if (!exist)
                    {
                        MessageBox.Show("没有符合条件的记录被查找到！");
                    }

                   
                }
            }
            this.Cursor = Cursors.Arrow;

        }
    }
}
