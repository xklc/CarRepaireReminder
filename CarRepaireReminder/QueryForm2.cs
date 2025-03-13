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
    public partial class QueryForm2 : Form
    {
        private Form1 form1;

        private Dictionary<string, ExpireReminderItem> carno_2_matchedItems = new Dictionary<string, ExpireReminderItem>();

        private Boolean queryButtonClieck = true;
        public void setForm1(Form1 form1)
        {
            this.form1 = form1;
        }

        public QueryForm2()
        {
            InitializeComponent();
        }

        public void fillData()
        {
            string car_no = this.textBox1.Text.Trim();
            this.dataGridView1.Rows.Clear();
            this.Cursor = Cursors.WaitCursor;
            string connStr = Form1.connectionString;
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

                        carno_2_matchedItems.Clear();

                        string one_month_ago = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-dd");

                        while (reader.Read())
                        {
                            exist = true;
                            car_no = reader["car_no"].ToString().Trim();
                            string next_baoyang_time = Convert.ToDateTime(reader["next_baoyang_time"]).ToString("yyyy-MM-dd");
                            string next_nianshen_time = Convert.ToDateTime(reader["next_nianshen_time"]).ToString("yyyy-MM-dd");
                            string next_baoxian_time = Convert.ToDateTime(reader["next_baoxian_time"]).ToString("yyyy-MM-dd");
                            string username = reader["username"].ToString().Trim();

                            var expireReminderItem = new ExpireReminderItem(car_no, next_baoyang_time, next_nianshen_time, next_baoxian_time, username);
                            carno_2_matchedItems[car_no] = expireReminderItem;
                            int index = this.dataGridView1.Rows.Add();
                            dataGridView1.Rows[index].Cells[0].Value = car_no;
                            dataGridView1.Rows[index].Cells[1].Value = next_nianshen_time;
                            dataGridView1.Rows[index].Cells[2].Value = next_baoyang_time;
                            dataGridView1.Rows[index].Cells[3].Value = next_baoxian_time;
                            dataGridView1.Rows[index].Cells[4].Value = username;

                            Boolean need_mark = next_baoyang_time.CompareTo(one_month_ago) < 0 || next_nianshen_time.CompareTo(one_month_ago) < 0 || next_baoxian_time.CompareTo(one_month_ago) < 0;
                            if (need_mark)
                            {
                                dataGridView1.Rows[index].Cells[0].Style.BackColor = Color.Red;
                                dataGridView1.Rows[index].Cells[1].Style.BackColor = Color.Red;
                                dataGridView1.Rows[index].Cells[2].Style.BackColor = Color.Red;
                                dataGridView1.Rows[index].Cells[4].Style.BackColor = Color.Red;
                                dataGridView1.Rows[index].Cells[3].Style.BackColor = Color.Red;
                            }
                        }
                    }

                    if (!exist)
                    {
                        MessageBox.Show("没有符合条件的记录被查找到！");
                    }
                }
            }
            this.Cursor = Cursors.Arrow;
            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Rows[0].Selected = false;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            queryButtonClieck = true;
            string car_no = this.textBox1.Text.Trim();
            //1. 判断车牌号是否为空
            if (car_no.Length==0)
            {
                MessageBox.Show("车牌号为空或者长度不对", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.textBox1.Focus();
                return;
            }

            fillData();

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows.Count < 1)
            {
                return;
            }
            queryButtonClieck = false;
            if (dataGridView1.Columns[e.ColumnIndex].Name == "EditButton" && e.RowIndex >= 0)
            {
                EditForm editForm = new EditForm();
                editForm.StartPosition = FormStartPosition.CenterParent;
                editForm.setQueryForm2(this);
                editForm.setForm1(this.form1);
                String car_no = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                ExpireReminderItem expireReminderItem = this.carno_2_matchedItems[car_no];
                editForm.setExpireReminderItem(expireReminderItem);
                editForm.ShowDialog();
            }
            else if (dataGridView1.Columns[e.ColumnIndex].Name == "DeleteButton" && e.RowIndex >= 0)
            {
                String car_no = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                String errorMsg = String.Format("是否删除车牌号为{0}的记录", car_no);

                if (MessageBox.Show(errorMsg, "删除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;
                    string connStr = Form1.connectionString;
                    using (System.Data.SqlClient.SqlConnection con = new SqlConnection(connStr))
                    {
                        using (SqlCommand cmd = con.CreateCommand())
                        {
                            con.Open();
                            String sql = String.Format("delete from dt_expire_reminder where car_no = '{0}'", car_no);
                            cmd.CommandText = sql;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    this.Cursor = Cursors.Arrow;
                    this.fillData();
                }
            }
        }

        private void QueryForm2_FormClosed(object sender, FormClosedEventArgs e)
        {
            form1.fillData();
        }
    }
}
