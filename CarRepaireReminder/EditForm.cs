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
    public partial class EditForm : Form
    {
        private ExpireReminderItem expireReminderItem;
        private Form1 form1;
        private QueryForm2 queryForm2;
        private Boolean is_new = true;
        private String old_car_num = "";

        public void setExpireReminderItem(ExpireReminderItem expireReminderItem)
        {
            this.expireReminderItem = expireReminderItem;
            this.old_car_num = new StringBuilder(expireReminderItem.car_no).ToString();
            is_new = false;
        }

        public void setForm1(Form1 form1)
        {
            this.form1 = form1;
        }

        public void setQueryForm2(QueryForm2 queryForm2)
        {
            this.queryForm2 = queryForm2;
        }

        public EditForm()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string car_no = this.textBox1.Text.Trim();
            //1. 判断车牌号是否为空
            if (car_no.Length < 5)
            {
                MessageBox.Show("车牌号为空或者长度不对", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.textBox1.Focus();
                return;
            }

            if (car_no.Length > 10)
            {
                MessageBox.Show("车牌号有误", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.textBox1.Focus();
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            string connStr = form1.GetConnectionStringsConfig("netmis");
            using (System.Data.SqlClient.SqlConnection con = new SqlConnection(connStr))
            {
                using (SqlCommand cmd = con.CreateCommand())
                {
                    con.Open();


                    String sql = String.Format("select * from dt_expire_reminder where car_no = '{0}'", car_no);
                    cmd.CommandText = sql;
                    Boolean exist = false;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        String today = DateTime.Now.ToString("yyyy-MM-dd");

                        while (reader.Read())
                        {
                            exist = true;
                            break;
                        }
                    }

                    //if (exist)
                    //{
                    //    MessageBox.Show("车牌号码已经存在， 请检查车牌号码", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;
                    //}


                    string next_baoxian_time = this.dateTimePicker3.Value.ToString("yyyy-MM-dd 00:00:00");
                    string next_baoyang_time = this.dateTimePicker2.Value.ToString("yyyy-MM-dd 00:00:00");
                    string next_nianshen_time = this.dateTimePicker1.Value.ToString("yyyy-MM-dd 00:00:00");
                    int operator_id = Form1.username_2_id[this.comboBox1.Text.Trim()];
                    String delete_sql = null;
                    if (exist)
                    {
                        //String msg = String.Format("车牌号码[{0}]已经存在， 是否更新该记录, 原有的车牌号{1}记录将会被删除?", car_no, old_car_num);
                        if ((!old_car_num.Equals(car_no) && MessageBox.Show("车牌号码已经存在， 是否更新已有号码的记录?", "修改提醒", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) || old_car_num.Equals(car_no))
                        {
                            // delete_sql =  String.Format("delete from dt_expire_reminder where car_no='{0}'", old_car_num);
                            if (!old_car_num.Equals(car_no) && old_car_num.Length > 0)
                            {
                                sql = String.Format("delete from dt_expire_reminder where car_no='{0}'", old_car_num);
                                cmd.CommandText = sql;
                                cmd.ExecuteNonQuery();
                            }
                            sql = String.Format("update dt_expire_reminder set next_baoyang_time='{0}' , next_nianshen_time='{1}', next_baoxian_time='{4}', operator_id={2} where car_no='{3}'", next_baoyang_time, next_nianshen_time, operator_id, car_no, next_baoxian_time);
                        }

                    }
                    else
                    {
                        if (!old_car_num.Equals(car_no) && old_car_num.Length>0)
                        {
                            sql = String.Format("update dt_expire_reminder set next_baoyang_time='{0}' , next_nianshen_time='{1}', next_baoxian_time='{4}', operator_id={2}, car_no='{3}' where car_no='{5}'", next_baoyang_time, next_nianshen_time, operator_id, car_no, next_baoxian_time, old_car_num);
                        }
                        else { 
                            sql = String.Format("insert into dt_expire_reminder(car_no,next_baoyang_time,next_nianshen_time,next_baoxian_time, operator_id) values('{3}', '{0}','{1}','{4}',{2})", next_baoyang_time, next_nianshen_time, operator_id, car_no, next_baoxian_time);
                        }
                    }
                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                    // if 
                }
            }
            this.Cursor = Cursors.Arrow;
            this.Close();
            if (this.queryForm2 != null)
            {
                this.queryForm2.fillData();
            }
            else if (this.form1 != null)
            {
                this.form1.fillData();
            }
        }



        private void EditForm_Load(object sender, EventArgs e)
        {
            if (is_new)
            {
                foreach (var username in Form1.username_2_id.Keys)
                {
                    this.comboBox1.Items.Add(username);
                }
                comboBox1.SelectedIndex = 0;
            }
            else
            {
                this.textBox1.Text = expireReminderItem.car_no;
                this.dateTimePicker1.Text = expireReminderItem.next_nianshen_time;
                this.dateTimePicker2.Text = expireReminderItem.next_baoyang_time;
                this.dateTimePicker3.Text = expireReminderItem.next_baoxian_time;
                this.comboBox1.Items.Clear();
                foreach (var username in Form1.username_2_id.Keys)
                {
                    this.comboBox1.Items.Add(username);
                }
                this.comboBox1.SelectedItem = expireReminderItem.username;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
