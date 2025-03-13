using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarRepaireReminder
{
    public class ExpireReminderItem
    {
        public String car_no;
        public String next_baoyang_time;
        public String next_nianshen_time;
        public String next_baoxian_time;
        public String username;

        public ExpireReminderItem()
        {

        }

        public ExpireReminderItem(String car_no, String next_baoyang_time, String next_nianshen_time, String next_baoxian_time, String username)
        {
            this.car_no = car_no;
            this.next_baoyang_time = next_baoyang_time;
            this.next_nianshen_time = next_nianshen_time;
            this.next_baoxian_time = next_baoxian_time;
            this.username = username;
        }
    }

    public class ExpireReminderItem2
    {
        public String car_no;
        public String expire_time;        
        private static string one_month_ago = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-dd");
        public Boolean need_mark;

        public ExpireReminderItem2()
        {

        }

        public ExpireReminderItem2(String car_no, String expire_time)
        {
            this.car_no = car_no;
            this.expire_time = expire_time;
            this.need_mark = expire_time.CompareTo(one_month_ago) < 0;
        }
    }
}
