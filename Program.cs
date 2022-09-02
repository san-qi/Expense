using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;

namespace Expense
{

    class Program
    {
        public static void Test()
        {
            // Generate_reception("禹虹机", "王俊民", "石泉管理所", "维萨拉、艾萨拉、文库、布里茨、安岚", "江舟鱼舫", "晚", "890", "0", "1", "2022.8.26", "2022.8.29", "公务接待", false);
            Generate_trip("禹虹机", "石泉", "西安", "十天高速石泉站收费撤站手续办理", "陈念、夏开俊", "", "2022-10-15", "2022-10-16", DateTime.Now.ToString(), "1234", "12.34", "1", "0");
        }

        public static void Generate_reception(string name, string colleagues, string reception_employer, string reception_people,
            string target_place, string meal_time, string total_count, string tax_count, string tax_paper_number,
            string start_date, string reimbursement_date, string reason, bool have_wine_paper)
        {
            string path;
            using (var conn = new SqliteConnection("Data Source=conf.db"))
            {
                conn.Open();
                var command = conn.CreateCommand();
                command.CommandText = @"Select source from config where name = 'recption_path'";
                var reader = command.ExecuteReader();

                if (reader.Read() && reader["source"].ToString().Trim().Length > 0)
                {
                    path = reader["source"] as string;
                }
                else
                {
                    var desktop_dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    path = System.IO.Path.Combine(desktop_dir, String.Format("运筹部{0}+业务招待+{1}元", name, total_count));
                }
            }
            var list_file_name = System.IO.Path.Combine(path, "接待清单.xlsx");
            Generate_reception(name, colleagues, reception_employer, reception_people,
                target_place, meal_time, total_count, tax_count, tax_paper_number,
                start_date, reimbursement_date, reason, have_wine_paper, list_file_name);
        }

        public static void Generate_reception(string name, string colleagues, string reception_employer, string reception_people,
            string target_place, string meal_time, string total_count, string tax_count, string tax_paper_number,
            string start_date, string reimbursement_date, string reason, bool have_wine_paper, string list_file_name)
        {
            var pairs = new Dictionary<String, String>();
            pairs.Add("name", name);
            pairs.Add("colleagues", colleagues); //陪同人员
            pairs.Add("reception_employer", reception_employer); //接待单位
            pairs.Add("reception_people", reception_people); //接待人员
            pairs.Add("target_place", target_place); //用餐接待地点
            pairs.Add("meal_time", meal_time); //用餐接待地点
            pairs.Add("reason", reason);
            pairs.Add("total_count", total_count);
            pairs.Add("tax_count", tax_count);
            pairs.Add("tax_paper_number", tax_paper_number); //发票张数
            pairs.Add("start_date", start_date); //招待日期
            pairs.Add("reimbursement_date", reimbursement_date); //报销日期

            pairs.Add("start_place", "石泉");
            var paper_number = Convert.ToInt32(tax_paper_number) * 2 + 1; //发票张数需要查验，故乘2;加1表示添加一张接待清单
            if (have_wine_paper)
            {
                ++paper_number;
            }
            pairs.Add("paper_number", paper_number.ToString());
            var desktop_dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var path = System.IO.Path.Combine(desktop_dir, String.Format("运筹部{0}+业务招待+{1}元", name, total_count));
            if (!System.IO.Directory.Exists(path))
            {
                var dic_info = new System.IO.DirectoryInfo(path);
                dic_info.Create();
            }

            var cover = new Reception.Cover();
            var cover_file_name = System.IO.Path.Combine(path, "封面.xlsx");
            cover.New(cover_file_name, pairs);

            var request = new Reception.Request();
            var request_file_name = System.IO.Path.Combine(path, "审批单.xlsx");
            request.New(request_file_name, pairs);

            if (!System.IO.File.Exists(list_file_name))
            {
                var list = new Reception.RecpList();
                list.New(list_file_name, pairs);
            }
            Reception.RecpList.Append(list_file_name, pairs);
        }

        static void Generate_trip(string name, string start_place, string target_place, string reason, string colleagues, string driver,
            string start_date, string end_date, string reimbursement_date, string total_count, string tax_count, string tax_paper_number, string other_paper_number)
        {
            var pairs = new Dictionary<String, String>();
            pairs.Add("name", name);
            pairs.Add("colleagues", colleagues); //同行人员
            pairs.Add("driver", driver); //司机,若不带公车为空
            pairs.Add("start_place", start_place); //出差起始地点
            pairs.Add("target_place", target_place); //出差地点
            pairs.Add("reason", reason);
            pairs.Add("total_count", total_count);
            pairs.Add("tax_count", tax_count);
            pairs.Add("tax_paper_number", tax_paper_number); //发票张数
            pairs.Add("start_date", start_date); //出差起始日期
            pairs.Add("end_date", end_date); //出差截至日期
            pairs.Add("reimbursement_date", reimbursement_date); //报销日期

            var paper_number = Convert.ToInt32(tax_paper_number) * 2 + Convert.ToInt32(other_paper_number);
            pairs.Add("paper_number", paper_number.ToString());

            var desktop_dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var path = System.IO.Path.Combine(desktop_dir, String.Format("运筹部{0}+公务出差+{1}元", name, total_count)); //TODO: 计算出准确的人
            if (!System.IO.Directory.Exists(path))
            {
                var dic_info = new System.IO.DirectoryInfo(path);
                dic_info.Create();
            }

            var cover = new Trip.Cover();
            var cover_file_name = System.IO.Path.Combine(path, "封面.xlsx");
            cover.New(cover_file_name, pairs);

            var request = new Trip.Request();
            var request_file_name = System.IO.Path.Combine(path, "审批单.xlsx");
            request.New(request_file_name, pairs);
        }
    }
}
