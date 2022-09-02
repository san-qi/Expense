using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

// 公务出差
namespace Trip
{
    // 公务出差封面
    public class Cover : Book
    {
        override internal void Init(HSSFWorkbook book)
        {
            var sheet = book.CreateSheet("差旅费报销单（本币）");
            // A4纸大小
            sheet.PrintSetup.PaperSize = 9;
            // 设置打印横向居中
            sheet.VerticallyCenter = true;
            sheet.HorizontallyCenter = true;
            sheet.PrintSetup.Landscape = true;

            var margin_scale = 2.54;
            sheet.SetMargin(MarginType.TopMargin, 1.50 / margin_scale);
            sheet.SetMargin(MarginType.BottomMargin, 1.50 / margin_scale);
            sheet.SetMargin(MarginType.LeftMargin, 2.50 / margin_scale);
            sheet.SetMargin(MarginType.RightMargin, 1.80 / margin_scale);
            sheet.SetMargin(MarginType.HeaderMargin, 0.80 / margin_scale);
            sheet.SetMargin(MarginType.FooterMargin, 0.80 / margin_scale);

            Set_row_heights(new double[] { 37.5, 31.5, 14.25, 25.5, 25.5, 38, 25.5, 25.5, 25.5, 28, 25.5, 28, 25.5, 25.5, 25.5, 25.5, 25.5, 25.5, 25.5, 25.5, 14.25, 21, 21 });
            Set_col_widths(new double[] { 4.76, 4.76, 10.51, 5.13, 5.38, 10.63, 10.26, 10.26, 5.26, 9.51, 6.01, 9.38, 11.38, 11.63, 4.38 });
        }
        internal override void Change_style(HSSFWorkbook book)
        {
            Fill_cell_by_style(8, 0, 15, 13, Get_form_style(book));
        }
        override internal Dictionary<RangeCell, HSSFRichTextString> Get_cell_informations(HSSFWorkbook book, Dictionary<String, String> pairs)
        {
            var sheet = book.GetSheetAt(0);
            var wrap_style = Get_form_style(book);
            wrap_style.WrapText = true;
            var underline_font = book.CreateFont();
            underline_font.Underline = FontUnderlineType.Single;
            underline_font.FontName = "宋体";
            underline_font.FontHeightInPoints = 12;
            var vertical_style = Get_normal_style(book);
            vertical_style.Alignment = HorizontalAlignment.Center;
            vertical_style.VerticalAlignment = VerticalAlignment.Center;
            vertical_style.Rotation = (short)0xff;

            DateTime start_date, end_date, reimbursement_date;
            DateTime.TryParse(pairs.GetValueOrDefault("start_date"), out start_date);
            DateTime.TryParse(pairs.GetValueOrDefault("end_date"), out end_date);
            DateTime.TryParse(pairs.GetValueOrDefault("reimbursement_date"), out reimbursement_date);
            var name = pairs.GetValueOrDefault("name");
            var start_place = pairs.GetValueOrDefault("start_place");
            var target_place = pairs.GetValueOrDefault("target_place");
            var reason = pairs.GetValueOrDefault("reason");
            var total_count = Convert.ToDouble(pairs.GetValueOrDefault("total_count"));
            var tax_count = Convert.ToDouble(pairs.GetValueOrDefault("tax_count"));
            var paper_number_array = "零壹贰叁肆伍陆柒捌玖";
            var paper_number = paper_number_array[Convert.ToInt32(pairs.GetValueOrDefault("paper_number"))];
            var remainder_count = total_count - tax_count;
            var string_total_count = String.Format("{0:N}", total_count);
            var string_tax_count = String.Format("{0:N}", tax_count);
            var string_remainder_count = String.Format("{0:N}", remainder_count);
            var string_count = String.Format("￥：  {0}  元（其中：增值税额  {1}  元，实际费用  {2}  元）", string_total_count, string_tax_count, string_remainder_count);
            var richstring_count = new HSSFRichTextString(string_count);
            richstring_count.ApplyFont(string_count.IndexOf(string_total_count) - 2, string_count.IndexOf(string_total_count) + string_total_count.Length + 2, underline_font);
            richstring_count.ApplyFont(string_count.IndexOf(string_tax_count) - 2, string_count.IndexOf(string_tax_count) + string_tax_count.Length + 2, underline_font);
            richstring_count.ApplyFont(string_count.LastIndexOf(string_remainder_count) - 2, string_count.LastIndexOf(string_remainder_count) + string_remainder_count.Length + 2, underline_font);

            return new Dictionary<RangeCell, HSSFRichTextString>
                {
                    { new RangeCell(0, 0, 0, 13, Get_title_style(book)), new HSSFRichTextString("陕西葛洲坝延黄宁石高速公路有限公司") },
                    { new RangeCell(1, 0, 1, 13, Get_subtitle_style(book)), new HSSFRichTextString("差 旅 费 报 销 单") },
                    { new RangeCell(3, 0, 3, 4, Get_normal_style(book)), new HSSFRichTextString("报销单位（盖章）：运营筹备部") },
                    { new RangeCell(3, 6, 3, 8, Get_normal_style(book)), new HSSFRichTextString(String.Format("{0} 年 {1} 月 {2} 日", reimbursement_date.Year, reimbursement_date.Month, reimbursement_date.Day)) },
                    { new RangeCell(3, 12, 3, 13, Get_normal_style(book)), new HSSFRichTextString("      第 1 页 共 1 页") },
                    { new RangeCell(4, 0, 4, 2, Get_form_style(book)), new HSSFRichTextString("姓     名") },
                    { new RangeCell(4, 3, 4, 5, Get_form_style(book)), new HSSFRichTextString("出 差 地 点") },
                    { new RangeCell(4, 6, 4, 11, Get_form_style(book)), new HSSFRichTextString("出  差  事  由") },
                    { new RangeCell(4, 12, 4, 13, Get_form_style(book)), new HSSFRichTextString("预借备用金") },
                    { new RangeCell(5, 0, 5, 2, Get_form_style(book)), new HSSFRichTextString(name) },
                    { new RangeCell(5, 3, 5, 5, Get_form_style(book)), new HSSFRichTextString(target_place) },
                    { new RangeCell(5, 6, 5, 11, Get_form_style(book)), new HSSFRichTextString(reason) },
                    { new RangeCell(5, 12, 5, 13, Get_form_style(book)), new HSSFRichTextString("") },
                    { new RangeCell(6, 0, 6, 2, Get_form_style(book)), new HSSFRichTextString("起") },
                    { new RangeCell(6, 3, 6, 5, Get_form_style(book)), new HSSFRichTextString("止") },
                    { new RangeCell(6, 8, 6, 9, Get_form_style(book)), new HSSFRichTextString("出差补助") },
                    { new RangeCell(6, 10, 6, 11, Get_form_style(book)), new HSSFRichTextString("乘车补助") },
                    { new RangeCell(6, 12, 6, 13, Get_form_style(book)), new HSSFRichTextString("其他") },
                    { new RangeCell(7, 0, 7, 0, Get_form_style(book)), new HSSFRichTextString("月") },
                    { new RangeCell(7, 0, 7, 0, Get_form_style(book)), new HSSFRichTextString("月") },
                    { new RangeCell(7, 1, 7, 1, Get_form_style(book)), new HSSFRichTextString("日") },
                    { new RangeCell(7, 2, 7, 2, Get_form_style(book)), new HSSFRichTextString("地点") },
                    { new RangeCell(7, 3, 7, 3, Get_form_style(book)), new HSSFRichTextString("月") },
                    { new RangeCell(7, 4, 7, 4, Get_form_style(book)), new HSSFRichTextString("日") },
                    { new RangeCell(7, 5, 7, 5, Get_form_style(book)), new HSSFRichTextString("地点") },
                    { new RangeCell(6, 6, 7, 6, Get_form_style(book)), new HSSFRichTextString("路 费") },
                    { new RangeCell(6, 7, 7, 7, Get_form_style(book)), new HSSFRichTextString("住宿费") },
                    { new RangeCell(7, 8, 7, 8, Get_form_style(book)), new HSSFRichTextString("人/天") },
                    { new RangeCell(7, 9, 7, 9, Get_form_style(book)), new HSSFRichTextString("金 额") },
                    { new RangeCell(7, 10, 7, 10, Get_form_style(book)), new HSSFRichTextString("小时") },
                    { new RangeCell(7, 11, 7, 11, Get_form_style(book)), new HSSFRichTextString("金 额") },
                    { new RangeCell(7, 12, 7, 12, Get_form_style(book)), new HSSFRichTextString("项  目") },
                    { new RangeCell(7, 13, 7, 13, Get_form_style(book)), new HSSFRichTextString("金  额") },
                    { new RangeCell(8, 0, 8, 0, Get_form_style(book)), new HSSFRichTextString(start_date.Month.ToString()) },
                    { new RangeCell(8, 1, 8, 1, Get_form_style(book)), new HSSFRichTextString(start_date.Day.ToString()) },
                    { new RangeCell(8, 2, 8, 2, Get_form_style(book)), new HSSFRichTextString(start_place) },
                    { new RangeCell(8, 3, 8, 3, Get_form_style(book)), new HSSFRichTextString(start_date.Month.ToString()) },
                    { new RangeCell(8, 4, 8, 4, Get_form_style(book)), new HSSFRichTextString(start_date.Day.ToString()) },
                    { new RangeCell(8, 5, 8, 5, Get_form_style(book)), new HSSFRichTextString(target_place) },
                    { new RangeCell(8, 12, 8, 12, Get_form_style(book)), new HSSFRichTextString("交通保险费") },
                    { new RangeCell(9, 0, 9, 0, Get_form_style(book)), new HSSFRichTextString(end_date.Month.ToString()) },
                    { new RangeCell(9, 1, 9, 1, Get_form_style(book)), new HSSFRichTextString(end_date.Day.ToString()) },
                    { new RangeCell(9, 2, 9, 2, Get_form_style(book)), new HSSFRichTextString(target_place) },
                    { new RangeCell(9, 3, 9, 3, Get_form_style(book)), new HSSFRichTextString(end_date.Month.ToString()) },
                    { new RangeCell(9, 4, 9, 4, Get_form_style(book)), new HSSFRichTextString(end_date.Day.ToString()) },
                    { new RangeCell(9, 5, 9, 5, Get_form_style(book)), new HSSFRichTextString(start_place) },
                    { new RangeCell(9, 12, 9, 12, wrap_style), new HSSFRichTextString("订票、退票\n手续费") },
                    { new RangeCell(10, 12, 10, 12, Get_form_style(book)), new HSSFRichTextString("邮寄费") },
                    { new RangeCell(11, 12, 11, 12, wrap_style), new HSSFRichTextString("差旅市内\n交通费") },
                    { new RangeCell(12, 12, 12, 12, Get_form_style(book)), new HSSFRichTextString("会 议 费") },
                    { new RangeCell(13, 12, 13, 12, Get_form_style(book)), new HSSFRichTextString("核酸检测") },
                    { new RangeCell(16, 0, 16, 5, Get_form_style(book)), new HSSFRichTextString("单 据 金 额 合 计") },
                    { new RangeCell(16, 6, 16, 13, Get_form_style(book)), richstring_count },
                    { new RangeCell(17, 0, 17, 2, Get_form_style(book)), new HSSFRichTextString("审 核 金 额") },
                    { new RangeCell(17, 3, 17, 13, Get_form_style(book)), new HSSFRichTextString("佰     拾     万     仟     佰     拾     元     角     分   ￥：_________元") },
                    { new RangeCell(19, 0, 19, 2, Get_normal_style(book)), new HSSFRichTextString("总经理（授权委托人）") },
                    { new RangeCell(20, 0, 20, 1, Get_normal_style(book)), new HSSFRichTextString("审批：") },
                    { new RangeCell(19, 3, 19, 5, Get_normal_style(book)), new HSSFRichTextString("总会计师") },
                    { new RangeCell(20, 3, 20, 5, Get_normal_style(book)), new HSSFRichTextString("审签：") },
                    { new RangeCell(19, 6, 20, 7, Get_normal_style(book)), new HSSFRichTextString("财务部复核：") },
                    { new RangeCell(19, 8, 19, 9, Get_normal_style(book)), new HSSFRichTextString("分管领导") },
                    { new RangeCell(20, 8, 20, 9, Get_normal_style(book)), new HSSFRichTextString("审 核：") },
                    { new RangeCell(19, 11, 20, 12, Get_normal_style(book)), new HSSFRichTextString("部门负责人：") },
                    { new RangeCell(19, 13, 20, 13, Get_normal_style(book)), new HSSFRichTextString("报销人") },
                    { new RangeCell(4, 14, 14, 14, vertical_style), new HSSFRichTextString(String.Format("附单据  {0}  张", paper_number)) },
                };
        }
        static ICellStyle Get_title_style(HSSFWorkbook book)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 26;
            title_font.FontName = "华文中宋";
            title_font.IsBold = true;
            style.SetFont(title_font);
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
        static ICellStyle Get_subtitle_style(HSSFWorkbook book)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 22;
            title_font.FontName = "华文楷体";
            title_font.IsBold = true;
            title_font.Underline = FontUnderlineType.Double;
            style.SetFont(title_font);
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
        static ICellStyle Get_form_style(HSSFWorkbook book)
        {
            var style = Get_normal_style(book);
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.WrapText = true;

            return style;
        }
        static ICellStyle Get_normal_style(HSSFWorkbook book)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 12;
            title_font.FontName = "宋体";
            style.SetFont(title_font);
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
    }

    // 出差审批单
    public class Request : Book
    {
        internal override void Init(HSSFWorkbook book)
        {
            var sheet = book.CreateSheet("出差审批单");
            // A4纸大小
            sheet.PrintSetup.PaperSize = 9;
            // 设置打印纵向居中
            sheet.HorizontallyCenter = true;
            sheet.PrintSetup.Landscape = false;

            var margin_scale = 2.54;
            sheet.SetMargin(MarginType.TopMargin, 1.91 / margin_scale);
            sheet.SetMargin(MarginType.BottomMargin, 1.91 / margin_scale);
            sheet.SetMargin(MarginType.LeftMargin, 1.78 / margin_scale);
            sheet.SetMargin(MarginType.RightMargin, 1.78 / margin_scale);
            sheet.SetMargin(MarginType.HeaderMargin, 0.76 / margin_scale);
            sheet.SetMargin(MarginType.FooterMargin, 0.76 / margin_scale);

            Set_row_heights(new double[] { 13.5, 13.5, 13.5, 13.5, 28, 28, 13.5, 13.5, 51, 51, 51, 51, 51, 51, 51, 51, 51, 28, 28 });
            Set_col_widths(new double[] { 17.6, 13.2, 7.4, 16.1, 8.6, 16 }, 267);
        }
        internal override Dictionary<RangeCell, HSSFRichTextString> Get_cell_informations(HSSFWorkbook book, Dictionary<string, string> pairs)
        {
            var handwrite_font = book.CreateFont();
            handwrite_font.FontName = "集萤映雪";
            handwrite_font.FontHeightInPoints = 14;
            var normal_font = book.CreateFont();
            normal_font.FontName = "仿宋_GB2312";
            normal_font.FontHeightInPoints = 14;

            DateTime start_date, end_date;
            DateTime.TryParse(pairs.GetValueOrDefault("start_date"), out start_date);
            DateTime.TryParse(pairs.GetValueOrDefault("end_date"), out end_date);
            var day = end_date - start_date;
            var date_span_string = String.Format("{0}年 {1}月 {2}日 至 {3}年 {4}月 {5}日，共 {6} 天",
                start_date.Year, start_date.Month, start_date.Day, end_date.Year, end_date.Month, end_date.Day, day.TotalDays);
            var date_span_richstring = new HSSFRichTextString(date_span_string);
            date_span_richstring.ApplyFont(handwrite_font);
            var regex = new Regex(".*(年).*(月).*(日).*(至).*(年).*(月).*(日，共).*(天).*");
            var groups = regex.Match(date_span_string).Groups;
            for (int i = 1; i <= groups.Count; ++i)
            {
                var tmp = groups[i];
                date_span_richstring.ApplyFont(tmp.Index, tmp.Index + tmp.Length, normal_font);
            }

            return new Dictionary<RangeCell, HSSFRichTextString>
            {
                    { new RangeCell(4, 0, 4, 5, Get_title_style(book)), new HSSFRichTextString("陕西葛洲坝延黄宁石高速公路有限公司") },
                    { new RangeCell(5, 0, 5, 5, Get_title_style(book)), new HSSFRichTextString("出差审批单") },
                    { new RangeCell(8, 0, 8, 0, Get_form_style(book)), new HSSFRichTextString("姓名")},
                    { new RangeCell(8, 1, 8, 1, Get_form_style(book)), new HSSFRichTextString(pairs.GetValueOrDefault("name"))},
                    { new RangeCell(8, 2, 8, 2, Get_form_style(book)), new HSSFRichTextString("部门") },
                    { new RangeCell(8, 3, 8, 3, Get_form_style(book)), new HSSFRichTextString("运营筹备部") },
                    { new RangeCell(8, 4, 8, 4, Get_form_style(book)), new HSSFRichTextString("职务或岗位") },
                    { new RangeCell(8, 5, 8, 5, Get_form_style(book)), new HSSFRichTextString("") },

                    { new RangeCell(9, 0, 9, 0, Get_form_style(book)), new HSSFRichTextString("出差地点") },
                    { new RangeCell(9, 1, 9, 5, Get_base_style(book)), new HSSFRichTextString(pairs.GetValueOrDefault("target_place"))},
                    { new RangeCell(10, 0, 10, 0, Get_form_style(book)), new HSSFRichTextString("计划出差时间") },
                    { new RangeCell(10, 1, 10, 5, Get_base_style(book)), date_span_richstring },
                    { new RangeCell(11, 0, 11, 0, Get_form_style(book)), new HSSFRichTextString("交通工具") },
                    { new RangeCell(11, 1, 11, 5, Get_base_style(book)), new HSSFRichTextString("□公车  □飞机  □火车  □汽车  □动车  □其他") },
                    { new RangeCell(12, 0, 12, 0, Get_form_style(book)), new HSSFRichTextString("同行人员") },
                    // { new RangeCell(12, 1, 12, 5, Get_base_style(book)), new HSSFRichTextString(pairs.GetValueOrDefault("colleagues"))},
                    { new RangeCell(13, 0, 13, 0, Get_form_style(book)), new HSSFRichTextString("出差事由") },
                    // { new RangeCell(13, 1, 13, 5, Get_base_style(book)), new HSSFRichTextString(pairs.GetValueOrDefault("reason"))},
                    { new RangeCell(14, 0, 14, 0, Get_form_style(book)), new HSSFRichTextString("部门意见") },
                    { new RangeCell(14, 1, 14, 5, Get_base_style(book)), new HSSFRichTextString("") },
                    { new RangeCell(15, 0, 15, 0, Get_form_style(book)), new HSSFRichTextString("分管领导审核") },
                    { new RangeCell(15, 1, 15, 5, Get_base_style(book)), new HSSFRichTextString("") },
                    { new RangeCell(16, 0, 16, 0, Get_form_style(book)), new HSSFRichTextString("总经理（授权委托人）审批") },
                    { new RangeCell(16, 1, 16, 5, Get_base_style(book)), new HSSFRichTextString("") },

                    { new RangeCell(17, 0, 17, 5, Get_note_style(book)), new HSSFRichTextString(" 说明：1.此审批表作为出差申请及差旅费核销必备凭证。") },
                    { new RangeCell(18, 0, 18, 5, Get_note_style(book)), new HSSFRichTextString("       2.如出差途中需改变行程计划需及时汇报。") },
            };
        }
        internal override void Change_style(HSSFWorkbook book)
        {
            Fill_cell_by_style(9, 1, 9, 5, Get_form_style(book), true);
            Fill_cell_by_style(10, 1, 10, 5, Get_form_style(book), true);
            Fill_cell_by_style(11, 1, 11, 5, Get_form_style(book), true);
            Fill_cell_by_style(12, 1, 12, 5, Get_form_style(book), true);
            Fill_cell_by_style(13, 1, 13, 5, Get_form_style(book), true);
            Fill_cell_by_style(14, 1, 14, 5, Get_form_style(book), true);
            Fill_cell_by_style(15, 1, 15, 5, Get_form_style(book), true);
            Fill_cell_by_style(16, 1, 16, 5, Get_form_style(book), true);
            Fill_cell_by_style(8, 0, 16, 5, Get_form_style(book, BorderStyle.Medium), true, true);
        }
        static ICellStyle Get_title_style(HSSFWorkbook book)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 22;
            title_font.FontName = "华文中宋";
            style.SetFont(title_font);
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
        static ICellStyle Get_note_style(HSSFWorkbook book)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 14;
            title_font.FontName = "仿宋_GB2312";
            style.SetFont(title_font);
            style.Alignment = HorizontalAlignment.Left;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
        static ICellStyle Get_form_style(HSSFWorkbook book, BorderStyle borderStyle = BorderStyle.Thin)
        {
            var style = Get_base_style(book);
            style.BorderTop = borderStyle;
            style.BorderBottom = borderStyle;
            style.BorderLeft = borderStyle;
            style.BorderRight = borderStyle;
            style.WrapText = true;

            return style;
        }
        static ICellStyle Get_base_style(HSSFWorkbook book)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 14;
            title_font.FontName = "仿宋_GB2312";
            style.SetFont(title_font);
            style.WrapText = true;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
    }

}