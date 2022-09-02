using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

// 业务招待
namespace Reception
{
    // 业务招待封面
    public class Cover : Book
    {
        override internal void Init(HSSFWorkbook book)
        {
            var sheet = book.CreateSheet("费用报销单");
            // A4纸大小
            sheet.PrintSetup.PaperSize = 9;
            // 设置打印横向居中
            sheet.VerticallyCenter = true;
            sheet.HorizontallyCenter = true;
            sheet.PrintSetup.Landscape = true;

            var margin_scale = 2.54;
            sheet.SetMargin(MarginType.TopMargin, 1.90 / margin_scale);
            sheet.SetMargin(MarginType.BottomMargin, 1.90 / margin_scale);
            sheet.SetMargin(MarginType.LeftMargin, 2.50 / margin_scale);
            sheet.SetMargin(MarginType.RightMargin, 1.80 / margin_scale);
            sheet.SetMargin(MarginType.HeaderMargin, 0.80 / margin_scale);
            sheet.SetMargin(MarginType.FooterMargin, 0.80 / margin_scale);

            Set_row_heights(new double[] { 42.75, 28.5, 27, 35.25, 28.5, 28.5, 28.5, 28.5, 28.5, 28.5, 28.5, 28.5, 28.5, 28.5, 13.5, 21.75, 21.75 });
            Set_col_widths(new double[] { 21.71, 15, 22, 19.5, 17.1, 3.88, 19.95, 4.26 }, 265);
        }
        override internal Dictionary<RangeCell, HSSFRichTextString> Get_cell_informations(HSSFWorkbook book, Dictionary<String, String> pairs)
        {
            var wrap_style = Get_form_style(book);
            wrap_style.WrapText = true;
            var underline_font = book.CreateFont();
            underline_font.FontName = "宋体";
            underline_font.FontHeightInPoints = 12;
            underline_font.Underline = FontUnderlineType.Single;
            var vertical_style = Get_normal_style(book);
            vertical_style.Alignment = HorizontalAlignment.Center;
            vertical_style.VerticalAlignment = VerticalAlignment.Center;
            vertical_style.Rotation = (short)0xff;

            DateTime reimbursement_date;
            DateTime.TryParse(pairs.GetValueOrDefault("reimbursement_date"), out reimbursement_date);
            var total_count = Convert.ToDouble(pairs.GetValueOrDefault("total_count"));
            var tax_count = Convert.ToDouble(pairs.GetValueOrDefault("tax_count"));
            var paper_number_array = "零壹贰叁肆伍陆柒捌玖";
            var paper_number = paper_number_array[Convert.ToInt32(pairs.GetValueOrDefault("paper_number"))];
            var tax_paper_number = paper_number_array[Convert.ToInt32(pairs.GetValueOrDefault("tax_paper_number"))];
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
                    { new RangeCell(0, 0, 0, 6, Get_title_style(book)), new HSSFRichTextString("陕西葛洲坝延黄宁石高速公路有限公司") },
                    { new RangeCell(1, 0, 1, 6, Get_subtitle_style(book)), new HSSFRichTextString("费    用    报    销    单") },
                    { new RangeCell(2, 0, 2, 0, Get_small_style(book)), new HSSFRichTextString("报销单位（盖章）：") },
                    { new RangeCell(2, 1, 2, 1, Get_small_style(book)), new HSSFRichTextString("运营筹备部") },
                    { new RangeCell(2, 2, 2, 3, Get_small_style(book)), new HSSFRichTextString(String.Format("{0} 年 {1} 月 {2} 日", reimbursement_date.Year, reimbursement_date.Month, reimbursement_date.Day)) },
                    { new RangeCell(2, 6, 2, 6, Get_small_style(book)), new HSSFRichTextString("第 1 页 共 1 页") },
                    { new RangeCell(3, 0, 3, 0, Get_form_style(book)), new HSSFRichTextString("项   目") },
                    { new RangeCell(3, 1, 3, 1, Get_form_style(book)), new HSSFRichTextString("单 据 张 数") },
                    { new RangeCell(3, 2, 3, 2, Get_form_style(book)), new HSSFRichTextString("金      额") },
                    { new RangeCell(3, 3, 3, 3, Get_form_style(book)), new HSSFRichTextString("项      目") },
                    { new RangeCell(3, 4, 3, 4, Get_form_style(book)), new HSSFRichTextString("单 据 张 数") },
                    { new RangeCell(3, 5, 3, 6, Get_form_style(book)), new HSSFRichTextString("金      额") },
                    { new RangeCell(4, 0, 4, 0, Get_form_style(book)), new HSSFRichTextString("办  公  费") },
                    { new RangeCell(4, 3, 4, 3, Get_form_style(book)), new HSSFRichTextString("固 定 资 产") },
                    { new RangeCell(5, 0, 5, 0, Get_form_style(book)), new HSSFRichTextString("业务招待费") },
                    { new RangeCell(5, 1, 5, 1, Get_form_style(book)), new HSSFRichTextString(tax_paper_number.ToString()) },
                    { new RangeCell(5, 2, 5, 2, Get_form_style(book)), new HSSFRichTextString(string_total_count) },
                    { new RangeCell(5, 3, 5, 3, Get_form_style(book)), new HSSFRichTextString("税      金") },
                    { new RangeCell(6, 0, 6, 0, Get_form_style(book)), new HSSFRichTextString("市内交通费") },
                    { new RangeCell(6, 3, 6, 3, Get_form_style(book)), new HSSFRichTextString("保  险  费") },
                    { new RangeCell(7, 0, 7, 0, Get_form_style(book)), new HSSFRichTextString("会  议  费") },
                    { new RangeCell(7, 3, 7, 3, Get_form_style(book)), new HSSFRichTextString("修  理  费") },
                    { new RangeCell(8, 0, 8, 0, Get_form_style(book)), new HSSFRichTextString("培 训 费") },
                    { new RangeCell(8, 3, 8, 3, Get_form_style(book)), new HSSFRichTextString("审 计 费") },
                    { new RangeCell(9, 0, 9, 0, Get_form_style(book)), new HSSFRichTextString("福  利  费") },
                    { new RangeCell(9, 3, 9, 3, Get_form_style(book)), new HSSFRichTextString("协 会 会 费") },
                    { new RangeCell(10, 0, 10, 0, Get_form_style(book)), new HSSFRichTextString("水  电  费") },
                    { new RangeCell(10, 3, 10, 3, Get_form_style(book)), new HSSFRichTextString("车辆运行费") },
                    { new RangeCell(11, 0, 11, 0, Get_form_style(book)), new HSSFRichTextString("租      金") },
                    { new RangeCell(11, 3, 11, 3, Get_form_style(book)), new HSSFRichTextString("其    他") },
                    { new RangeCell(12, 0, 12, 1, Get_form_style(book)), new HSSFRichTextString("单 据 金 额 合 计") },
                    { new RangeCell(12, 2, 12, 6, Get_form_style(book)), richstring_count },
                    { new RangeCell(13, 0, 13, 1, Get_form_style(book)), new HSSFRichTextString("审 核 金 额") },
                    { new RangeCell(13, 2, 13, 6, Get_form_style(book)), new HSSFRichTextString("仟    佰    拾    万    仟    佰    拾    元    角   分 ￥： ________元   ") },
                    { new RangeCell(15, 0, 15, 0, Get_left_font_style(book)), new HSSFRichTextString("总经理（授权委托人）") },
                    { new RangeCell(16, 0, 16, 0, Get_left_font_style(book)), new HSSFRichTextString(" 审批：") },
                    { new RangeCell(15, 1, 15, 1, Get_normal_style(book)), new HSSFRichTextString("总会计师") },
                    { new RangeCell(16, 1, 16, 1, Get_left_font_style(book)), new HSSFRichTextString("     审 签：") },
                    { new RangeCell(15, 2, 16, 2, Get_right_font_style(book)), new HSSFRichTextString("财务部复核：  ") },
                    { new RangeCell(15, 3, 15, 3, Get_right_font_style(book)), new HSSFRichTextString("分管领导") },
                    { new RangeCell(16, 3, 16, 3, Get_right_font_style(book)), new HSSFRichTextString("审 核：") },
                    { new RangeCell(15, 4, 16, 5, Get_right_font_style(book)), new HSSFRichTextString("部门负责人：") },
                    { new RangeCell(15, 6, 16, 6, Get_right_font_style(book)), new HSSFRichTextString("报销人") },
                    { new RangeCell(3, 7, 13, 7, vertical_style), new HSSFRichTextString(String.Format("附单据  {0}  张", paper_number)) },
                };
        }
        internal override void Change_style(HSSFWorkbook book)
        {
            Fill_cell_by_style(3, 0, 11, 4, Get_form_style(book));
            for (int row = 3; row <= 11; ++row)
            {
                Fill_cell_by_style(row, 5, row, 6, Get_form_style(book), true);
            }
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
        static ICellStyle Get_small_style(HSSFWorkbook book)
        {
            var style = Get_normal_style(book);
            var font = style.GetFont(book);
            font.FontHeightInPoints = 11;
            style.SetFont(font);

            return style;
        }
        static ICellStyle Get_left_font_style(HSSFWorkbook book)
        {
            var style = Get_normal_style(book);
            style.Alignment = HorizontalAlignment.Left;

            return style;
        }
        static ICellStyle Get_right_font_style(HSSFWorkbook book)
        {
            var style = Get_normal_style(book);
            style.Alignment = HorizontalAlignment.Right;

            return style;
        }
    }

    // 业务接待审批单
    public class Request : Book
    {
        internal override void Init(HSSFWorkbook book)
        {
            var sheet = book.CreateSheet("业务接待审批单");
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

            Set_row_heights(new double[] { 13.5, 13.5, 20.25, 13.5, 30.75, 13.5, 13.5, 52.5, 56.65, 56.65, 28, 28, 59.5, 70.85, 88.55, 13.5, 20.25 });
            Set_col_widths(new double[] { 22.05, 17.33, 17.33, 17.33, 17.33 }, 270);
        }
        internal override Dictionary<RangeCell, HSSFRichTextString> Get_cell_informations(HSSFWorkbook book, Dictionary<string, string> pairs)
        {
            var handwrite_font = book.CreateFont();
            handwrite_font.FontName = "集萤映雪";
            handwrite_font.FontHeightInPoints = 14;
            var normal_font = book.CreateFont();
            normal_font.FontName = "仿宋_GB2312";
            normal_font.FontHeightInPoints = 14;

            var details = "";
            DateTime start_date;
            DateTime.TryParse(pairs.GetValueOrDefault("start_date"), out start_date);
            start_date = start_date.AddDays(-1);
            var note_date_string = String.Format("填表日期：{0} 年 {1} 月 {2} 日",
                start_date.Year, start_date.Month, start_date.Day);
            var note_date_richstring = new HSSFRichTextString(note_date_string);
            note_date_richstring.ApplyFont(handwrite_font);
            var regex = new Regex(".*(填表日期：).*(年).*(月).*(日).*");
            var groups = regex.Match(note_date_string).Groups;
            for (int i = 1; i <= groups.Count; ++i)
            {
                var tmp = groups[i];
                note_date_richstring.ApplyFont(tmp.Index, tmp.Index + tmp.Length, normal_font);
            }

            return new Dictionary<RangeCell, HSSFRichTextString>
            {
                    { new RangeCell(2, 0, 2, 0, Get_header_style(book)), new HSSFRichTextString("  附件 1") },
                    { new RangeCell(4, 0, 4, 3, Get_title_style(book)), new HSSFRichTextString("业务接待审批单") },
                    { new RangeCell(7, 0, 7, 0, Get_form_style(book)), new HSSFRichTextString("责任部门") },
                    { new RangeCell(7, 1, 7, 1, Get_form_style(book)), new HSSFRichTextString("运营筹备部")},
                    { new RangeCell(7, 2, 7, 2, Get_form_style(book)), new HSSFRichTextString("经办人")},
                    { new RangeCell(7, 3, 7, 3, Get_form_style(book)), new HSSFRichTextString(pairs.GetValueOrDefault("name"))},
                    { new RangeCell(8, 0, 8, 0, Get_form_style(book)), new HSSFRichTextString("接待事由") },
                    { new RangeCell(8, 1, 8, 3, Get_form_style(book)),new HSSFRichTextString(pairs.GetValueOrDefault("reason"))},
                    { new RangeCell(9, 0, 9, 0, Get_form_style(book)), new HSSFRichTextString("接待（拜访）单位及人数") },
                    { new RangeCell(9, 1, 9, 3, Get_form_style(book)), new HSSFRichTextString(details) },
                    { new RangeCell(10, 0, 11, 0, Get_form_style(book)), new HSSFRichTextString("接待类别") },
                    { new RangeCell(10, 1, 10, 3, Get_form_style(book)), new HSSFRichTextString("□重要商务接待       □一般商务接待") },
                    { new RangeCell(11, 1, 11, 3, Get_form_style(book)), new HSSFRichTextString("□重要公务接待       □一般公务接待") },
                    { new RangeCell(12, 0, 12, 0, Get_form_style(book)), new HSSFRichTextString("办公室意见") },
                    { new RangeCell(12, 1, 12, 3, Get_form_style(book)), new HSSFRichTextString("") },
                    { new RangeCell(13, 0, 13, 0, Get_form_style(book)), new HSSFRichTextString("业务分管领导审批") },
                    { new RangeCell(13, 1, 13, 3, Get_form_style(book)), new HSSFRichTextString("") },
                    { new RangeCell(14, 0, 14, 0, Get_form_style(book)), new HSSFRichTextString("主要领导审批") },
                    { new RangeCell(14, 1, 14, 3, Get_form_style(book)), new HSSFRichTextString("") },

                    { new RangeCell(16, 1, 16, 3, Get_note_style(book)), note_date_richstring },
            };
        }
        internal override void Change_style(HSSFWorkbook book)
        {
            Fill_cell_by_style(10, 1, 11, 3, Get_form_style(book), true);
            Fill_cell_by_style(7, 0, 14, 3, Get_form_style(book, BorderStyle.Medium), true, true);
        }
        static ICellStyle Get_header_style(HSSFWorkbook book)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 16;
            title_font.FontName = "黑体";
            style.SetFont(title_font);
            style.Alignment = HorizontalAlignment.Left;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
        static ICellStyle Get_title_style(HSSFWorkbook book)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 22;
            title_font.FontName = "小标宋";
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
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
        static ICellStyle Get_form_style(HSSFWorkbook book, BorderStyle borderStyle = BorderStyle.Thin)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 14;
            title_font.FontName = "仿宋_GB2312";
            style.SetFont(title_font);
            style.WrapText = true;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.BorderTop = borderStyle;
            style.BorderBottom = borderStyle;
            style.BorderLeft = borderStyle;
            style.BorderRight = borderStyle;

            return style;
        }
    }

    // 业务接待清单
    class RecpList : Book
    {
        // Used for Method append;
        private HSSFWorkbook book;
        internal override void Init(HSSFWorkbook book)
        {
            var sheet = book.CreateSheet("业务接待清单");
            // A4纸大小
            sheet.PrintSetup.PaperSize = 9;
            // 设置打印横向居中
            sheet.HorizontallyCenter = true;
            sheet.PrintSetup.Landscape = true;

            // 设置打印区域
            book.SetPrintArea(0, 0, 8, 0, 5);

            var margin_scale = 2.54;
            sheet.SetMargin(MarginType.TopMargin, 2.54 / margin_scale);
            sheet.SetMargin(MarginType.BottomMargin, 2.54 / margin_scale);
            sheet.SetMargin(MarginType.LeftMargin, 3.50 / margin_scale);
            sheet.SetMargin(MarginType.RightMargin, 1.91 / margin_scale);
            sheet.SetMargin(MarginType.HeaderMargin, 1.27 / margin_scale);
            sheet.SetMargin(MarginType.FooterMargin, 1.27 / margin_scale);

            this.book = book;
            Set_row_heights(new double[] { 20.25, 28.5, 38, 91, 45, 37, 13.5, 13.5, 13.5, 28.5, 38 });
            Set_col_widths(new double[] { 11.37, 8.38, 8.38, 10.96, 18.43, 18.98, 10.68, 12.86, 12.99 }, 270);
        }
        static public void Append(String file_name, Dictionary<string, string> pairs)
        {
            HSSFWorkbook book;
            using (var fs = new FileStream(file_name, FileMode.Open, FileAccess.ReadWrite))
            {
                book = new HSSFWorkbook(fs);

                var sheet = book.GetSheetAt(0);
                var row_num = sheet.LastRowNum;
                var row = sheet.CreateRow(row_num + 1);
                var target_row = sheet.GetRow(3);
                row.Height = 1340;

                // 复用generate
                var key_array = Parse_pairs(pairs);
                for (int id = 0; id < key_array.Length; ++id)
                {
                    var cell = row.CreateCell(id);
                    var target_cell = target_row.GetCell(id);
                    cell.SetCellValue(key_array[id]);
                    target_cell.SetCellValue(key_array[id]);
                    cell.CellStyle = Get_form_style(book);
                }
            }
            using (var fs = new FileStream(file_name, FileMode.Open, FileAccess.Write))
            {
                book.Write(fs);
            }
        }
        internal override Dictionary<RangeCell, HSSFRichTextString> Get_cell_informations(HSSFWorkbook book, Dictionary<string, string> pairs)
        {
            var key_array = Parse_pairs(pairs);

            return new Dictionary<RangeCell, HSSFRichTextString>
            {
                    { new RangeCell(1, 0, 1, 8, Get_title_style(book)), new HSSFRichTextString("业务接待清单") },
                    { new RangeCell(2, 0, 2, 0, Get_form_style(book)), new HSSFRichTextString("日期") },
                    { new RangeCell(2, 1, 2, 1, Get_form_style(book)), new HSSFRichTextString("早/中/晚") },
                    { new RangeCell(2, 2, 2, 2, Get_form_style(book)), new HSSFRichTextString("用餐地点") },
                    { new RangeCell(2, 3, 2, 3, Get_form_style(book)), new HSSFRichTextString("接待（拜访）单位") },
                    { new RangeCell(2, 4, 2, 4, Get_form_style(book)), new HSSFRichTextString("我方人员（陪同人员）") },
                    { new RangeCell(2, 5, 2, 5, Get_form_style(book)), new HSSFRichTextString("对方人员（接待人员）") },
                    { new RangeCell(2, 6, 2, 6, Get_form_style(book)), new HSSFRichTextString("费用（元）") },
                    { new RangeCell(2, 7, 2, 7, Get_form_style(book)), new HSSFRichTextString("报销日期") },
                    { new RangeCell(2, 8, 2, 8, Get_form_style(book)), new HSSFRichTextString("备注") },

                    { new RangeCell(3, 0, 3, 0, Get_form_style(book)), new HSSFRichTextString(key_array[0]) },
                    { new RangeCell(3, 1, 3, 1, Get_form_style(book)), new HSSFRichTextString(key_array[1]) },
                    { new RangeCell(3, 2, 3, 2, Get_form_style(book)), new HSSFRichTextString(key_array[2]) },
                    { new RangeCell(3, 3, 3, 3, Get_form_style(book)), new HSSFRichTextString(key_array[3]) },
                    { new RangeCell(3, 4, 3, 4, Get_form_style(book)), new HSSFRichTextString(key_array[4]) },
                    { new RangeCell(3, 5, 3, 5, Get_form_style(book)), new HSSFRichTextString(key_array[5]) },
                    { new RangeCell(3, 6, 3, 6, Get_form_style(book)), new HSSFRichTextString(key_array[6]) },
                    { new RangeCell(3, 7, 3, 7, Get_form_style(book)), new HSSFRichTextString(key_array[7]) },
                    { new RangeCell(3, 8, 3, 8, Get_form_style(book)), new HSSFRichTextString() },

                    { new RangeCell(9, 0, 9, 8, Get_title_style(book)), new HSSFRichTextString("业务接待清单台账（之前填过的复制到下面，不影响打印）") },
                    { new RangeCell(10, 0, 10, 0, Get_form_style(book)), new HSSFRichTextString("日期") },
                    { new RangeCell(10, 1, 10, 1, Get_form_style(book)), new HSSFRichTextString("早/中/晚") },
                    { new RangeCell(10, 2, 10, 2, Get_form_style(book)), new HSSFRichTextString("用餐地点") },
                    { new RangeCell(10, 3, 10, 3, Get_form_style(book)), new HSSFRichTextString("接待（拜访）单位") },
                    { new RangeCell(10, 4, 10, 4, Get_form_style(book)), new HSSFRichTextString("我方人员（陪同人员）") },
                    { new RangeCell(10, 5, 10, 5, Get_form_style(book)), new HSSFRichTextString("对方人员（接待人员）") },
                    { new RangeCell(10, 6, 10, 6, Get_form_style(book)), new HSSFRichTextString("费用（元）") },
                    { new RangeCell(10, 7, 10, 7, Get_form_style(book)), new HSSFRichTextString("报销日期") },
                    { new RangeCell(10, 8, 10, 8, Get_form_style(book)), new HSSFRichTextString("备注") },
            };
        }
        internal override void Change_style(HSSFWorkbook book)
        {
            Fill_cell_by_style(4, 0, 5, 8, Get_form_style(book));
        }
        static private String[] Parse_pairs(Dictionary<string, string> pairs)
        {
            var key_array = new String[9];
            var name = pairs.GetValueOrDefault("name");
            var colleagues = pairs.GetValueOrDefault("colleagues");
            var reception_people = pairs.GetValueOrDefault("reception_people");
            DateTime start_date, reimbursement_date;
            DateTime.TryParse(pairs.GetValueOrDefault("start_date"), out start_date);
            DateTime.TryParse(pairs.GetValueOrDefault("reimbursement_date"), out reimbursement_date);

            String Join_name(String[] names)
            {
                var _name = names[0];
                for (int id = 1; id < names.Length; ++id)
                {
                    if (id % 2 != 0)
                    {
                        _name += '、' + names[id];
                    }
                    else
                    {
                        _name += "、\n" + names[id];
                    }
                }
                return _name;
            }
            var our_people_list = String.Join(' ', name, colleagues).Split(new char[] { ' ', ',', '，', '\t', ';', '；', '、' }, StringSplitOptions.RemoveEmptyEntries);
            var our_people = Join_name(our_people_list);
            var other_people_list = reception_people.Split(new char[] { ' ', ',', '，', '\t', ';', '；', '、' }, StringSplitOptions.RemoveEmptyEntries);
            var other_people = Join_name(other_people_list);

            key_array[0] = String.Format("{0}.{1}.{2}", start_date.Year, start_date.Month, start_date.Day);
            key_array[1] = pairs.GetValueOrDefault("meal_time");
            key_array[2] = pairs.GetValueOrDefault("target_place");
            key_array[3] = pairs.GetValueOrDefault("reception_employer");
            key_array[4] = our_people;
            key_array[5] = other_people;
            key_array[6] = pairs.GetValueOrDefault("total_count");
            key_array[7] = String.Format("{0}.{1}.{2}", reimbursement_date.Year, reimbursement_date.Month, reimbursement_date.Day);

            return key_array;
        }
        static ICellStyle Get_title_style(HSSFWorkbook book)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 22;
            title_font.FontName = "小标宋";
            style.SetFont(title_font);
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            return style;
        }
        static ICellStyle Get_form_style(HSSFWorkbook book, BorderStyle borderStyle = BorderStyle.Thin)
        {
            var style = book.CreateCellStyle();
            var title_font = book.CreateFont();
            title_font.FontHeightInPoints = 12;
            title_font.FontName = "仿宋_GB2312";
            style.SetFont(title_font);
            style.WrapText = true;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.BorderTop = borderStyle;
            style.BorderBottom = borderStyle;
            style.BorderLeft = borderStyle;
            style.BorderRight = borderStyle;

            return style;
        }
    }
}
