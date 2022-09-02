using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

public abstract class Book
{
    private HSSFWorkbook book;
    private double[] row_heights;
    private double[] col_widths;
    public void New(String file_name, Dictionary<String, String> pairs, bool covered = true)
    {
        var file_type = covered ? FileAccess.Write : FileAccess.Read;
        using (var fs = new FileStream(file_name, FileMode.Create, file_type))
        {
            book = new HSSFWorkbook();
            Init(book);
            Init_sheet_by_shape();
            Generete(Get_cell_informations(book, pairs));
            Change_style(book);
            book.Write(fs);
        }
    }
    private void Generete(Dictionary<RangeCell, HSSFRichTextString> keyValuePairs)
    {
        var sheet = book.GetSheetAt(0);
        foreach (var keyValue in keyValuePairs)
        {
            var key = keyValue.Key;
            var cell = sheet.GetRow(key.start_x).CreateCell(key.start_y);
            if (key.Is_merge_cell())
            {
                sheet.AddMergedRegion(key.To_cell_range_address());
            }
            Fill_cell_by_style(key.start_x, key.start_y, key.end_x, key.end_y, key.style);
            if (keyValue.Value.Length != 0)
            {
                cell.SetCellValue(keyValue.Value);
            }
        }
    }
    private void Init_sheet_by_shape()
    {
        var sheet = book.GetSheetAt(0);
        for (int i = 0; i < col_widths.Length; ++i)
        {
            sheet.SetColumnWidth(i, (int)(col_widths[i]));
        }
        for (int i = 0; i < row_heights.Length; ++i)
        {
            var row = sheet.CreateRow(i);
            row.Height = (short)(row_heights[i] * 20);
            for (int j = 0; j < col_widths.Length; ++j)
            {
                row.CreateCell(j);
            }
        }
    }
    internal void Fill_cell_by_style(int start_x, int start_y, int end_x, int end_y, ICellStyle style, bool cannulated = false, bool keep_others = false)
    {
        var sheet = book.GetSheetAt(0);
        for (int i = start_x; i <= end_x; ++i)
        {
            for (int j = start_y; j <= end_y; ++j)
            {
                var form_cell = sheet.GetRow(i).GetCell(j);
                var _style = book.CreateCellStyle();
                _style.CloneStyleFrom(style);
                if (cannulated)
                {
                    if (i == start_x)
                    {
                        _style.BorderTop = style.BorderTop;
                    }
                    else if (keep_others)
                    {
                        _style.BorderTop = form_cell.CellStyle.BorderTop;
                    }
                    else
                    {
                        _style.BorderTop = BorderStyle.None;
                    }
                    if (i == end_x)
                    {
                        _style.BorderBottom = style.BorderBottom;
                    }
                    else if (keep_others)
                    {
                        _style.BorderBottom = form_cell.CellStyle.BorderBottom;
                    }
                    else
                    {
                        _style.BorderBottom = BorderStyle.None;
                    }
                    if (j == start_y)
                    {
                        _style.BorderLeft = style.BorderLeft;
                    }
                    else if (keep_others)
                    {
                        _style.BorderLeft = form_cell.CellStyle.BorderLeft;
                    }
                    else
                    {
                        _style.BorderLeft = BorderStyle.None;
                    }
                    if (j == end_y)
                    {
                        _style.BorderRight = style.BorderRight;
                    }
                    else if (keep_others)
                    {
                        _style.BorderRight = form_cell.CellStyle.BorderRight;
                    }
                    else
                    {
                        _style.BorderRight = BorderStyle.None;
                    }
                }
                form_cell.CellStyle = _style;
            }
        }
    }
    internal void Set_row_heights(double[] array)
    {
        row_heights = array;
    }
    internal void Set_col_widths(double[] array, double scale = 295)
    {
        for (int id = 0; id < array.Length; ++id)
        {
            array[id] *= scale;
        }
        col_widths = array;
    }
    virtual internal void Change_style(HSSFWorkbook book) { }
    abstract internal void Init(HSSFWorkbook book);
    abstract internal Dictionary<RangeCell, HSSFRichTextString> Get_cell_informations(HSSFWorkbook book, Dictionary<String, String> pairs);
}

