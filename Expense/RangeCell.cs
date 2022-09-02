using NPOI.SS.UserModel;
using NPOI.SS.Util;

class RangeCell
{
    public int start_x { get; set; }
    public int start_y { get; set; }
    public int end_x { get; set; }
    public int end_y { get; set; }
    public ICellStyle style { get; set; }

    public RangeCell(int start_x, int start_y, int end_x, int end_y, ICellStyle style)
    {
        this.start_x = start_x;
        this.start_y = start_y;
        this.end_x = end_x;
        this.end_y = end_y;
        this.style = style;
    }

    public bool Is_merge_cell()
    {
        return start_x != end_x || start_y != end_y;
    }

    public CellRangeAddress To_cell_range_address()
    {
        return new CellRangeAddress(start_x, end_x, start_y, end_y);
    }
}
