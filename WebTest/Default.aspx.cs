using System;
using System.IO;
using System.Data;
using DHCC.HR.Common;
using System.Collections.Generic;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void btnExport_Click(object sender, EventArgs e)
    {
        Dictionary<string, string> dic = new Dictionary<string, string>();
        dic.Add("姓名", "姓名");
        dic.Add("序号", "序号");
        DataTable dt = GetDataTable();
        ExcelOperate.ToExcel(dt,dic,"cs.xls");
    }
    protected void btnImport_Click(object sender, EventArgs e)
    {
        string path = Server.MapPath("~/Temp/");
        if (fileUpload.HasFile)
        {
            fileUpload.SaveAs(path + fileUpload.FileName);
        }

        DataTable dt = ExcelOperate.ToDataTable(path + fileUpload.FileName);

        gvPS.DataSource = dt;
        gvPS.DataBind();
    }
    private DataTable GetDataTable()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("序号", typeof(int));
        dt.Columns.Add("姓名", typeof(string));
        dt.Columns.Add("性别", typeof(string));
        dt.Columns.Add("身份证", typeof(string));
        dt.Columns.Add("随机唯一标识码", typeof(string));

        dt.Rows.Add(1, "傅芷若", "女", "511702197407135024", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(2, "顾岚彩", "女", "511702198304257904", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(3, "韦问萍", "女", "511702198107283986", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(4, "唐芷文", "女", "511702199001103486", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(5, "姜娟巧", "女", "511702197301289703", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(6, "郎芳芳", "女", "451025197709242781", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(7, "罗忆梅", "女", "451025198607141183", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(8, "廉清逸", "女", "451025197606178342", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(9, "冯凌雪", "女", "45102519840920354X", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(10, "柏娜兰", "女", "411525197204252845", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(11, "卞涵韵", "女", "120000198806269580", Guid.NewGuid().ToString("N"));
        dt.Rows.Add(12, "岑安卉", "女", "120000198301207800", Guid.NewGuid().ToString("N"));

        return dt;
    }
}