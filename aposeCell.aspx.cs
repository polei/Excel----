using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Excel导入导出
{
    public partial class aposeCell : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void import_Click(object sender, EventArgs e)
        {
            HttpPostedFile file = FileUpload1.PostedFile;
            if (file.ContentLength <= 0)
            {
                info.Text = "请上传文件";
                return;
            }

            string[] vaildExt = new string[] {".xls",".xlsx" };
            string fileExt = System.IO.Path.GetExtension(file.FileName);
            if (!vaildExt.Contains(fileExt))
            {
                info.Text = "文件类型不符合要求";
                return;
            }

            #region 保存excel文件
            string newName = Guid.NewGuid().ToString() + fileExt;

            string saveDir = "upload/Excel/";

            if(!System.IO.Directory.Exists(Server.MapPath(saveDir)))
            {
                System.IO.Directory.CreateDirectory(Server.MapPath(saveDir));
            }

            try
            {
                file.SaveAs(Server.MapPath(saveDir + newName));
                info.Text = "upload success";
            }
            catch
            {
                info.Text = "upload failed";
            }
            #endregion
            improtExcel(Server.MapPath(saveDir + newName)); //从本地文件 读取excel
            Response.Flush();
            improtExcel(file.InputStream);


            //从文件流中读取文件
        }


        //导入excel，从本地文件读取
        private void improtExcel(string ExcelName)
        {
            Aspose.Cells.Workbook wk = new Aspose.Cells.Workbook();
            wk.Open(ExcelName);// 这儿是需要导入的文件

            
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("col1", typeof(string));
            DataColumn dc2 = new DataColumn("col2", typeof(string));
            DataColumn dc3 = new DataColumn("col3", typeof(string));
            DataColumn dc4 = new DataColumn("col4", typeof(string));
            DataColumn dc5 = new DataColumn("col5", typeof(string));
            dt.Columns.AddRange(new DataColumn [] {dc1,dc2,dc3,dc4,dc5});

            int totalRowCount = wk.Worksheets[0].Cells.Rows.Count;
            for (int i = 0; i <totalRowCount; i++)//用于EXCEL数据的等号，可以自行固定如：149，也可以自行去读取它的等号；
            {
                DataRow dr = dt.NewRow();

                dr["col1"] = wk.Worksheets[0].Cells[i, 0].Value;//读取文件里面对应的信息
                dr["col2"] = wk.Worksheets[0].Cells[i, 1].Value;
                dr["col3"] = wk.Worksheets[0].Cells[i, 2].Value;
                dr["col4"] = wk.Worksheets[0].Cells[i, 3].Value;
                dr["col5"] = wk.Worksheets[0].Cells[i, 4].Value;

                dt.Rows.Add(dr);
            }

            GridView1.DataSource = dt;
            GridView1.DataBind();

            if (System.IO.File.Exists(ExcelName))
                System.IO.File.Delete(ExcelName);
        }



        //导入excel，从流中读取
        private void improtExcel(System.IO.Stream excelStream)
        {
            Aspose.Cells.Workbook wk = new Aspose.Cells.Workbook(excelStream);
            wk.Open(excelStream);// 这儿是需要导入的文件


            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("col1", typeof(string));
            DataColumn dc2 = new DataColumn("col2", typeof(string));
            DataColumn dc3 = new DataColumn("col3", typeof(string));
            DataColumn dc4 = new DataColumn("col4", typeof(string));
            DataColumn dc5 = new DataColumn("col5", typeof(string));
            dt.Columns.AddRange(new DataColumn[] { dc1, dc2, dc3, dc4, dc5 });

            int totalRowCount = wk.Worksheets[0].Cells.Rows.Count;
            for (int i = 0; i < totalRowCount; i++)//用于EXCEL数据的等号，可以自行固定如：149，也可以自行去读取它的等号；
            {
                DataRow dr = dt.NewRow();

                dr["col1"] = wk.Worksheets[0].Cells[i, 0].Value;//读取文件里面对应的信息
                dr["col2"] = wk.Worksheets[0].Cells[i, 1].Value;
                dr["col3"] = wk.Worksheets[0].Cells[i, 2].Value;
                dr["col4"] = wk.Worksheets[0].Cells[i, 3].Value;
                dr["col5"] = wk.Worksheets[0].Cells[i, 4].Value;

                dt.Rows.Add(dr);
            }

            GridView2.DataSource = dt;
            GridView2.DataBind();

        }

        protected void export_Click(object sender, EventArgs e)
        {
             
             Workbook workbook = new Workbook(); //工作簿 
             Worksheet sheet = workbook.Worksheets[0]; //工作表 
             workbook.Worksheets[0].Name = "排考表";

             #region 样式 已注释
             //Cells cells = sheet.Cells;//单元格 
             ////样式2 
             //Aspose.Cells.Style style2 = workbook.Styles[workbook.Styles.Add()];//新增样式 
             //style2.HorizontalAlignment = TextAlignmentType.Center;//文字居中 
             //style2.VerticalAlignment = TextAlignmentType.Center;
             //style2.Font.Name = "宋体";//文字字体 
             //style2.Font.Size = 12;//文字大小 
             //style2.Font.IsBold = true;//粗体 
             //style2.BackgroundColor = System.Drawing.Color.Yellow;
             //style2.IsTextWrapped = true;//单元格内容自动换行 
             //style2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
             //style2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
             //style2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
             //style2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;    
             //cells.SetRowHeight(3, 10);//设高
             //cells.SetColumnWidth(1, 50);//设宽
             //cells[0, 2].PutValue("样式使用");
             //cells[0, 2].SetStyle(style2); //0表示行号,2表示列号
             //cells.Merge(1, 2, 3, 4);//合并单元格  1表示行号,2表示列号,3表示合并的行号,4表示合并的列数; 把3或者4其中一个改变成1 ,表示不合并行或者列;如cells.Merge(1, 2, 3, 1);只合并三行,不合并列
             #endregion

             List<string> headName = new List<string> { "col1", "col2", "col3", "col4", "col5" };

             
            
             //动态生成excel  利用apose.cells
             for(int i=0;i<headName.Count;i++)
             {
                 sheet.Cells[0, i].PutValue(headName[i]);
             }

             for (int i = 0; i < GridView2.Rows.Count; i++)
             {
                 for(int j=0;j<GridView2.Rows[i].Cells.Count;j++)
                 {
                     sheet.Cells[i + 1, j].PutValue(GridView2.Rows[i].Cells[j].Text);
                 }
             }


             #region   先保存 本地本件，再下载，已注释
             //string filename = Server.MapPath(Guid.NewGuid().ToString() + ".xls");
             //workbook.Save(filename);

             ////以字符流的形式下载文件 
             //FileStream fs = new FileStream(filename, FileMode.Open);
             //byte[] bytes1 = new byte[(int)fs.Length];
             //fs.Read(bytes1, 0, bytes1.Length);
             //fs.Close();
             #endregion




             Response.ContentType = "application/octet-stream";
             //通知浏览器下载文件而不是打开 
             Response.AddHeader("Content-Disposition", "attachment;  filename=" + HttpUtility.UrlEncode("zxc.xls", System.Text.Encoding.UTF8));
           //  Response.BinaryWrite(bytes1);
             Response.BinaryWrite(workbook.SaveToStream().ToArray());
             Response.Flush();
             Response.End();


        }


        //导出模板
        protected void downloadModel_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook(); //工作簿 
            Worksheet sheet = workbook.Worksheets[0]; //工作表 
            workbook.Worksheets[0].Name = "排考表";








            #region 样式 已注释
            //Cells cells = sheet.Cells;//单元格 
            ////样式2 
            //Aspose.Cells.Style style2 = workbook.Styles[workbook.Styles.Add()];//新增样式 
            //style2.HorizontalAlignment = TextAlignmentType.Center;//文字居中 
            //style2.VerticalAlignment = TextAlignmentType.Center;
            //style2.Font.Name = "宋体";//文字字体 
            //style2.Font.Size = 12;//文字大小 
            //style2.Font.IsBold = true;//粗体 
            //style2.BackgroundColor = System.Drawing.Color.Yellow;
            //style2.IsTextWrapped = true;//单元格内容自动换行 
            //style2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            //style2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            //style2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            //style2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;    
            //cells.SetRowHeight(3, 10);//设高
            //cells.SetColumnWidth(1, 50);//设宽
            //cells[0, 2].PutValue("样式使用");
            //cells[0, 2].SetStyle(style2); //0表示行号,2表示列号
            //cells.Merge(1, 2, 3, 4);//合并单元格  1表示行号,2表示列号,3表示合并的行号,4表示合并的列数; 把3或者4其中一个改变成1 ,表示不合并行或者列;如cells.Merge(1, 2, 3, 1);只合并三行,不合并列
            #endregion


            // Get the first worksheet.
            Worksheet worksheet1 = workbook.Worksheets[0];
            // Add a new worksheet and access it.
          
            // Get the validations collection.
            ValidationCollection validations = worksheet1.Validations;
            // Create a new validation to the validations list.
            Validation validation = validations[validations.Add()];
            // Set the validation type.
            validation.Type = Aspose.Cells.ValidationType.List;
            // Set the operator.
            validation.Operator = OperatorType.None;
            // Set the in cell drop down.
            validation.InCellDropDown = true;
            // Set the formula1.
            validation.Formula1 = string.Join(",",new string[] { "102教室", "202教室" } );
            // Enable it to show error.
            validation.ShowError = true;
            // Set the alert type severity level.
            validation.AlertStyle = ValidationAlertType.Stop;
            // Set the error title.
            validation.ErrorTitle = "请选择正确的班级";
            // Set the error message.
            validation.ErrorMessage = "班级不存在于下拉框中！";
            // Specify the validation area.
            CellArea area;
            area.StartRow = 0;
            area.EndRow = 4;
            area.StartColumn = 0;
            area.EndColumn = 0;
            // Add the validation area.
            validation.AreaList.Add(area);
            // Save the Excel file.
         //   workbook.Save("d:\\test\\validationtypelist.xls");













            Response.ContentType = "application/octet-stream";
            //通知浏览器下载文件而不是打开 
            Response.AddHeader("Content-Disposition", "attachment;  filename=" + HttpUtility.UrlEncode("zxc.xls", System.Text.Encoding.UTF8));
            Response.BinaryWrite(workbook.SaveToStream().ToArray());
            Response.Flush();
            Response.End();
        }
    }
}