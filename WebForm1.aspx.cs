using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
namespace Excel导入导出
{
    public partial class WebForm1 : System.Web.UI.Page
    {
       
        private void Page_Load(object sender, System.EventArgs e)
        {
            // 在此处放置用户代码以初始化页面
        }

        #region Web 窗体设计器生成的代码
        override protected void OnInit(EventArgs e)
        {
            //
            // CODEGEN: 该调用是 ASP.NET Web 窗体设计器所必需的。
            //
            InitializeComponent();
            base.OnInit(e);
        }

        /// <summary>
        /// 设计器支持所需的方法 - 不要使用代码编辑器修改
        /// 此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.BtnImport.Click += new System.EventHandler(this.BtnImport_Click);
            this.Load += new System.EventHandler(this.Page_Load);

        }
        #endregion


        //// <summary>
        /// 从Excel提取数据--》Dataset
        /// </summary>
        /// <param name="filename">Excel文件路径名</param>
        private void ImportXlsToData(string fileName)
        {
            try
            {
                if (fileName == string.Empty)
                {
                    throw new ArgumentNullException("Excel文件上传失败！");
                }

                string oleDBConnString = String.Empty;
                oleDBConnString = "Provider=Microsoft.Jet.OLEDB.4.0;";
                oleDBConnString += "Data Source=";
                oleDBConnString += fileName;
                oleDBConnString += ";Extended Properties=Excel 8.0;";
                OleDbConnection oleDBConn = null;
                OleDbDataAdapter oleAdMaster = null;
                DataTable m_tableName = new DataTable();
                DataSet ds = new DataSet();

                oleDBConn = new OleDbConnection(oleDBConnString);
                oleDBConn.Open();
                m_tableName = oleDBConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (m_tableName != null && m_tableName.Rows.Count > 0)
                {

                    m_tableName.TableName = m_tableName.Rows[0]["TABLE_NAME"].ToString();

                }
                string sqlMaster;
                sqlMaster = " SELECT *  FROM [" + m_tableName.TableName + "]";
                oleAdMaster = new OleDbDataAdapter(sqlMaster, oleDBConn);
                oleAdMaster.Fill(ds, "m_tableName");
                oleAdMaster.Dispose();
                oleDBConn.Close();
                oleDBConn.Dispose();

                AddDatasetToSQL(ds, 14);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 上传Excel文件
        /// </summary>
        /// <param name="inputfile">上传的控件名</param>
        /// <returns></returns>
        private string UpLoadXls(System.Web.UI.HtmlControls.HtmlInputFile inputfile)
        {
            string orifilename = string.Empty;
            string uploadfilepath = string.Empty;
            string modifyfilename = string.Empty;
            string fileExtend = "";//文件扩展名
            int fileSize = 0;//文件大小
            try
            {
                if (inputfile.Value != string.Empty)
                {
                    //得到文件的大小
                    fileSize = inputfile.PostedFile.ContentLength;
                    if (fileSize == 0)
                    {
                        throw new Exception("导入的Excel文件大小为0，请检查是否正确！");
                    }
                    //得到扩展名
                    fileExtend = inputfile.Value.Substring(inputfile.Value.LastIndexOf(".") + 1);
                    if (fileExtend.ToLower() != "xls")
                    {
                        throw new Exception("你选择的文件格式不正确，只能导入EXCEL文件！");
                    }
                    //路径
                    uploadfilepath = Server.MapPath("~/Service/GraduateChannel/GraduateApply/ImgUpLoads");
                    //新文件名
                    modifyfilename = System.Guid.NewGuid().ToString();
                    modifyfilename += "." + inputfile.Value.Substring(inputfile.Value.LastIndexOf(".") + 1);
                    //判断是否有该目录
                    System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(uploadfilepath);
                    if (!dir.Exists)
                    {
                        dir.Create();
                    }
                    orifilename = uploadfilepath + "\\" + modifyfilename;
                    //如果存在,删除文件
                    if (File.Exists(orifilename))
                    {
                        File.Delete(orifilename);
                    }
                    // 上传文件
                    inputfile.PostedFile.SaveAs(orifilename);
                }
                else
                {
                    throw new Exception("请选择要导入的Excel文件!");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return orifilename;
        }

        /// <summary>
        /// 将Dataset的数据导入数据库
        /// </summary>
        /// <param name="pds">数据集</param>
        /// <param name="Cols">数据集列数</param>
        /// <returns></returns>
        private bool AddDatasetToSQL(DataSet pds, int Cols)
        {
            int ic, ir;
            ic = pds.Tables[0].Columns.Count;
            if (pds.Tables[0].Columns.Count < Cols)
            {
                throw new Exception("导入Excel格式错误！Excel只有" + ic.ToString() + "列");
            }
            ir = pds.Tables[0].Rows.Count;
            string str="";
            if (pds != null && pds.Tables[0].Rows.Count > 0)
            {
                for (int i = 1; i < pds.Tables[0].Rows.Count; i++)
                {
                    str += string.Format(@"
                                {0}   {1}   {2}   {3}   {4}   {5}",
                                pds.Tables[0].Rows[i][0].ToString(),
                                pds.Tables[0].Rows[i][1].ToString(),
                                pds.Tables[0].Rows[i][2].ToString(),
                                pds.Tables[0].Rows[i][3].ToString(),
                                pds.Tables[0].Rows[i][4].ToString(),
                                pds.Tables[0].Rows[i][5].ToString());
                }
            }
            else
            {
                throw new Exception("导入数据为空！");
            }
            return true;
        }

      
        private void BtnImport_Click(object sender, System.EventArgs e)
        {
            string filename = string.Empty;
            try
            {
                filename = UpLoadXls(FileExcel);//上传XLS文件
                ImportXlsToData(filename);//将XLS文件的数据导入数据库                
                if (filename != string.Empty && System.IO.File.Exists(filename))
                {
                    System.IO.File.Delete(filename);//删除上传的XLS文件
                }
                LblMessage.Text = "数据导入成功！";
            }
            catch (Exception ex)
            {
                LblMessage.Text = ex.Message;
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("col1", typeof(string));
            DataColumn dc2 = new DataColumn("col2", typeof(string));
            DataColumn dc3 = new DataColumn("col3", typeof(string));
            DataColumn dc4 = new DataColumn("col4", typeof(string));
            DataColumn dc5 = new DataColumn("col5", typeof(string));
            dt.Columns.AddRange(new DataColumn [5] {dc1,dc2,dc3,dc4,dc5});

            for(int i=0;i<100;i++)
            {
                DataRow dr = dt.NewRow();
                dr[0]="a"+i.ToString();
                dr[1]="b"+i.ToString();
                dr[2]="c"+i.ToString();
                dr[3]="d"+i.ToString();
                dr[4]="e"+i.ToString();
                dt.Rows.Add(dr);
            }
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
        //    CreateExcel(ds,"testExcel.xls");
        }

    }
}