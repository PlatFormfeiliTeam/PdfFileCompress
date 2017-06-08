using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace PdfFileCompress
{
    class Program
    {
        string shrinkdir = ConfigurationManager.AppSettings["shrinkdir"];
        string filedir = ConfigurationManager.AppSettings["filedir"];

        static void Main(string[] args)
        {
            if (ConfigurationManager.AppSettings["AutoRun"].ToString().Trim() == "Y")
            {
                Program p = new Program();

                int count = Convert.ToInt32(ConfigurationManager.AppSettings["count"].ToString());
                int sleepsecond = Convert.ToInt32(ConfigurationManager.AppSettings["sleepsecond"].ToString());
                p.pdfcompress(count, sleepsecond);
            }
        }

        private void pdfcompress(int count, int sleepsecond)
        {
            string sql = string.Empty;
            try
            {
                sql = @"select t.attachmentid,t.ID,a.filename,a.filesuffix 
                        from pdfshrinklog t   
                            left join List_Attachment a on t.attachmentid=a.id  
                        where t.iscompress=0 order by ID desc";//id=264160
                DataTable dt = DBMgr.GetDataTable(sql);


               for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i > count) { break; }
                    if (File.Exists(filedir + dt.Rows[i]["FILENAME"]))//先判断原始文件存在
                    {
                        //再对扩展名判断
                        if ((dt.Rows[i]["FILESUFFIX"] + "").ToUpper() == "PDF" || (dt.Rows[i]["FILESUFFIX"] + "").ToUpper() == ".PDF")
                        {
                            if ((new FileInfo(filedir + dt.Rows[i]["FILENAME"])).Length > 0)
                            {
                                System.Diagnostics.Process.Start(shrinkdir, filedir + dt.Rows[i]["FILENAME"]);
                                sql = "update pdfshrinklog set iscompress='1',shrinktime=sysdate WHERE ID='" + dt.Rows[i]["ID"] + "'";
                            }
                            else
                            {
                                sql = "update pdfshrinklog set iscompress='999',shrinktime=sysdate WHERE ID='" + dt.Rows[i]["ID"] + "'";//--大小为0KB，标记为999
                            }
                            
                            DBMgr.ExecuteNonQuery(sql);
                        }
                    }
                    if (sleepsecond > 0) { Thread.Sleep(sleepsecond); }                    
                }

                 /*if (dt.Rows.Count > 0)//如果一次性送多个文件进入压缩程序，会有错误提示,所以每次只送一条记录
                {
                    DataRow dr = dt.Rows[0];
                    if (File.Exists(ConfigurationManager.AppSettings["filedir"] + dr["FILENAME"]))//先判断原始文件存在
                    {
                        //再对扩展名判断
                        if ((dr["FILESUFFIX"] + "").ToUpper() == "PDF" || (dr["FILESUFFIX"] + "").ToUpper() == ".PDF")
                        {
                            System.Diagnostics.Process.Start(@"D:\Apago\PDFShrink\PDFShrink.exe", @"D:\ftpserver\" + dr["FILENAME"]);
                            sql = "update pdfshrinklog set iscompress='1',shrinktime=sysdate WHERE ID='" + dr["ID"] + "'";
                            DBMgr.ExecuteNonQuery(sql);
                        }
                    }
                }*/
            }
            catch (Exception ex)
            {

            }
        }

    }
}
