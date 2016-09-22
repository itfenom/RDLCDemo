using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Reporting.WebForms;

namespace RDLCReportDemo
{
    class Program
    {
        static void Main(string[] args)
        {

            Warning[] warnings;
            String[] streamIds;
            var mimeType = string.Empty;
            var encoding = string.Empty;
            var extension = string.Empty;
            //// Setup the report viewer object and get the array of bytes
            using (var lr = new LocalReport())
            {
                ReportParameter rp1 = new ReportParameter("Parameter1", "AAAAAAAAAAAAAAAAAAAAAAA\r\nAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA\r\nAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaA");
                ReportParameter rp2 = new ReportParameter("Parameter2", "BBB");

                lr.ReportPath = @"C:\Users\chengcui.AVEPOINT\Documents\Visual Studio 2015\Projects\RDLCReportDemo\RDLCReportDemo\DemoReport.rdlc";

                lr.SetParameters(new[] { rp1, rp2 });
                lr.DataSources.Add(new ReportDataSource("DataSet1", GetData()));


                var bytes = lr.Render("WORD", null, out mimeType, out encoding, out extension, out streamIds, out warnings);

                var filename = @"Monthly Sales.doc";
                if (File.Exists((filename)))
                {
                    File.Delete(filename);
                }

                using (FileStream fs = new FileStream(filename, FileMode.Create))
                {
                    fs.Write(bytes, 0, bytes.Length);
                }

            }

        }

        private static DataTable GetData1()
        {
            DataTable dt = new DataTable("DataTable1");
            dt.Columns.Add(new DataColumn("d1", typeof(string)));
            dt.Columns.Add(new DataColumn("d2", typeof(decimal)));
            dt.Columns.Add(new DataColumn("d3", typeof(string)));
            DataRow dr = dt.NewRow();
            dr["c1"] = "张三";
            dr["c2"] = 3300.00m;
            dr["c3"] = "人事";
            dt.Rows.Add(dr);
            return dt;
        }


        private static DataTable GetData()
        {
            DataTable dt = new DataTable("DataTable1");
            dt.Columns.Add(new DataColumn("c1", typeof(string)));
            dt.Columns.Add(new DataColumn("c2", typeof(decimal)));
            dt.Columns.Add(new DataColumn("c3", typeof(string)));

            DataRow dr = dt.NewRow();
            dr["c1"] = "张三";
            dr["c2"] = 3300.00m;
            dr["c3"] = "人事";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["c1"] = "李四";
            dr["c2"] = 3500.00m;
            dr["c3"] = "后勤";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["c1"] = "XJ";
            dr["c2"] = 7500.00m;
            dr["c3"] = "技术";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["c1"] = "CSC";
            dr["c2"] = 8500.00m;
            dr["c3"] = "技术";
            dt.Rows.Add(dr);
            return dt;
        }
    }


}
