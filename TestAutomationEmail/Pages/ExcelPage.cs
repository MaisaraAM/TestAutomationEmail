using ExcelSol;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTable = System.Data.DataTable;
using HtmlAgilityPack;
using System.Data;

namespace TestAutomationEmail.Pages
{
    public class ExcelPage : TestFixtureBase
    {
        ExcelApi excelApi;

        public string getExcelPath(string FolderName, string FileName)
        {
            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            var applicationPath_new = application_path.Replace("\\bin\\Debug", "");
            var applicationPath_new_name = applicationPath_new + "\\" + FolderName;
            applicationPath_new_name = @applicationPath_new_name.Replace(@"\\" + FolderName, @"\" + FolderName);
            string final_path = $@"{applicationPath_new_name}\{FileName}";
            return final_path;
        }

        public DataTable getTable(string body)
        {
            DataTable dataTable = new DataTable();

            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(body);
            var table = doc.DocumentNode.SelectSingleNode("//table");
            var tableRows = table.SelectNodes("tr");
            var columns = tableRows[0].SelectNodes("td");

            dataTable.Columns.Add(columns[0].InnerText);
            dataTable.Columns.Add(columns[1].InnerText);

            for (int i = 0; i < tableRows.Count; i++)
            {
                var column = tableRows[i].SelectNodes("td");
                dataTable.Rows.Add(column[0].InnerText.ToString(), column[1].InnerText.ToString());
            }
            
            return dataTable;
        }

        public void insertRow(DataTable dt, string excelFilePath)
        {
            excelApi = new ExcelApi(excelFilePath);
            excelApi.OpenExcel();
            List<string> sheetList = excelApi.getSheetName();
            excelApi.deleteRowsInExcel(0);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dataRow = dt.Rows[i];

                excelApi.UpdateCellData(sheetList[0], 1, i + 1, dataRow["Curr Date"].ToString());
                excelApi.UpdateCellData(sheetList[0], 2, i + 1, dataRow["1st Rent Date"].ToString());
            }

            excelApi.CloseExcel();
        }
    }

}

