using ExcelSol;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using TestAutomationEmail.Pages;
using System.IO;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace TestAutomationEmail
{
    public class getEmailData : TestFixtureBase
    {
        [Test]
        public void getData()
        {
            EmailPage emailPage = new EmailPage();
            string HTMLbody = emailPage.searchEmail("Test Automation Email");

            ExcelPage excelPage = new ExcelPage();
            
            string path = excelPage.getExcelPath("Files", "EmailData.xlsx");
            DataTable dt = excelPage.getTable(HTMLbody);

            excelPage.insertRow(dt, path);
        }
    }
}
