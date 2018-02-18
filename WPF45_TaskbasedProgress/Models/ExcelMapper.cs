using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using ImportProducts.Models;

namespace CSV.Models
{
    public class ExcelMapper
    {

        public List<Descriptions> MapToDescriptions()
        {
            var dvEmp = new DataView();
            _logger.LogWrite("Getting refs from description file");
            try
            {
                using (var connectionHandler = new OleDbConnection(System.Configuration.ConfigurationManager.AppSettings["ExcelConnectionString"]))
                {
                    connectionHandler.Open();
                    var adp = new OleDbDataAdapter("SELECT * FROM [Sheet1$A:R]", connectionHandler);

                    var dsXls = new DataSet();
                    adp.Fill(dsXls);
                    dvEmp = new DataView(dsXls.Tables[0]);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                _logger.LogWrite("Error occured getting refs from description file: " + e);
            }

            var descriptions = new List<Descriptions>();

            for (var i = 0; i < dvEmp.Table.Rows.Count; i++)
            {
                var descrip = new Descriptions()
                {
                    T2TRef = (from DataRow row in dvEmp.Table.Rows select row["T2TREF"] != DBNull.Value ? (string)row["T2TREF"] : "").ElementAt(i),
                    Descriptio = (from DataRow row in dvEmp.Table.Rows select row["TITLE"] != DBNull.Value ? (string)row["TITLE"] : "").ElementAt(i),
                    Description = (from DataRow row in dvEmp.Table.Rows select row["DESCRIPTION"] != DBNull.Value ? (string)row["DESCRIPTION"] : "").ElementAt(i),
                    Bullet1 = (from DataRow row in dvEmp.Table.Rows select row["BULLET 1"] != DBNull.Value ? (string)row["BULLET 1"] : "").ElementAt(i),
                    Bullet2 = (from DataRow row in dvEmp.Table.Rows select row["BULLET 2"] != DBNull.Value ? (string)row["BULLET 2"] : "").ElementAt(i),
                    Bullet3 = (from DataRow row in dvEmp.Table.Rows select row["BULLET 3"] != DBNull.Value ? (string)row["BULLET 3"] : "").ElementAt(i),
                    Bullet4 = (from DataRow row in dvEmp.Table.Rows select row["BULLET 4"] != DBNull.Value ? (string)row["BULLET 4"] : "").ElementAt(i),
                    Bullet5 = (from DataRow row in dvEmp.Table.Rows select row["BULLET 5"] != DBNull.Value ? (string)row["BULLET 5"] : "").ElementAt(i),
                    Bullet6 = (from DataRow row in dvEmp.Table.Rows select row["BULLET 6"] != DBNull.Value ? (string)row["BULLET 6"] : "").ElementAt(i),
                    Bullet7 = (from DataRow row in dvEmp.Table.Rows select row["BULLET 7"] != DBNull.Value ? (string)row["BULLET 7"] : "").ElementAt(i)

                };

                descriptions.Add(descrip);
            }
            return descriptions;
        }

        private LogWriter _logger = new LogWriter();
    }
}
