using Grpc.Core;
using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.SqlServer.Management.Common;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Runtime.InteropServices;
using Microsoft.SqlServer.Management.Smo;
using Server = Microsoft.SqlServer.Management.Smo.Server;

public class Program
{

    public static void Main()
    {

        //We create our download directory
        string dynamicPath = Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.Personal));
        dynamicPath = Path.Combine(dynamicPath, "Downloads//MTM_Project");
        string downloadPathDynamic = Path.Combine(dynamicPath, "MTM_Files");
        string sqlPathDynamic = Path.Combine(dynamicPath, "SQL");

        if (!Directory.Exists(downloadPathDynamic))
            Directory.CreateDirectory(downloadPathDynamic);

        DownloadMTMFiles(downloadPathDynamic).Wait();

        CreateDatabase("MTM_Assessment", sqlPathDynamic);

        string exampleFilePath = Path.Combine(sqlPathDynamic, "Create & Data Script for Table [DailyMTM].sql");
        ProcessExampleFile(exampleFilePath);

        WriteExcelFilesToTable(downloadPathDynamic);

        DateTime from = new DateTime(2021, 01, 04);
        DateTime to = new DateTime(2021, 01, 05);
        WriteReportToExcel(dynamicPath, from, to);
    }

    public static async Task DownloadMTMFiles(string folderPath)
    {
        string url = @"https://clientportal.jse.co.za/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM";

        var links = GetLinks(url);

        var allValidDownloadLinks = new List<string>();

        foreach (var item in links)
        {
            string fileName = item.Split('/').LastOrDefault();
            if (!fileName.Contains("2023"))
            {
                continue;
            }

            //Here we go down recursively, just in case they update the structures per month 
            string possibleFileName = item.Split('/').LastOrDefault();

            if (possibleFileName != null && !possibleFileName.Contains("2023"))
            {
                continue;
            }

            if (Int32.TryParse(possibleFileName, out int val))
            {
                string inner_url = url + ("/" + possibleFileName);

                var inner_links = GetLinks(inner_url);

                foreach (var inner_item in inner_links)
                {
                    string inner_possibleFileName = inner_item.Split('/').LastOrDefault();
                    if (IsValidFile(inner_possibleFileName))
                    {
                        allValidDownloadLinks.Add(inner_item);
                    }
                }
            }
            else if (IsValidFile(possibleFileName))
            {
                allValidDownloadLinks.Add(item);
            }

        }

        //Download all files based off saved and validated links
        string baseURI = @"https://clientportal.jse.co.za/";
        foreach (var link in allValidDownloadLinks)
        {
            string fileName = link.Split('/').LastOrDefault();

            string filePath = Path.Combine(folderPath, fileName);
            Path.GetFullPath(filePath);
            if (File.Exists(filePath))
            {
                continue;
            }

            using (var httpClient = new HttpClient())
            {
                try
                {
                    Console.WriteLine("Downloading:" + fileName);

                    Stream fileStream = await httpClient.GetStreamAsync(baseURI + link);


                    using (FileStream outputFileStream = new FileStream(filePath, FileMode.CreateNew))
                    {
                        await fileStream.CopyToAsync(outputFileStream);
                    }

                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(fileName + " was not able to be downloaded, please evaluate link manually:\n" + link);
                }
            }
        }
        
    }

    private static bool IsValidFile(string fileName)
    {
        //Ensures we have an expected file name at the end
        if (fileName == null)
        {
            Console.Error.WriteLine(fileName + "represented and unknown file name");
            return false;
        }
        //We don't want pdf files
        else if (fileName.EndsWith(".pdf"))
        {
            return false;
        }
        return true;
    }

    private static List<string> GetLinks(string url)
    {

        HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
        HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

        HtmlDocument doc = new HtmlDocument();
        doc.Load(webResponse.GetResponseStream());

        //We start by trying to find the year directory, we check this by looking for the class of margin-bot-10
        var links = new List<string>();
        var divs = doc.DocumentNode.SelectNodes("//div[@class='col__inner__container margin-bot-10']");
        // If 2 is not found, either the site was updated, or under our assumption the directory does not exist
        if (divs.Count == 2)
        {
            if (divs[0] != null)
            {
                return divs[0].Descendants("a")
                               .Select(a => a.GetAttributeValue("href", String.Empty))
                               .ToList();

            }
        }
        return new List<string>();
    }

    private static void CreateDatabase(string programName, string folderPath)
    {

        folderPath = Path.Combine(folderPath, "SQL");
        SqlConnection connection = new SqlConnection("Server=localhost;Integrated security=SSPI;Trusted_Connection=yes");
        string dbName = programName + "_db";
        string dataName = programName + "_data";
        string dbFile = Path.Combine(folderPath, programName + ".mdf");
        string dbLogName = programName + "_Log";
        string logFile = Path.Combine(folderPath, dbLogName + ".ldf");

        Console.WriteLine("Checking if db already exists");
        //if(CheckDatabaseExists(connection, dbName))
        //{
        //    return;
        //}

        String str = String.Format("CREATE DATABASE {0} ON PRIMARY " +
         "(NAME = {1}, " +
         "FILENAME = '{2}', " +
         "SIZE = 2MB, MAXSIZE = 10MB, FILEGROWTH = 10%)" +
         "LOG ON (NAME = {3}, " +
         "FILENAME = '{4}', " +
         "SIZE = 1MB, " +
         "MAXSIZE = 15MB, " +
         "FILEGROWTH = 10%)",
         dbName, dataName, dbFile, dbLogName, logFile);

        SqlCommand command = new SqlCommand(str, connection);
        try
        {
            connection.Open();
            command.ExecuteNonQuery();
            Console.WriteLine("Successfully created DATABASE " + programName);
        }
        catch (System.Exception ex)
        {
            Console.Error.WriteLine("Failed to create DATABASE " + programName);
        }
        finally
        {
            if (connection.State == ConnectionState.Open)
            {
                connection.Close();
            }
        }
    }

    private static void ProcessExampleFile(string filePath)
    {
        SqlConnection myConn = new SqlConnection("Server=localhost;Integrated security=SSPI;database=master");

        string script = File.ReadAllText(filePath);

        SqlCommand myCommand = new SqlCommand(script, myConn);
        try
        {
            myConn.Open();
            myCommand.ExecuteNonQuery();
            Console.WriteLine("Example file process successfully");
        }
        catch (System.Exception ex)
        {
            Console.Error.WriteLine("Unable to process example file");
        }
        finally
        {
            if (myConn.State == ConnectionState.Open)
            {
                myConn.Close();
            }
        }
    }

    private static void WriteExcelFilesToTable(string folderPath)
    {
        var files = Directory.EnumerateFiles(folderPath).ToList();
        foreach (var file in files)
{
            int year = Int32.Parse(file.Substring(0, 4));
            int month = Int32.Parse(file.Substring(4, 2));
            int day = Int32.Parse(file.Substring(6, 2));

            DateTime fileDate = new DateTime(year, month, day);

            Excel.Application xcl_Application = new Excel.Application();
            Excel.Workbook xcl_workBook = xcl_Application.Workbooks.Open(file);

            Excel._Worksheet xcl_Worksheet = (Excel._Worksheet)xcl_workBook.Sheets[1];
            Excel.Range xcl_Range = xcl_Worksheet.UsedRange;

            int rowCount = xcl_Range.Rows.Count;
            int colCount = xcl_Range.Columns.Count;

            DataTable dt = GetDataTable();
            for (int row = 6; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    var dataRow = dt.NewRow();
                    dataRow.SetField("Id", 0);
                    dataRow.SetField("FileDate", fileDate);

                    if (xcl_Range.Cells[row, col] != null)
                    {
                        //As per instruction we skip collumn 2
                        if (col == 2)
                        {
                            continue;
                        }
                        var cell = xcl_Range.Cells[row, col];
                        dataRow.SetField(GetCollumnName(col), cell);
                    }
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            xcl_workBook.Close();

            xcl_Application.Quit();
            xcl_Application.Workbooks.Close();

            //Write data table to sql
            WriteDataTableToSQL(dt, fileDate);
        }
    }

    private static DataTable GetDataTable()
    {
        DataTable dt = new DataTable();

        dt.Columns.Add("Id");
        dt.Columns.Add("FileDate");
        dt.Columns.Add("Contract");
        dt.Columns.Add("ExpiryDate");
        dt.Columns.Add("Classification");
        dt.Columns.Add("Strike");
        dt.Columns.Add("CallPut");
        dt.Columns.Add("MTMYield");
        dt.Columns.Add("MarkPrice");
        dt.Columns.Add("SpotRate");
        dt.Columns.Add("PreviousMTM");
        dt.Columns.Add("PreviousPrice");
        dt.Columns.Add("PremiumOnOption");
        dt.Columns.Add("Volatility");
        dt.Columns.Add("Delta");
        dt.Columns.Add("DeltaValue");
        dt.Columns.Add("ContractsTraded");
        dt.Columns.Add("OpenInterest");


        return dt;
    }

    private static string GetCollumnName(int index)
    {
        switch (index)
        {
            case 1: return "Contract"; 
            case 3: return "ExpiryDate";
            case 4: return "Classification";
            case 5: return "Strike";
            case 6: return "CallPut";
            case 7: return "MTMYield";
            case 8: return "MarkPrice";
            case 9: return "SpotRate";
            case 10: return "PreviousMTM";
            case 11: return "PreviousPrice";
            case 12: return "PremiumOnOption";
            case 13: return "Volatility";
            case 14: return "Delta";
            case 15: return "DeltaValue";
            case 16: return "ContractsTraded";
            case 17: return "OpenInterest";

            default: return "";
        }
    }

    private static void WriteDataTableToSQL(DataTable dt, DateTime fileDate)
    {
        string connString = @"Server=localhost;Integrated security=SSPI;database=master";
        using (SqlConnection connection = new SqlConnection(connString))
        {
            connection.Open();
            using (SqlDataAdapter adapter = new SqlDataAdapter())
            {
                SqlCommand selectCommand = new SqlCommand(@"select * from DailyMTM where FileDate = '" + fileDate.ToString("yyyy-MM-dd") + "'");
                selectCommand.Connection = connection;
                adapter.SelectCommand = selectCommand;
                SqlCommandBuilder cb = new SqlCommandBuilder(adapter);
                adapter.UpdateCommand = cb.GetUpdateCommand();
                adapter.Update(dt);
            }
        }
        
    }

    private static void WriteReportToExcel(string folderPath, DateTime from, DateTime to)
    {
        string connString = @"Server=localhost;Integrated security=SSPI;database=master";
        using (SqlConnection connection = new SqlConnection(connString))
        {
            connection.Open();
            using (SqlDataAdapter adapter = new SqlDataAdapter())
            {
                string command = String.Format(@"EXEC SP_Total_Contracts_Traded_Report "
                    + "@DateFrom = N'{0:YYYY-MM-dd}', @DateTo = N'{1:YYYY-MM-dd}'", from, to);

                SqlCommand queryCommand = new SqlCommand(command);
                queryCommand.Connection = connection;
                adapter.SelectCommand = queryCommand;

                using(SqlDataReader reader = queryCommand.ExecuteReader())
                {
                    Excel._Application xcl_Application = new Excel.Application();
                    Excel.Workbook xcl_workBook = xcl_Application.Workbooks.Add();

                    Excel._Worksheet xcl_Worksheet = (Excel._Worksheet)xcl_workBook.Sheets.Add();
                    Excel.Range xcl_Range = xcl_Worksheet.UsedRange;

                    xcl_Worksheet = (Excel.Worksheet)xcl_Application.ActiveSheet;

                    var dataTypeDateTime = Type.GetType("System.DateTime");
                    var dataTypeString = Type.GetType("System.String");
                    var dataTypeDecimal = Type.GetType("System.Decimal");

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add(new DataColumn(Headers.FileDate, dataTypeDateTime));
                    dataTable.Columns.Add(new DataColumn(Headers.Contract, dataTypeString));
                    dataTable.Columns.Add(new DataColumn(Headers.ContractsTraded, dataTypeDecimal));
                    dataTable.Columns.Add(new DataColumn(Headers.TotalContractsTraded, dataTypeDecimal));

                    do
                    {
                        int fielcCount = reader.FieldCount;
                        for(int i = 0; i< fielcCount; i++)
                        {
                            var row = dataTable.NewRow();
                            row.SetField(Headers.FileDate, reader.GetDateTime(Headers.FileDate));
                            row.SetField(Headers.Contract, reader.GetString(Headers.Contract));
                            row.SetField(Headers.ContractsTraded, reader.GetDecimal(Headers.ContractsTraded));
                            row.SetField(Headers.TotalContractsTraded, reader.GetString(Headers.TotalContractsTraded));
                            
                            dataTable.Rows.Add(row);
                        }

                    } while (reader.Read());


                    foreach (DataColumn col in dataTable.Columns)
                    {
                        xcl_Range.Cells[1, col.Ordinal + 1] = col.ColumnName;
                    }

                    int rowIndex = 2;

                    foreach(DataRow dataRow in dataTable.Rows)
                    { 
                        xcl_Range.Cells[rowIndex, 1] = dataRow.Field<DateTime>(Headers.FileDate);
                        xcl_Range.Cells[rowIndex, 2] = dataRow.Field<DateTime>(Headers.Contract);
                        xcl_Range.Cells[rowIndex, 3] = dataRow.Field<DateTime>(Headers.ContractsTraded);
                        xcl_Range.Cells[rowIndex, 4] = dataRow.Field<DateTime>(Headers.TotalContractsTraded);
                    }

                    xcl_workBook.SaveAs(Path.Combine(folderPath,"SP_Total_Contracts_Traded_Report Result.xls"));

                    //cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    xcl_workBook.Close();

                    xcl_Application.Quit();
                    xcl_Application.Workbooks.Close();
                }

            }
        }
    }

    private static class Headers
    {
        public static readonly string FileDate = "FileDate";
        public static readonly string Contract = "Contract";
        public static readonly string ContractsTraded = "ContractsTraded";
        public static readonly string TotalContractsTraded = "Percentage Of Total ContractsTraded";
    }
}
