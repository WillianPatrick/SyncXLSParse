using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using Parse;

namespace SyncXLSParse
{
    public static class Program
    {
        static async Task Main(string[] args)
        {
            string XlsSyncFolderPath = "";
            string _username = "";
            string _password = "";
            int _rowsbuffer = 0;

            ParseClient.Configuration config;

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
                    case "-applicationid":
                        config.ApplicationId = (args[i + 1] != null) ? args[i + 1] : string.Empty;
                        break;
                    case "-server":
                        config.Server = (args[i + 1] != null) ? args[i + 1] : string.Empty;
                        break;
                    case "-key":
                        config.WindowsKey = (args[i + 1] != null) ? args[i + 1] : string.Empty;
                        break;
                    case "-username":
                        _username = (args[i + 1] != null) ? args[i + 1] : string.Empty;
                        break;
                    case "-password":
                        _password = (args[i + 1] != null) ? args[i + 1] : string.Empty;
                        break;
                    case "-xlssyncfolderpath":
                        XlsSyncFolderPath = (args[i + 1] != null) ? args[i + 1] : string.Empty;
                        break;
                    case "-rowsbuffer":
                        _rowsbuffer = args.Length > i ? int.Parse(args[i + 1]) : 0;
                        break;
                    case "-cleantable":
                        _rowsbuffer = args.Length > i ? int.Parse(args[i + 1]) : 0;
                        break;                       
                    default:
                        break;
                } 
            }

            if (!Directory.Exists(XlsSyncFolderPath)){
                throw new Exception(String.Format("XML directory path '{0}' not exists!", XlsSyncFolderPath));
            }

            if (!Directory.Exists(Path.Combine(XlsSyncFolderPath, "Pending")))
                Directory.CreateDirectory(Path.Combine(XlsSyncFolderPath, "Pending"));

            if (!Directory.Exists(Path.Combine(XlsSyncFolderPath, "Processing")))
                Directory.CreateDirectory(Path.Combine(XlsSyncFolderPath, "Processing"));

            if (!Directory.Exists(Path.Combine(XlsSyncFolderPath, "Success")))
                Directory.CreateDirectory(Path.Combine(XlsSyncFolderPath, "Success"));

            if (!Directory.Exists(Path.Combine(XlsSyncFolderPath, "Error")))
                Directory.CreateDirectory(Path.Combine(XlsSyncFolderPath, "Error"));

            ParseClient.Initialize(config);

            await ParseUser.LogInAsync(_username, _password);

            Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - Parse Platform - Sucess initialize.");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var pendingDirectory = new DirectoryInfo(Path.Combine(XlsSyncFolderPath, "Pending"));
            foreach (var xlsxFilePath in pendingDirectory.GetFilesByExtensions(".xls", ".xlsx").OrderBy(x => x.Name))
            {
                string processingFilePath = Path.Combine(XlsSyncFolderPath, "Processing", string.Format("{0}", xlsxFilePath.Name));

                if(File.Exists(processingFilePath)) File.Delete(processingFilePath);
                
                xlsxFilePath.MoveTo(processingFilePath);

                try
                {
                    Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - XLS - Loading data from file: " + xlsxFilePath.Name);
                    using (var stream = File.Open(processingFilePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });

                            Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - XLS - Sheets count: " + result.Tables.Count);

                            var myClassObjects = new List<ParseObject>();
                            foreach (DataTable table in result.Tables)
                            {
                                int totalRows = (table.Rows.Count - 2);
                                int buffer = _rowsbuffer > 0 ? _rowsbuffer : (totalRows > 10000) ? totalRows / 5000 : (totalRows > 1000) ? (totalRows / 100) : 100;
                                int buffercount = 0;
                                int totalRowsSync = 0;

                                Console.WriteLine(string.Format(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - XLS - Current sheet {0} - {1} buffer size: {2}", table.TableName, _rowsbuffer > 0 ? "Dynamic" : "User value", buffer));

                                Console.WriteLine(string.Format(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - XLS - Current sheet {0} - Total rows: {1}", table.TableName, totalRows));
                                Console.WriteLine(string.Format(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - XLS - Current sheet {0} - Importing data...", table.TableName, totalRows));
                                Console.WriteLine("");

                                for (int i = 1; i < table.Rows.Count; i++)
                                {
                                    if (buffercount == buffer)
                                    {
                                        buffercount = 0;
                                        await ParseObject.SaveAllAsync(myClassObjects);
                                        totalRowsSync += buffer;
                                        decimal totalPercent = Math.Round((decimal)((totalRowsSync / totalRows) * 100), 2);
                                        Console.WriteLine(string.Format(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - Integrated rows: {1} - {2}%", totalRowsSync, totalPercent));
                                        myClassObjects = new List<ParseObject>();
                                    }

                                    ParseObject testObject = new ParseObject("XLS_" + table.TableName);
                                    for (int ix = 0; ix <= table.Columns.Count - 1; ix++)
                                        testObject.Add(SyncXLSParse.Program.RemoveSpecialCharacters(table.Columns[ix].ColumnName), SyncXLSParse.Program.ConvertFromDBVal<object>(table.Rows[i].ItemArray[ix]));

                                    myClassObjects.Add(testObject);
                                    buffercount++;
                                }
                                Console.WriteLine(string.Format(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - XLS - Current sheet {0} -  Done!", table.TableName));
                                Console.WriteLine("");
                                Console.WriteLine("");
                            }
                        }
                    }
                    string successFilePath = Path.Combine(XlsSyncFolderPath, "Success", string.Format("{0}", xlsxFilePath.Name));
                    if(File.Exists(successFilePath)) File.Delete(successFilePath);
                    xlsxFilePath.MoveTo(successFilePath);
                    Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - XLS - file: {0} - Imported!", xlsxFilePath.Name);

                    Console.WriteLine("");
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - XLS - file: {0} - Error: ", xlsxFilePath.Name);
                    if(ex.Message != null){
                        Console.WriteLine("");
                        Console.WriteLine("{0}", ex.Message);
                    }

                    if(ex.InnerException != null && ex.Message == null){
                        Console.WriteLine("");
                        Console.WriteLine("{0}", ex.InnerException.Message);
                    }

                    Console.WriteLine("");
                    string errorFilePath = Path.Combine(XlsSyncFolderPath, "Error", string.Format("{0}", xlsxFilePath.Name));
                    if(File.Exists(errorFilePath)) File.Delete(errorFilePath);
                    xlsxFilePath.MoveTo(errorFilePath);
                    continue;
                }
            }
        }

        public static string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < str.Length; i++)
            {
                if ((str[i] >= '0' && str[i] <= '9')
                    || (str[i] >= 'A' && str[i] <= 'z'
                        || (str[i] == '.' || str[i] == '_')))
                {
                    sb.Append(str[i]);
                }
            }

            return sb.ToString();
        }

        static T ConvertFromDBVal<T>(object obj)
        {
            if (obj == null || obj == DBNull.Value)
            {
                return default(T); // returns the default value for the type
            }
            else
            {
                return (T)obj;
            }
        }


        public static IEnumerable<FileInfo> GetFilesByExtensions(this DirectoryInfo dir, params string[] extensions)
        {
            if (extensions == null)
                throw new ArgumentNullException("extensions");
            IEnumerable<FileInfo> files = dir.EnumerateFiles();
            return files.Where(f => extensions.Contains(f.Extension));
        }
    }


}

