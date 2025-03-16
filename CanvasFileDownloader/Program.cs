// See https://aka.ms/new-console-template for more information
using ExcelDataReader;
using System;
using System.Data;
using System.Data.Common;
using System.IO.Compression;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;

internal class Program
{
    private static async Task Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        string targetFolder = "D:\\Columbia\\Game225\\Backlog";
        string extractFolder = "D:\\Columbia\\Game225\\BacklogExtract";
        //await DownloadSubmissions(targetFolder);

        ExtractZips(targetFolder, extractFolder);

    }

    private static void ExtractZips(string targetFolder, string extractFolder)
    {
        DirectoryInfo extractFolderInfo = new(extractFolder);
        if (!extractFolderInfo.Exists)
            extractFolderInfo.Create();
        else if (extractFolderInfo.GetFiles().Length > 0)
            extractFolderInfo.EnumerateFiles().ToList().ForEach(act => act.Delete());


        foreach (var zf in Directory.EnumerateFiles(targetFolder))
        {
            string revZf = zf;
            if (zf.EndsWith("_.zip"))
                revZf = zf.Replace("_.zip", ".zip");

            string extractFolder2 = new FileInfo(zf).Name;
            if (extractFolder2.EndsWith("_"))
                extractFolder2 = extractFolder2.Substring(0, extractFolder2.Length - 1);
            extractFolder2 = extractFolder2.Substring(0, extractFolder2.Length - 5);
            string extractFolder3 = Path.Combine(extractFolder, extractFolder2);
            Console.WriteLine(revZf);

            try
            {
                ZipArchive archive = ZipFile.OpenRead(zf);
                ZipFileExtensions.ExtractToDirectory(archive, extractFolder3);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Skipped file:{zf} due to {e.Message}");
            }
        }
    }

    private static async Task DownloadSubmissions(string targetDir)
    {
        string dataFileName = "D:\\Columbia\\Game225\\SubmittedAssignmentsNoGrade.xlsx";

        byte[] content;
        var httpClient = new HttpClient();

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var enc1252 = Encoding.GetEncoding(1252);

        DataTable table = new DataTable("Assns");
        using (var stream = File.Open(dataFileName, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream, new
                   ExcelReaderConfiguration()
            { FallbackEncoding = Encoding.GetEncoding(1252) }))
            {
                var result = reader.AsDataSet();

                // Examples of data access
                table = result.Tables[0];
                DataRow row = table.Rows[0];
                string cell = row[0].ToString()!;
            }
        }
        int assignment_id;
        string url = String.Empty;
        string StudentName = String.Empty;
        string NewFileName = String.Empty;


        EnumerableRowCollection<DataRow> query = from order in table.AsEnumerable()
                                                 orderby order.Field<String>("Column5"),
                                                 order.Field<Object>("Column1")
                                                 select order;

        DataView view = query.AsDataView();



        foreach (DataRowView dr in view)
        {
            if (dr["Column1"].ToString()! == "Column1.assignment_id")
                continue;
            assignment_id = Convert.ToInt32(dr["Column1"].ToString()!);

            url = dr["Column4"].ToString()!;
            StudentName = dr["Column5"].ToString()!;

            NewFileName = Path.Combine(targetDir, dr["Column6"].ToString()!);

            Console.WriteLine($"{assignment_id}, {url}, {StudentName}, {NewFileName}");

            using (var response = await httpClient.GetAsync(url))
            {
                content = await response.Content.ReadAsByteArrayAsync();
                File.WriteAllBytes(NewFileName, content);
            }
        }
    }
}