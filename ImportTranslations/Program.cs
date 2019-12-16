using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using MsCrmTools.Translator;
using Microsoft.Xrm.Tooling.Connector;
using System.Configuration;
using System.IO;

namespace ImportTranslations
{
    class Program
    {
        private static string _filePath = null;

        static void Main(string[] args)
        {
            string directory = string.Empty;
            int? lcidToProcess = null;
            string connStr = null;

            Directory.CreateDirectory("Logs");
            Log($"Running the translation import in batch mode");

            //get arguments
            if (args.Length < 1)
            {
                directory = @"C:\Users\arvind-v\Documents\UNHCR\Data\Translation\XrmToolbox\toprocess";
                lcidToProcess = 1036;

                Log($"Directory not specified! - running test mode for path: {directory}");
            }
            else
            {
                directory = args[0];

                if (args.Length > 1)
                {
                    connStr = args[1];

                    if (ConfigurationManager.ConnectionStrings.Cast<ConnectionStringSettings>()
                            .FirstOrDefault(s => connStr.Equals(s.Name, StringComparison.OrdinalIgnoreCase)) == null)
                    {
                        Log($"Connection string '{connStr}' is not available in config. Aborting ....");
                        return;
                    }
                }

                if (args.Length > 2)
                {
                    if (int.TryParse(args[2], out int parsedInt))
                    {
                        lcidToProcess = parsedInt;
                    }
                }
            }

            Log($"Processing for directory {directory}");
            Log($"Processing for Language {lcidToProcess}");

            if (!Directory.Exists(directory))
            {
                Log($"Incorrect path! {directory}");
                return;
            }

            //Compose CRM Service
            CrmServiceClient c = new CrmServiceClient(ConfigurationManager.ConnectionStrings[connStr].ConnectionString);

            Log("Testing CRM service");
            //test service
            var resp = c.Execute(new WhoAmIRequest()) as WhoAmIResponse;
            Log($"Service is working fine. Current user is {resp.UserId}");

            var filesToImport = Directory.GetFiles(directory);

            Engine e = new Engine();

            foreach (var file in filesToImport)
            {
                try
                {
                    Log($"*************Importing File: {file}");

                    e.Import(file, c, new BackgroundWorker(), lcidToProcess);
                }
                catch (Exception ex)
                {
                    Log($"Error processing file {file}. Error Message: {ex.Message}, More Details: {ex.StackTrace}");
                }

                if (string.IsNullOrWhiteSpace(file) || !file.EndsWith(".xlsx"))
                    continue;
                try
                {
                    //move the file to processed folder
                    var destDir = Path.Combine(Path.GetDirectoryName(file), "Processed");
                    Directory.CreateDirectory(destDir);
                    File.Move(file, Path.Combine(destDir, Path.GetFileName(file)));
                }
                catch (Exception ex)
                {
                    Log($"Error in moving file {file}. Error Message: {ex.Message}, More Details: {ex.StackTrace}");
                }
            }

            Log($"Translation Import Complete");
        }

        private static void Log(string msg)
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                _filePath = "Logs\\ImportTranslations_" + DateTime.Now.Date.ToString("MMddyyyy") + ".log";
            }

            Console.WriteLine(msg);
            File.AppendAllText(_filePath, $"{Environment.NewLine}{DateTime.Now.ToString()} {msg}");
        }
    }
}
