using System;
using System.IO;
using System.Text.RegularExpressions;
using ZOSAPI;
using ZOSAPI.Analysis;
using ZOSAPI.Analysis.Data;
using ZOSAPI.Analysis.Settings;

namespace CSharpUserOperandApplication
{
    class Program
    {
        private static IA_ footprint_diagram;
        private static IAS_ footprint_settings;
        private static IAR_ footprint_diagram_results;

        static void Main(string[] args)
        {
            // Find the installed version of OpticStudio
            bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
            // Note -- uncomment the following line to use a custom initialization path
            //bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(@"C:\Program Files\OpticStudio\");
            if (isInitialized)
            {
                LogInfo("Found OpticStudio at: " + ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory());
            }
            else
            {
                HandleError("Failed to locate OpticStudio!");
                return;
            }

            BeginUserOperand();
        }

        static void BeginUserOperand()
        {
            // Create the initial connection class
            ZOSAPI_Connection TheConnection = new ZOSAPI_Connection();

            // Attempt to connect to the existing OpticStudio instance
            IZOSAPI_Application TheApplication = null;
            try
            {
                TheApplication = TheConnection.ConnectToApplication(); // this will throw an exception if not launched from OpticStudio
            }
            catch (Exception ex)
            {
                HandleError(ex.Message);
                return;
            }
            if (TheApplication == null)
            {
                HandleError("An unknown connection error occurred!");
                return;
            }

            // Check the connection status
            if (!TheApplication.IsValidLicenseForAPI)
            {
                HandleError("Failed to connect to OpticStudio: " + TheApplication.LicenseStatus);
                return;
            }
            if (TheApplication.Mode != ZOSAPI_Mode.Operand)
            {
                HandleError("User plugin was started in the wrong mode: expected Operand, found " + TheApplication.Mode.ToString());
                return;
            }

            // Read the operand arguments
            double Hx = TheApplication.OperandArgument1;
            double Hy = TheApplication.OperandArgument2;
            double Px = TheApplication.OperandArgument3;
            double Py = TheApplication.OperandArgument4;

            // Initialize the output array
            int maxResultLength = TheApplication.OperandResults.Length;
            double[] operandResults = new double[maxResultLength];

            IOpticalSystem TheSystem = TheApplication.PrimarySystem;
            // Add your custom code here...
            // Create a footprint diagram analysis window
            footprint_diagram = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.FootprintSettings);

            // Modify settings
            if (Hx > 0)
            {
                // Retrieve settings of the footprint diagram
                footprint_settings = footprint_diagram.GetSettings();

                // Modifiy settings
                string filename = System.IO.Path.GetTempFileName();
                footprint_settings.SaveTo(filename);
                footprint_settings.ModifySettings(filename, "FOO_SURFACE", Hx.ToString());
                footprint_settings.LoadFrom(filename);
                System.IO.File.Delete(filename);
            }

            // Apply settings and update analysis
            footprint_diagram.ApplyAndWaitForCompletion();

            // Retrieve settings from the analysis window and write to a text file
            string temporary_filename = @"Footprint_diagram_results.txt";
            string full_path = Path.Combine(TheApplication.SamplesDir, temporary_filename);
            footprint_diagram_results = footprint_diagram.GetResults();
            footprint_diagram_results.GetTextFile(full_path);

            // Parse the text file
            string new_line = "";
            StreamReader sr = new StreamReader(full_path);

            // Ignore header
            for(int ii=0; ii < 8; ii++)
            {
                new_line = sr.ReadLine();
            }

            // Scientific notation regex
            string sci_regex = @"-?[\d.]+(?:E-?\d+)?";

            // Get X min
            new_line = sr.ReadLine();
            string str_value = Regex.Match(new_line, sci_regex).Value;
            double x_min = Convert.ToDouble(str_value);

            // Get X max
            new_line = sr.ReadLine();
            str_value = Regex.Match(new_line, sci_regex).Value;
            double x_max = Convert.ToDouble(str_value);

            // Get Y min
            new_line = sr.ReadLine();
            str_value = Regex.Match(new_line, sci_regex).Value;
            double y_min = Convert.ToDouble(str_value);

            // Get Y max
            new_line = sr.ReadLine();
            str_value = Regex.Match(new_line, sci_regex).Value;
            double y_max = Convert.ToDouble(str_value);

            sr.Close();
            File.Delete(full_path);

            operandResults[0] = x_min;
            operandResults[1] = x_max;
            operandResults[2] = y_min;
            operandResults[3] = y_max;

            // Clean up
            FinishUserOperand(TheApplication, operandResults);
        }

        static void FinishUserOperand(IZOSAPI_Application TheApplication, double[] resultData)
        {
            // Note - OpticStudio will wait for the operand to complete until this application exits 
            if (TheApplication != null)
            {
                TheApplication.OperandResults.WriteData(resultData.Length, resultData);
            }
        }

        static void LogInfo(string message)
        {
            // TODO - add custom logging
            Console.WriteLine(message);
        }

        static void HandleError(string errorMessage)
        {
            // TODO - add custom error handling
            throw new Exception(errorMessage);
        }

    }
}
