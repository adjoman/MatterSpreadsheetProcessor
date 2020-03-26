using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using ExcelDataReader;
using Newtonsoft.Json;
using ImportMatterExcel;

namespace ImportMatterExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            try
            {
                // steps
                // 1. read the questions into a collection so we can iterate and find the answers
                // 2. read the Excel file. Process the Worksheets individually
                // 3. write the answers out in this format ("75007aca-e9e6-4959-b79d-a06809b9d960":{"answer":"a"})
                // the answer is the value of the question -- not the item

                #region Read Questions from JSON into Collection
                var notificationList = JsonConvert.DeserializeObject<List<Question>>(getListOfQuestions());
                #endregion

                #region Open the entire Excel File
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                FileStream stream = File.Open("/Users/adjo/Import/MAS_Boarding_WEB_Application_Assessment_Questionnaire.xlsx", FileMode.Open, FileAccess.Read);

                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                #endregion

                //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet();

                foreach(DataTable table in result.Tables)
                {
                    switch(table.TableName.ToLower())
                    {
                        case "infrastructure":
                            processInfra(table);
                            break;
                    }
                }

                //// how does the Excel file look -- change this stuff
                //DataTable dr_Intr               = result.Tables[0];
                //DataTable dr_Optimization       = result.Tables[1];
                ////DataTable dr_TransactionLoad    = result.Tables[2];

                //foreach(DataRow dr in dr_Optimization.Rows)
                //{
                //    Console.WriteLine(dr[1].ToString());
                //}

                stream.Close(); // release the resources of the stream
            }
            catch (Exception ex)
            {
                string test = string.Empty;
            }

        }

        // read the Matter output as a JSON object
        static string getListOfQuestions()
        {
            string returnString = string.Empty;

            try
            {
                string path = @"/Users/adjo/Import/questions.json";

                if (File.Exists(path))
                {
                    // Open the file to read from.
                    returnString = File.ReadAllText(path);
                }
            }
            catch (Exception ex)
            {
                throw;
            }

            return returnString;
        }

        // process the infrstructure worksheet
        static void processInfra(DataTable data)
        {
            try
            {
                // infra looks like this    
                // column 1 column 2    column 3    column 4        column 5
                // 1        Question    Possible Answers    answer      X(if answered)

                foreach(DataRow dr in data.Rows)
                {
                    string test = string.Empty;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
