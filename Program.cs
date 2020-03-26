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
        static string outputFilePath    = @"/Users/adjo/Import/processed/DWS-DONOTUSE-Processed.json";
        static string questionsPath = @"/Users/adjo/Import/questions.json";

        static List<Question> questions = new List<Question>();

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
                questions = JsonConvert.DeserializeObject<List<Question>>(getListOfQuestions());
                #endregion

                #region Open the entire Excel File
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                FileStream stream = File.Open("/Users/adjo/Import/DWS.xlsx", FileMode.Open, FileAccess.Read);

                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                #endregion

                //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet();

                foreach(DataTable table in result.Tables)
                {
                    if ( table.TableName.ToLower() != "introduction")
                    {
                        processInfra(table);
                    }

                    //switch(table.TableName.ToLower())
                    //{
                    //    case "infrastructure":
                    //        processInfra(table);
                    //        break;
                    //    case "optimization patterns":
                    //        processInfra(table);
                    //        break;
                    //    case "transaction load":
                    //        processInfra(table);
                    //        break;
                    //    case "business value":
                    //        processInfra(table);
                    //        break;
                    //    case "business value":
                    //        processInfra(table);
                    //        break;
                    //    case "business value":
                    //        processInfra(table);
                    //        break;
                    //    case "securituy":
                    //        processInfra(table);
                    //        break;

                    //}
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
                Console.WriteLine(ex);
                string test = string.Empty; // just a line to stop on
            }

        }

        // read the Matter output as a JSON object
        static string getListOfQuestions()
        {
            string returnString = string.Empty;

            try
            {
                string path = questionsPath;

                if (File.Exists(path))  // don't try and read it if it's not there
                {
                    // Open the file to read from.
                    returnString = File.ReadAllText(path);
                }
                else
                {
                    throw new Exception("Questions file not loaded into location: " + path);
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

                // find the question in the Excel, find out where the answer should be and pluck the answer out

                bool lookingForAnswer   = false;
                string currentQuestion  = string.Empty;

                foreach(DataRow dr in data.Rows)
                {
                    // whenever there's a value in dr[4] (Column 4) = this is a new question -- grab the guid
                    // whenever there's a value in dr[3] (Column 3) = the user says this is the answer -- grab the value to search the questions
                    //              what's in the question can give us the value

                    // new question. let's see how they answered
                    if (dr[4].ToString() != string.Empty )
                    {
                        currentQuestion = dr[4].ToString();
                        lookingForAnswer = true;
                    }

                    // we are on the same question but are looking in cell 3 for an X or any text really
                    if ( lookingForAnswer )
                    {
                        if( dr[3].ToString() != string.Empty )
                        {
                            // we've got action -- this is the answered question
                            // take the value in cell 2 and look up the value in the questions
                            string questionTextValue = dr[2].ToString();

                            Question questionWithAnswers = questions.Find(x => x.uuid == currentQuestion);  // get the question and all it's answers

                            if (questionWithAnswers != null)
                            {
                                // get the correct answer
                                var questionActualValue = questionWithAnswers.answers.Find(x => x.text.ToString() == questionTextValue).answer;

                                // example "75007aca-e9e6-4959-b79d-a06809b9d960":{ "answer":"a"}
                                string finalOutputString = "\"" + currentQuestion + "\":";
                                finalOutputString += "{" + "\"" + "answer\"" + ":";
                                finalOutputString += "\"" + questionActualValue + "\"}," + Environment.NewLine;

                                File.AppendAllText(outputFilePath, finalOutputString, Encoding.UTF8);   // write the answer to a file
                                lookingForAnswer = false;   // allow the next question to be found
                            }
                        }

                    }

                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
