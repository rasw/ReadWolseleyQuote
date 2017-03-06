using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace Read_Wolseley_Quote
{
    class OutputToFile
    {

        static string m_fullOutputFilePath = Directory.GetCurrentDirectory() + @"\Output";


        #region Constructors
        public OutputToFile() { }

        #endregion

        public static void WriteToFile(string Data)
        {
           WriteOutput(Data);
        }

        private static async void WriteOutput(string data)
        {
            try
            {
                if (!Directory.Exists(m_fullOutputFilePath))
                    Directory.CreateDirectory(m_fullOutputFilePath);

                if (m_fullOutputFilePath != null)
                {
                    if (m_fullOutputFilePath.Length != 0)
                    {
                        using (StreamWriter sr = new StreamWriter(m_fullOutputFilePath + @"\Output.txt", true))
                        {
                            if (data != null)
                                await sr.WriteLineAsync(data);
                        }
                    }
                }
            }
            catch //(FileNotFoundException fileNotFound)
            {
            }
        }
    }
}
