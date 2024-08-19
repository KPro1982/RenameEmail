using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RenameEmail
{
    public class Helper
    {
        

        
        public static string GetDLLPath()
        {
            return Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Programs\Python\Python312\python312.dll");

        }
        public static string GetPythonRoot()
        {
            return Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Programs\Python\Python312");
        }

        public static string GetModulePath()
        {
            return Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Programs\Python\Python312\Lib\site-packages");
        }

        public static string GetExecutionPath()
        {
            string? currentPath = Path.GetDirectoryName(System.AppContext.BaseDirectory);
            if (currentPath != null) {
                return  Path.GetFullPath(Path.Combine(currentPath, "."));
            }
            else{
                return "";
            }
            

        }
        public static string GetJsonPath()
        {
            return Helper.GetExecutionPath() + @"\timeentries.json";
        }

          public static string GetClientDataFileName()
        {
            return @"G:\clientdata.csv";
        }
        public static string GetExampleFileName()
        {
            return @"G:\timeexamples.csv";
        }

        
        
        public static string GetEmailFolderPath()
        {
            return @"G:\Data\Email";
        }

        public static string ConvertDate(string msgdate)
        {
            DateTime dt = DateTime.Parse(msgdate);
            string dtstr = dt.ToString("yyyyMMdd");
            

            return dtstr;
;
        }
        public static string ConvertTime(string msgtime)
        {
            DateTime dt = DateTime.Parse(msgtime);
            string dtstr = dt.ToString("H:mm:ss");
            

            return dtstr;
;
        }
        public static void RenameEmail()
        {

            MessageData Msg_Extract = new MessageData();



            foreach (var old_file in 
                Directory.EnumerateFiles(Helper.GetExecutionPath(), "*.msg"))
                {


                    if(!old_file.StartsWith("_-_"))
                    {
                        string msgidstr = "";
                        string msgdatestr = "";
                        
                        Msg_Extract.Get(old_file, out msgdatestr, out msgidstr);
                        string msgtime = Helper.ConvertTime(msgdatestr);
                        msgtime = msgtime.Replace(":","-");
                        string datestr = Helper.ConvertDate(msgdatestr);
                        FileInfo fileInfo = new FileInfo(old_file);
                        string pathstr = fileInfo.Directory.FullName;
                        string new_name = "_-_" + datestr + "_" + msgtime;
                        string new_file = pathstr + "\\" + new_name + ".msg";
                        int i = 2;
                        while(File.Exists(new_file))
                        {
                              new_file = pathstr + "\\" + new_name + "-" + i.ToString() + ".msg";
                              i++;
                        } 
                        File.Move(old_file, new_file);

                    }



                }

        }
        
    }
}
