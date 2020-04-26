using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace WordFormat
{
    class Program
    {
        static void Main(string[] args)
        {
            string strFolder = null;
            Console.WriteLine("输入文件夹路径开始任务:");
            while (true)
            {
                strFolder = Console.ReadLine();
                if (Directory.Exists(strFolder))
                {
                    break;
                }
                else
                {
                    Console.WriteLine("文件夹不存在，请重新输入:");
                }
            }

            string[] strsFiles = Directory.GetFiles(strFolder, "*.*", SearchOption.AllDirectories).Where(f => !f.Contains("~$") && f.Contains(".doc")).ToArray();
            Application app = GetApplication();
            foreach (string file in strsFiles)
            {
                try
                {
                    FormatFile(file, app);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                
                Console.WriteLine(file);
            }
            Console.WriteLine("全部完成");
            Console.ReadKey();
        }

        private static void FormatFile(string strFile, Application app)
        {
            Document doc = app.Documents.Open(strFile);
            app.Selection.WholeStory();
            app.Selection.Font.Size = 10;
            app.Selection.Font.Name = "微软雅黑";

            app.Selection.GoTo(WdGoToItem.wdGoToLine, WdGoToDirection.wdGoToFirst);
            app.Selection.MoveEndUntil("\v");
            if (app.Selection.Characters.Count == 1)
            {
                app.Selection.MoveEndUntil("\r");
            }
            app.Selection.Font.Size = 20;
            app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            app.Selection.ParagraphFormat.LineUnitBefore = 0.5f;
            app.Selection.ParagraphFormat.LineUnitAfter = 0.5f;

            doc.Save();
            doc.Close();
        }

        private static Application GetApplication()
        {
            Application app = Application.GetActiveInstance();
            if (app == null)
            {
                app = new Application();
            }
            return app;
        }
    }
}
