using Aspose.Cells;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopySheets
{
    public class CopySheets : CodeActivity
    {
        [Category("输入")]
        [DisplayName("源文件路径")]
        [RequiredArgument]
        public InArgument<string> SoursePath { get; set; }

        [Category("输入")]
        [DisplayName("工作表名称")]
        [RequiredArgument]
        [Description("key表示目标工作表名称，value表示源工作表名称")]
        public InArgument<Dictionary<string,string>> Sheets { get; set; }

        [Category("输入")]
        [DisplayName("目标文件路径")]
        [RequiredArgument]
        public InArgument<string> DestinationPath { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var soursePath = SoursePath.Get(context);
            var sheets = Sheets.Get(context);
            var destinationPath = DestinationPath.Get(context);            
            Workbook destinationWorkbook = null;
            Workbook sourseWorkbook = new Workbook(soursePath);
            if (!File.Exists(destinationPath))
            {
                destinationWorkbook = new Workbook();
                //destinationWorkbook.Worksheets.Add("Sheet1");
                destinationWorkbook.Save(destinationPath);
            }
            destinationWorkbook = new Workbook(destinationPath);
            foreach (KeyValuePair<string,string> item in sheets)
            {
                destinationWorkbook.Worksheets.Add(item.Key).Copy(sourseWorkbook.Worksheets[item.Value]);
            }
            destinationWorkbook.Worksheets.RemoveAt("Sheet1");
            destinationWorkbook.Save(destinationPath);






        }
    }
}
