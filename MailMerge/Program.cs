using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            var docFields = new Dictionary<string, string>()
            {
                {"FirstName","Saied" },
                {"LastName", "Hassaninia" }
            };

            var application = new Application();
            var document = new Document();

            document = application.Documents.Add(@"c:\mail-merge\template.docx");

            for (int i = 0; i < 1000; i++)
            {


                foreach (Field field in document.Fields)
                {
                    var fieldCode = field.Code;
                    var fieldText = fieldCode.Text;

                    if (fieldText.StartsWith(" MERGEFIELD"))
                    {
                        var endMerge = fieldText.IndexOf("\\");
                        var fieldNameLength = fieldText.Length - endMerge;
                        var fieldName = fieldText.Substring(11, endMerge - 11);

                        fieldName = fieldName.Trim();

                        foreach (var item in docFields)
                        {
                            if (fieldName == item.Key)
                            {
                                field.Select();
                                application.Selection.TypeText(item.Value);
                            }
                        }
                    }
                }
                document.SaveAs2(@"c:\mail-merge\saved" + i + ".docx");
            }
            

            document.Close();
            application.Quit();
        }
    }
}
