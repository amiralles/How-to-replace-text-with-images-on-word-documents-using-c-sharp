namespace TemplateProcessor {
    using System;
    using System.IO;
    using Microsoft.Office.Interop.Word;    
    using sys = System;

    class Program {
        static void Main(string[] args) {
             Console.WriteLine("Demo - Replacing text with images on word templates.");

            var app = new Application();
            try {
                //This code creates a document based on the specified template.
                var doc = app.Documents.Add(
                    Path.GetFullPath(@"Docs\foo.dotx"),
                    Visible: false);

                doc.Activate();

                //for each keyword you want to replace.
                //************************************************
                var keyword = "angus-young";
                Console.WriteLine("Replacing keyword: {0} ...",keyword);
                var sel = app.Selection;                
                sel.Find.Text = string.Format("[{0}]", keyword);                
                sel.Find.Execute(Replace: WdReplace.wdReplaceNone);
                sel.Range.Select();                

                //This code inserts the image
                var imgPath = Path.GetFullPath(string.Format(@"Img\{0}.jpg", keyword));
                sel.InlineShapes.AddPicture(
                    FileName: imgPath,
                    LinkToFile: false,
                    SaveWithDocument: true);
                //************************************************

                //finally, save the doc.
                doc.SaveAs(Path.GetFullPath(@"Docs\foo.docx"));
                doc.Close();
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
            finally {                
                app.Quit();
                sys.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);                
            }
            Console.WriteLine("Press [Enter] to exit");
            Console.ReadLine();
        }        
    }
}
