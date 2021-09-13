using TemplateEngine.Docx;
using System.IO;
using System;

namespace DocGeneratorExamlle
{
    class Program
    {
        static string templateFile = " ";
        static string outputFile = " ";
        static string image1 = " ";
        static string image2 = " ";
        static string image3 = " ";

        static void Main(string[] args)
        {
            File.Copy(templateFile, outputFile, true);

            TableContent table = new("Table", new TableRowContent(new FieldContent("Table Column 1", "Value 1"), new FieldContent("Table Column 2", "Value 2"), new FieldContent("Table Column 3", "Value 3)")));
            ImageContent imageContent1 = new ImageContent("Image", File.ReadAllBytes(image1));
            ImageContent imageContent2 = new ImageContent("Image", File.ReadAllBytes(image2));
            ImageContent imageContent3 = new ImageContent("Image", File.ReadAllBytes(image3));

            ListContent list = new("List");

            list.AddItem(new ListItemContent(new FieldContent("List Item", "List Item 1"), table, imageContent1, imageContent2, imageContent3));

            Content valuesToFill = new(list);

            using TemplateProcessor outputDocument = new TemplateProcessor(outputFile).SetRemoveContentControls(true);
            outputDocument.FillContent(valuesToFill);
            outputDocument.SaveChanges();
        }
    }
}
