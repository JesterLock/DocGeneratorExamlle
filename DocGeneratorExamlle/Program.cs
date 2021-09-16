using TemplateEngine.Docx;
using System.IO;

namespace DocGeneratorExamlle
{
    class Program
    {
        static string templateFile = @"C:\Users\kirpichev\Downloads\Doc Generation Example.docx";
        static string outputFile = @"C:\Users\kirpichev\Downloads\Example Doc Generated.docx";
        static string image1 = @"C:\Users\kirpichev\Downloads\nustar20171030-16.jpg";
        static string image2 = @"C:\Users\kirpichev\Downloads\retro-wallpaper-35.jpg";
        static string image3 = @"C:\Users\kirpichev\Downloads\10-unsolved-mysteries-of-space.jpg";

        static void Main(string[] args)
        {
            File.Copy(templateFile, outputFile, true);

            ListContent list = new("List");

            //  создание объектов с изображениями
            ImageContent imageContent1 = new ImageContent("Image 1", File.ReadAllBytes(image1));
            ImageContent imageContent2 = new ImageContent("Image 2", File.ReadAllBytes(image2));
            ImageContent imageContent3 = new ImageContent("Image 3", File.ReadAllBytes(image3));

            //  создание контента с таблицей; в примере таблица заполняется построчно
            TableContent table1 = new("Table", 
                new TableRowContent(new FieldContent("Table Column 1", "Value 1.1.1"), new FieldContent("Table Column 2", "Value 1.1.2"), new FieldContent("Table Column 3", "Value 1.1.3")),
                new TableRowContent(new FieldContent("Table Column 1", "Value 1.2.1"), new FieldContent("Table Column 2", "Value 1.2.2"), new FieldContent("Table Column 3", "Value 1.2.3")),
                new TableRowContent(new FieldContent("Table Column 1", "Value 1.3.1"), new FieldContent("Table Column 2", "Value 1.3.2"), new FieldContent("Table Column 3", "Value 1.3.3")),
                new TableRowContent(new FieldContent("Table Column 1", "Value 1.4.1"), new FieldContent("Table Column 2", "Value 1.4.2"), new FieldContent("Table Column 3", "Value 1.4.3")));

            list.AddItem(new ListItemContent(new FieldContent("List Item", "List Item 1"), table1, imageContent1, imageContent2, imageContent3));

            TableContent table2 = new("Table", 
                new TableRowContent(new FieldContent("Table Column 1", "Value 2.1.1"), new FieldContent("Table Column 2", "Value 2.1.2"), new FieldContent("Table Column 3", "Value 2.1.3")),
                new TableRowContent(new FieldContent("Table Column 1", "Value 2.2.1"), new FieldContent("Table Column 2", "Value 2.2.2"), new FieldContent("Table Column 3", "Value 2.2.3")),
                new TableRowContent(new FieldContent("Table Column 1", "Value 2.3.1"), new FieldContent("Table Column 2", "Value 2.3.2"), new FieldContent("Table Column 3", "Value 2.3.3")),
                new TableRowContent(new FieldContent("Table Column 1", "Value 2.4.1"), new FieldContent("Table Column 2", "Value 2.4.2"), new FieldContent("Table Column 3", "Value 2.4.3")));

            list.AddItem(new ListItemContent(new FieldContent("List Item", "List Item 2"), table2, imageContent1, imageContent2, imageContent3));

            TableContent table3 = new("Table", 
                new TableRowContent(new FieldContent("Table Column 1", "Value 3.1.1"), new FieldContent("Table Column 2", "Value 3.1.2"), new FieldContent("Table Column 3", "Value 3.1.3")),
                new TableRowContent(new FieldContent("Table Column 1", "Value 3.2.1"), new FieldContent("Table Column 2", "Value 3.2.2"), new FieldContent("Table Column 3", "Value 3.2.3")),
                new TableRowContent(new FieldContent("Table Column 1", "Value 3.3.1"), new FieldContent("Table Column 2", "Value 3.3.2"), new FieldContent("Table Column 3", "Value 3.3.3")),
                new TableRowContent(new FieldContent("Table Column 1", "Value 3.4.1"), new FieldContent("Table Column 2", "Value 3.4.2"), new FieldContent("Table Column 3", "Value 3.4.3")));

            list.AddItem(new ListItemContent(new FieldContent("List Item", "List Item 3"), table3, imageContent1, imageContent2, imageContent3));

            //  создание объекта с контентом всего документа
            Content valuesToFill = new Content(list);

            using TemplateProcessor outputDocument = new TemplateProcessor(outputFile).SetRemoveContentControls(true);
            outputDocument.FillContent(valuesToFill);
            outputDocument.SaveChanges();
        }
    }
}
