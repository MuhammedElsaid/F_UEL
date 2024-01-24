using System;
using Spire.Doc;
using SkiaSharp;
using System.IO;
using System.Linq;
using Spire.Doc.Documents;
using YamlDotNet.Serialization;
using System.Collections.Generic;

namespace F_UEL
{
    public class StudentInfo
    {
        [YamlMember(Alias = "Name")]
        public string Name { get; set; }

        [YamlMember(Alias = "Program")]
        public string Program { get; set; }

        [YamlMember(Alias = "ASU_ID")]
        public string ASU_ID { get; set; }

        [YamlMember(Alias = "UEL_ID")]
        public int UEL_ID { get; set; }

        [YamlMember(Alias = "Semester")]
        public string Semester { get; set; }

        [YamlMember(Alias = "Academic_Year")]
        public string AcademicYear { get; set; }

        [YamlMember(Alias = "Submission_Date")]
        public string SubmissionDate { get; set; }

        [YamlMember(Alias = "Courses")]
        public List<Course> Courses { get; set; }

        [YamlMember(Alias = "Format")]
        public string Format { get; set; }
    }

    public class Course
    {
        [YamlMember(Alias = "Name")]
        public string Name { get; set; }

        [YamlMember(Alias = "ASU_Name")]
        public string ASU_Name { get; set; }

        [YamlMember(Alias = "ASU_Code")]
        public int ASU_Code { get; set; }

        [YamlMember(Alias = "UEL_Name")]
        public string UEL_Name { get; set; }

        [YamlMember(Alias = "UEL_Code")]
        public int UEL_Code { get; set; }
    }

    internal class Program
    {
        static List<string> ConvertPDFToJpeg(string pdfFilePath)
        {
            var images = PDFtoImage.Conversion.ToImages(File.Open(pdfFilePath, FileMode.Open));
            var fileName = Path.GetFileNameWithoutExtension(pdfFilePath);

            var listOfImages = new List<string>();

            int iterator = 0;
            foreach(var image in images)
            {
                var imagePath = Path.Combine(Path.GetDirectoryName(pdfFilePath), $"{fileName}{iterator}.jpg");

                using (var data = image.Encode(SKEncodedImageFormat.Jpeg, 90))
                using (var stream = File.OpenWrite(imagePath))
                {
                    listOfImages.Add(imagePath);
                    data.SaveTo(stream);
                }
                
                iterator++;
            }

            return listOfImages;
        }

        static void AppendCoursePartHeadline(Paragraph paragraph, string text)
        {
            paragraph.ApplyStyle(style.Name);
            paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;

            paragraph.AppendText(text);
        }

        static void AppendImage(Paragraph paragraph, float sectionWidth, string imagePath)
        {
            paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;

            var image = paragraph.AppendPicture(imagePath);
            image.Width = sectionWidth - 60;
        }

        static void GenerateCourseDoc(Document document, string coursePath)
        {
            int iterator = 1;
            foreach(var dir in Directory.GetDirectories(coursePath))
            {
                var section = document.AddSection();

                var coursePart = dir.Split('\\').Last();
                var paragraph = section.AddParagraph();

                AppendCoursePartHeadline(paragraph, $"{iterator}- {coursePart}");

                foreach (var file in Directory.GetFiles(Path.Combine(Environment.CurrentDirectory, dir)))
                {
                    paragraph = section.AddParagraph();

                    var fileName = Path.GetFileName(file);

                    if (fileName.Split('.').Last() == "pdf")
                    {
                        foreach(var imagePath in ConvertPDFToJpeg(file))
                        {
                            AppendImage(paragraph, section.PageSetup.ClientWidth, imagePath);
                            paragraph = section.AddParagraph();
                        }
                    }
                    else
                    {
                        AppendImage(paragraph, section.PageSetup.ClientWidth, file);
                    }
                }

                iterator++;
            }
        }

        const string OutputPath = "GeneratedCourses";
        private static ParagraphStyle style;

        static void Main(string[] args)
        {
            var studentInfo = new Deserializer().Deserialize<StudentInfo>(File.ReadAllText("config.yml"));

            //Load Document
            Document document = new Document();
            document.LoadFromFile("Template.docx");

            style = new ParagraphStyle(document);
            style.Name = "FontStyle";
            style.CharacterFormat.FontName = "Century Gothic";
            style.CharacterFormat.FontSize = 20;
            document.Styles.Add(style);

            foreach(var course in Directory.GetDirectories("Courses"))
            {
                var courseName = course.Split('\\').Last();
                GenerateCourseDoc(document, course);

                document.Replace("{Name}", studentInfo.Name, false, true);
                document.Replace("{Program}", studentInfo.Program, false, true);
                document.Replace("{ASU_ID}", studentInfo.ASU_ID, false, true);
                document.Replace("{UEL_ID}", studentInfo.UEL_ID.ToString(), false, true);
                document.Replace("{Semester}", studentInfo.Semester, false, true);
                document.Replace("{AcademicYear}", studentInfo.AcademicYear, false, true);
                document.Replace("{SubmissionDate}", studentInfo.SubmissionDate, false, true);

                var courseConfig = studentInfo.Courses.FirstOrDefault(x => x.Name == courseName);

                if(courseConfig == null)
                {
                    Console.WriteLine($"Error: Couldn't find Course {course} in config.yml file.");
                    continue;
                }

                document.Replace("{CourseASU_Name}", courseConfig.ASU_Name, false, true);
                document.Replace("{CourseASU_Code}", courseConfig.ASU_Code.ToString(), false, true);

                document.Replace("{CourseUEL_Name}", courseConfig.UEL_Name, false, true);
                document.Replace("{CourseUEL_Code}", courseConfig.UEL_Code.ToString(), false, true);

                var generatedCourseName = studentInfo.Format
                .Replace("{Name}", studentInfo.Name)
                .Replace("{Program}", studentInfo.Program)
                .Replace("{ASU_ID}", studentInfo.ASU_ID)
                .Replace("{UEL_ID}", studentInfo.UEL_ID.ToString())
                .Replace("{Semester}", studentInfo.Semester)
                .Replace("{AcademicYear}", studentInfo.AcademicYear)
                .Replace("{SubmissionDate}", studentInfo.SubmissionDate)
                .Replace("{ASU_Name}", courseConfig.ASU_Name)
                .Replace("{ASU_Code}", courseConfig.ASU_Code.ToString())
                .Replace("{UEL_Name}", courseConfig.UEL_Name)
                .Replace("{UEL_Code}", courseConfig.UEL_Code.ToString());

                document.SaveToFile(Path.Combine(OutputPath, generatedCourseName + ".docx"), FileFormat.Docx2010);
            }
        }
    }
}
