using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Channels;
using Remotion.Configuration;
using Remotion.Configuration.TypeDiscovery;

namespace BitMiracle.Docotic.Pdf.Samples
{
    public static class FillForm
    {
        public static void Main()
        {
            //LicenseManager.AddLicenseData("4PYMJ-LEFYO-VS6BC-IYM3Z-9MXX9"); // 30 günlük deneme lisansı
            LicenseManager.AddLicenseData("4O2QO-V13H3-3GO82-F08IH-BVMLQ");

            //farklı fakültelerdeki sınıf isimleri aynı olursa?
            List<Classroom> classrooms = new List<Classroom>()
            {
                new Classroom() {Id = 1, Faculty = "Yabancı Diller Y.O.", Name = "A1", Capacity = 20},
                new Classroom() {Id = 2, Faculty = "Eğitim Fakültesi",   Name = "A2", Capacity = 20},
                new Classroom() {Id = 3, Faculty = "Yabancı Diller Y.O.", Name = "A3", Capacity = 20},
                new Classroom() {Id = 4, Faculty = "Eğitim Fakültesi",   Name = "A4", Capacity = 20},
                new Classroom() {Id = 5, Faculty = "Yabancı Diller Y.O.",   Name = "A5", Capacity = 20},
                new Classroom() {Id = 6, Faculty = "Eğitim Fakültesi",   Name = "A6", Capacity = 20},
                new Classroom() {Id = 7, Faculty = "Yabancı Diller Y.O.", Name = "A7", Capacity = 20},
                new Classroom() {Id = 8, Faculty = "Eğitim Fakültesi",   Name = "A8", Capacity = 20},
                new Classroom() {Id = 9, Faculty = "Yabancı Diller Y.O.", Name = "A9", Capacity = 20},
                new Classroom() {Id = 10, Faculty = "Eğitim Fakültesi",   Name = "A10", Capacity = 20},
                new Classroom() {Id = 11, Faculty = "Yabancı Diller Y.O.",   Name = "A11", Capacity = 20},
                new Classroom() {Id = 12, Faculty = "Eğitim Fakültesi",   Name = "A12", Capacity = 20}
            };

            var applicantStudents = ExcellFileOperations.GetApplicantStudentsFromExcel(classrooms);

            CreateApplicantForm(applicantStudents, classrooms);


            foreach (var classroom in classrooms)
            {
                CreateClassroomList(classrooms, classroom.Name);
            }



            Console.WriteLine("Oluşturulan ilk belge açılıyor...");
            Process.Start($@"C:\Users\ismet\Desktop\FillForm\bin\Debug\Outputs\{applicantStudents[0].IdentityNo}{applicantStudents[0].Name}.pdf"); // path değiştir.
        }

        private static void CreateApplicantForm(List<ApplicantStudent> applicantStudents, List<Classroom> classrooms)
        {
            string pathToFile = "";

            //EditPdf(applicantStudents); // pdf'e sınav salonunu vs. girmek için, şimdilik gerekli değil 

            using (PdfDocument pdf = new PdfDocument("Sample Data/son.pdf")) // doldurulabilir boş form pdf'i
            {
                List<PdfControl> controls = pdf.GetControls().ToList();
                for (int i = 0; i < applicantStudents.Count; i++)
                {
                    PdfCanvas canvas = pdf.Pages[0].Canvas;
                   
                    using (Bitmap bitmap = applicantStudents[i].Image)
                    {
                        PdfImage image2;

                        if (bitmap is null)
                        {
                            Console.WriteLine(applicantStudents[i].Name + " " + applicantStudents[i].Surname +
                                              " Adlı kişinin fotoğrafı bulunamadı, varsayılan resim atanıyor...");
                            image2 = pdf.AddImage(new Bitmap("Sample Data/user.png"));
                        }
                        else
                        {
                            image2 = pdf.AddImage(bitmap);
                        }

                        canvas.DrawImage(image2, 461, 110, 90, 113, 0); // kişinin fotoğrafının yerleştirileceği koordinat

                    }

                    for (int j = 0; j < controls.Count; j++)
                    {
                        ((PdfTextBox)controls[j]).FontSize = 10;
                        int examGroup = 0;
                        // son.pdf dosyasındaki fillable alanları bulup, excell dosyasından gelen verileri yazdırır.

                        switch (controls[j].Name)
                        {
                            case "basvuruNo":
                                ((PdfTextBox)controls[j]).Text =
                                    applicantStudents[i].ApplicationNo.ToString().Substring(0, 8).ToUpper();
                                break;

                            case "kimlikNo":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].IdentityNo.ToString();
                                break;

                            case "ad":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].Name.ToUpper();
                                break;

                            case "soyad":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].Surname.ToUpper();
                                break;

                            //case "sinavTarihi":
                            //    ((PdfTextBox) controls[j]).Text = DateTime.Now.ToString("g");
                            //    break;

                            case "sinavYeri":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].Classroom.Faculty.ToUpper();
                                break;

                            case "sinavSalonu":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].Classroom.Name;
                                break;

                            case "sinavGrup":
                                examGroup = ClassroomOperations
                                    .GetApplicantListOfGivenClass(classrooms, applicantStudents[i].Classroom.Name)
                                    .FindIndex(I => I.IdentityNo == applicantStudents[i].IdentityNo) + 1;
                                ((PdfTextBox)controls[j]).Text = examGroup.ToString();
                                break;
                        }
                    }

                    pathToFile = $"Outputs/{applicantStudents[i].IdentityNo}.pdf";
                    pdf.ReplaceDuplicateObjects();
                    pdf.SaveOptions.Compression = PdfCompression.Flate;
                    pdf.SaveOptions.UseObjectStreams = true;
                    pdf.SaveOptions.RemoveUnusedObjects = true;
                    pdf.SaveOptions.OptimizeIndirectObjects = true;
                    pdf.SaveOptions.WriteWithoutFormatting = true;
                    pdf.RemoveStructureInformation();
                    pdf.Metadata.Basic.Clear();
                    pdf.Metadata.DublinCore.Clear();
                    pdf.Metadata.MediaManagement.Clear();
                    pdf.Metadata.Pdf.Clear();
                    pdf.Metadata.RightsManagement.Clear();
                    pdf.Metadata.Custom.Properties.Clear();
                    foreach (XmpSchema schema in pdf.Metadata.Schemas)
                    {
                        schema.Properties.Clear();
                    }
                    pdf.Info.Clear(false);
                    unembedFonts(pdf);
                    pdf.RemoveUnusedResources();
                    pdf.RemovePieceInfo();
                    pdf.FlattenControls();

                    Console.WriteLine(
                        $"{i + 1}. {applicantStudents[i].Name} {applicantStudents[i].Surname} adlı kişinin belgesi oluşturuldu.");
                    pdf.Save(pathToFile);
                }
            }
        }

        private static void unembedFonts(PdfDocument pdf)
        {
            string[] alwaysUnembedList = new string[] { "MyriadPro-Regular" };
            string[] alwaysKeepList = new string[] { "ImportantFontName", "AnotherImportantFontName" };

            foreach (PdfFont font in pdf.GetFonts())
            {
                if (!font.Embedded ||
                    font.EncodingName == "Built-In" ||
                    Array.Exists(alwaysKeepList, name => font.Name == name))
                {
                    continue;
                }

                if (font.Format == PdfFontFormat.TrueType || font.Format == PdfFontFormat.CidType2)
                {
                    GdiFontLoader loader = new GdiFontLoader();
                    byte[] fontBytes = loader.Load(font.Name, font.Bold, font.Italic);
                    if (fontBytes != null)
                    {
                        font.Unembed();
                        continue;
                    }
                }

                if (Array.Exists(alwaysUnembedList, name => font.Name == name))
                    font.Unembed();
            }
        }


        public static void CreateClassroomList(List<Classroom> classrooms, string classroomName)
        {
            using (PdfDocument pdf = new PdfDocument("Sample Data/classroomm.pdf")) // doldurulabilir boş form pdf'i
            {
                List<PdfControl> controls = pdf.GetControls().ToList();
                List<ApplicantStudent> applicantStudents = ClassroomOperations.GetApplicantListOfGivenClass(classrooms, classroomName);
                //List<ApplicantStudent> applicantStudents = ClassroomOperations.GetClassroomLists(classrooms);

                PdfCanvas canvas = pdf.Pages[0].Canvas;
                double axisy = 87.5;
                int examGroup = 0;
                int k = 1;

                PdfTextBox classNoTBox = pdf.GetControl("classNo") as PdfTextBox;
                classNoTBox.FontSize = 20;
                Debug.Assert(classNoTBox != null);
                classNoTBox.Text = applicantStudents[0].Classroom.Name.ToUpper();

                for (int i = 0; i < applicantStudents.Count; i++)
                {
                    using (Bitmap bitmap = new Bitmap($"Outputs/Photos/{applicantStudents[i].IdentityNo}.jpg"))
                    {
                        PdfImage image2;

                        image2 = pdf.AddImage(bitmap);

                        if (i < 10)
                        {
                            canvas.DrawImage(image2, 32, axisy, 52, 64, 0);
                            axisy += 66.5;
                            if (i == 9)
                            {
                                axisy = 87.5;
                            }
                        }
                        else
                        {
                            canvas.DrawImage(image2, 311, axisy, 46.5, 64.5, 0);
                            axisy += 66.5;
                        }

                    }

                    string cname = String.Empty;

                    for (int j = 0; j < applicantStudents[i].Classroom.Capacity; j++)
                    {
                        cname = $"p{k.ToString()}name";
                        PdfTextBox nameTextBox = pdf.GetControl(cname) as PdfTextBox;
                        nameTextBox.FontSize = 6;
                        Debug.Assert(nameTextBox != null);
                        nameTextBox.Text = applicantStudents[i].Name.ToUpper();

                        cname = $"p{k.ToString()}surname";
                        PdfTextBox surnameTextBox = pdf.GetControl(cname) as PdfTextBox;
                        surnameTextBox.FontSize = 6;
                        Debug.Assert(surnameTextBox != null);
                        surnameTextBox.Text = applicantStudents[i].Surname.ToUpper();

                        cname = $"p{k.ToString()}id";
                        PdfTextBox idTextBox = pdf.GetControl(cname) as PdfTextBox;
                        idTextBox.FontSize = 6;
                        Debug.Assert(idTextBox != null);
                        idTextBox.Text = applicantStudents[i].IdentityNo.ToString().ToUpper();

                        cname = $"p{k.ToString()}classNo";
                        PdfTextBox classNoTextBox = pdf.GetControl(cname) as PdfTextBox;
                        classNoTextBox.FontSize = 6;
                        Debug.Assert(classNoTextBox != null);
                        classNoTextBox.Text = applicantStudents[i].Classroom.Name.ToUpper();

                        cname = $"p{k.ToString()}deskNo";
                        PdfTextBox deskNoTextBox = pdf.GetControl(cname) as PdfTextBox;
                        deskNoTextBox.FontSize = 6;
                        Debug.Assert(deskNoTextBox != null);
                        examGroup = ClassroomOperations.GetApplicantListOfGivenClass(classrooms, applicantStudents[i].Classroom.Name)
                                        .FindIndex(I => I.IdentityNo == applicantStudents[i].IdentityNo) + 1;
                        deskNoTextBox.Text = examGroup.ToString();

                        if (k < applicantStudents[i].Classroom.Capacity)
                        {
                            k++;
                            break;
                        }
                    }

                }

                string pathToFile = $"Outputs/{applicantStudents[0].Classroom.Name}.pdf";
                pdf.FlattenControls();
                pdf.Save(pathToFile);

            }
        }

    }
}