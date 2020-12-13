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
            LicenseManager.AddLicenseData("4O2QO-V13H3-3GO82-F08IH-BVMLQ");


            var applicantStudents = ExcellFileOperations.GetApplicantStudentsFromExcel();

            CreateApplicantForm(applicantStudents);

            var classroomList = ClassroomOperations.GetClassesWithApplicants(applicantStudents);

            CreateClassroomParticipantsForm(classroomList);

            //CreateClassroomParticipantsForm(ClassroomOperations.GetClassesWithApplicants(applicantStudents));


            Console.WriteLine("İşlem bitti...");
        }

        private static void CreateApplicantForm(List<ApplicantStudent> applicantStudents)
        {
            string pathToFile = "";


            using (PdfDocument pdf = new PdfDocument("Sample Data/sonSedat.pdf")) // doldurulabilir boş form pdf'i
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

                            case "sinavTuru":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].ExamType.ToUpper();
                                break;

                            case "sinavTarihi":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].ExamDate.ToUpper();
                                break;

                            case "sinavYeri":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].ExamBuilding.ToUpper();
                                break;

                            case "sinavSalonu":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].ExamClass.ToUpper();
                                break;

                            case "sinavGrup":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].ExamDeskNo.ToUpper();
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


        public static void CreateClassroomParticipantsForm(List<Classroom> classrooms)
        {
            foreach (var classroom in classrooms)
            {
                using (PdfDocument pdf = new PdfDocument("Sample Data/classroomm.pdf")) // doldurulabilir boş form pdf'i
                {
                    List<PdfControl> controls = pdf.GetControls().ToList();

                    PdfCanvas canvas = pdf.Pages[0].Canvas;
                    double axisy = 87.5;
                    int examGroup = 0;
                    int k = 1;



                    PdfTextBox classNoTBox = pdf.GetControl("classNo") as PdfTextBox;
                    classNoTBox.FontSize = 20;
                    Debug.Assert(classNoTBox != null);
                    classNoTBox.Text = classroom.Name.ToUpper();

                    for (int i = 0; i < classroom.ExamDeskCount; i++)
                    {
                        using (Bitmap bitmap =
                            new Bitmap($"Outputs/Photos/{classroom.ApplicantStudents[i].IdentityNo}.jpg"))
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


                        for (int j = 0; j < classroom.ExamDeskCount; j++)
                        {
                            cname = $"p{k.ToString()}name";
                            PdfTextBox nameTextBox = pdf.GetControl(cname) as PdfTextBox;
                            nameTextBox.FontSize = 6;
                            Debug.Assert(nameTextBox != null);
                            nameTextBox.Text = classroom.ApplicantStudents[i].Name.ToUpper();

                            cname = $"p{k.ToString()}surname";
                            PdfTextBox surnameTextBox = pdf.GetControl(cname) as PdfTextBox;
                            surnameTextBox.FontSize = 6;
                            Debug.Assert(surnameTextBox != null);
                            surnameTextBox.Text = classroom.ApplicantStudents[i].Surname.ToUpper();

                            cname = $"p{k.ToString()}id";
                            PdfTextBox idTextBox = pdf.GetControl(cname) as PdfTextBox;
                            idTextBox.FontSize = 6;
                            Debug.Assert(idTextBox != null);
                            idTextBox.Text = classroom.ApplicantStudents[i].IdentityNo.ToString().ToUpper();

                            cname = $"p{k.ToString()}classNo";
                            PdfTextBox classNoTextBox = pdf.GetControl(cname) as PdfTextBox;
                            classNoTextBox.FontSize = 6;
                            Debug.Assert(classNoTextBox != null);
                            classNoTextBox.Text = classroom.Name.ToUpper();

                            cname = $"p{k.ToString()}deskNo";
                            PdfTextBox deskNoTextBox = pdf.GetControl(cname) as PdfTextBox;
                            deskNoTextBox.FontSize = 6;
                            Debug.Assert(deskNoTextBox != null);
                            deskNoTextBox.Text = classroom.ApplicantStudents[i].ExamDeskNo.ToString();

                            if (k < classroom.ExamDeskCount)
                            {
                                k++;
                                break;
                            }
                        }
                    }
                    string pathToFile = $"Outputs/{classroom.Building} {classroom.Name}.pdf";
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
                    pdf.Save(pathToFile);
                }
            }
        }

        //public static void CreateClassroomList(List<ApplicantStudent> applicantStudents)
        //{
        //    using (PdfDocument pdf = new PdfDocument("Sample Data/classroomm.pdf")) // doldurulabilir boş form pdf'i
        //    {
        //        List<PdfControl> controls = pdf.GetControls().ToList();
        //        //List<ApplicantStudent> applicantStudents = ClassroomOperations.GetApplicantListOfGivenClass(classrooms, classroomName);
        //        //List<ApplicantStudent> applicantStudents = ClassroomOperations.GetClassroomLists(classrooms);

        //        PdfCanvas canvas = pdf.Pages[0].Canvas;
        //        double axisy = 87.5;
        //        int examGroup = 0;
        //        int k = 1;

        //        PdfTextBox classNoTBox = pdf.GetControl("classNo") as PdfTextBox;
        //        classNoTBox.FontSize = 20;
        //        Debug.Assert(classNoTBox != null);

        //        var test = applicantStudents.Where(p => p.ExamBuilding == "Eğitim Fakültesi" && p.ExamClass == "AMFİ 1")
        //            .ToList();



        //        //classNoTBox.Text = applicantStudents[0].Classroom.Name.ToUpper();

        //        //for (int i = 0; i < applicantStudents.Count; i++)
        //        //{
        //        //    using (Bitmap bitmap = new Bitmap($"Outputs/Photos/{applicantStudents[i].IdentityNo}.jpg"))
        //        //    {
        //        //        PdfImage image2;

        //        //        image2 = pdf.AddImage(bitmap);

        //        //        if (i < 10)
        //        //        {
        //        //            canvas.DrawImage(image2, 32, axisy, 52, 64, 0);
        //        //            axisy += 66.5;
        //        //            if (i == 9)
        //        //            {
        //        //                axisy = 87.5;
        //        //            }
        //        //        }
        //        //        else
        //        //        {
        //        //            canvas.DrawImage(image2, 311, axisy, 46.5, 64.5, 0);
        //        //            axisy += 66.5;
        //        //        }

        //        //    }

        //        //    string cname = String.Empty;

        //        //    for (int j = 0; j < applicantStudents[i].Classroom.Capacity; j++)
        //        //    {
        //        //        cname = $"p{k.ToString()}name";
        //        //        PdfTextBox nameTextBox = pdf.GetControl(cname) as PdfTextBox;
        //        //        nameTextBox.FontSize = 6;
        //        //        Debug.Assert(nameTextBox != null);
        //        //        nameTextBox.Text = applicantStudents[i].Name.ToUpper();

        //        //        cname = $"p{k.ToString()}surname";
        //        //        PdfTextBox surnameTextBox = pdf.GetControl(cname) as PdfTextBox;
        //        //        surnameTextBox.FontSize = 6;
        //        //        Debug.Assert(surnameTextBox != null);
        //        //        surnameTextBox.Text = applicantStudents[i].Surname.ToUpper();

        //        //        cname = $"p{k.ToString()}id";
        //        //        PdfTextBox idTextBox = pdf.GetControl(cname) as PdfTextBox;
        //        //        idTextBox.FontSize = 6;
        //        //        Debug.Assert(idTextBox != null);
        //        //        idTextBox.Text = applicantStudents[i].IdentityNo.ToString().ToUpper();

        //        //        cname = $"p{k.ToString()}classNo";
        //        //        PdfTextBox classNoTextBox = pdf.GetControl(cname) as PdfTextBox;
        //        //        classNoTextBox.FontSize = 6;
        //        //        Debug.Assert(classNoTextBox != null);
        //        //        classNoTextBox.Text = applicantStudents[i].Classroom.Name.ToUpper();

        //        //        cname = $"p{k.ToString()}deskNo";
        //        //        PdfTextBox deskNoTextBox = pdf.GetControl(cname) as PdfTextBox;
        //        //        deskNoTextBox.FontSize = 6;
        //        //        Debug.Assert(deskNoTextBox != null);
        //        //        examGroup = ClassroomOperations.GetApplicantListOfGivenClass(classrooms, applicantStudents[i].Classroom.Name)
        //        //                        .FindIndex(I => I.IdentityNo == applicantStudents[i].IdentityNo) + 1;
        //        //        deskNoTextBox.Text = examGroup.ToString();

        //        //        if (k < applicantStudents[i].Classroom.Capacity)
        //        //        {
        //        //            k++;
        //        //            break;
        //        //        }
        //        //    }

        //        //}

        //        //string pathToFile = $"Outputs/{applicantStudents[0].Classroom.Name}.pdf";
        //        //pdf.FlattenControls();
        //        //pdf.Save(pathToFile);

        //    }
        //}

    }
}