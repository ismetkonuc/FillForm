using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using Remotion.Configuration.TypeDiscovery;

namespace BitMiracle.Docotic.Pdf.Samples
{
    public static class FillForm
    {
        public static void Main()
        {
            LicenseManager.AddLicenseData("4PYMJ-LEFYO-VS6BC-IYM3Z-9MXX9"); // 30 günlük deneme lisansı

            var applicantStudents = ExcellFileOperations.GetApplicantStudentsFromExcel();

            string pathToFile ="";

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
                            Console.WriteLine(applicantStudents[i].Name + " " + applicantStudents[i].Surname + " Adlı kişinin fotoğrafı bulunamadı, varsayılan resim atanıyor...");
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
                        ((PdfTextBox) controls[j]).FontSize = 10;

                        // son.pdf dosyasındaki fillable alanları bulup, excell dosyasından gelen verileri yazdırır.

                        switch (controls[j].Name) 
                        {
                            case "basvuruNo":
                                ((PdfTextBox)controls[j]).Text = applicantStudents[i].ApplicationNo.ToString().Substring(0, 8).ToUpper();
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

                                //case "sinavYeri":
                                //    ((PdfTextBox) controls[j]).Text = "YABANCI DİLLER YO";
                                //    break;

                                //case "sinavSalonu":
                                //    ((PdfTextBox) controls[j]).Text = "DERSLİK6";
                                //    break;

                                //case "sinavGrup":
                                //    ((PdfTextBox) controls[j]).Text = "4";
                                //    break;
                        }

                    }

                    pathToFile = $"Outputs/{applicantStudents[i].IdentityNo}{applicantStudents[i].Name}.pdf";
                    //pdf.FlattenControls();
                    Console.WriteLine($"{i+1}. {applicantStudents[i].Name} { applicantStudents[i].Surname} adlı kişinin belgesi oluşturuldu.");
                    pdf.Save(pathToFile);
                }
            }

            Console.WriteLine("Oluşturulan ilk belge açılıyor...");
            Process.Start($@"C:\Users\ismet\Desktop\FillForm\bin\Debug\Outputs\{applicantStudents[0].IdentityNo}{applicantStudents[0].Name}.pdf"); // path değiştir.

            Console.ReadKey();
        }

        //burası gerekli değil şimdilik.
        public static void EditPdf(List<ApplicantStudent> applicantStudents)
        {
            using (PdfDocument pdf = new PdfDocument($"Sample Data/Test/{applicantStudents[0].IdentityNo}{applicantStudents[0].Name}.pdf")) //Sample Data/2020002216.pdf
            {
                List<PdfControl> controls = pdf.GetControls().ToList();
                
                    for (int j = 0; j < controls.Count; j++)
                    {
                        ((PdfTextBox)controls[j]).FontSize = 10;

                        switch (controls[j].Name)
                        {
                        //case "basvuruNo":
                        //    ((PdfTextBox)controls[j]).Text = applicantStudents[i].ApplicationNo.ToString().Substring(0, 8).ToUpper();
                        //    break;

                        //case "kimlikNo":
                        //    ((PdfTextBox)controls[j]).Text = applicantStudents[i].IdentityNo.ToString();
                        //    break;

                        //case "ad":
                        //    ((PdfTextBox)controls[j]).Text = applicantStudents[i].Name.ToUpper();
                        //    break;

                        //case "soyad":
                        //    ((PdfTextBox)controls[j]).Text = applicantStudents[i].Surname.ToUpper();
                        //    break;

                        case "sinavTarihi":
                            ((PdfTextBox)controls[j]).Text = DateTime.Now.ToString("g");
                            break;

                        case "sinavYeri":
                            ((PdfTextBox)controls[j]).Text = "YABANCI DİLLER YO";
                            break;

                        case "sinavSalonu":
                            ((PdfTextBox)controls[j]).Text = "DERSLİK6";
                            break;

                        case "sinavGrup":
                            ((PdfTextBox)controls[j]).Text = "4";
                            break;
                    }

                    }

                    pdf.Save($"Sample Data/Test/{applicantStudents[0].IdentityNo}{applicantStudents[0].Name}2.pdf");
                
            }
        }
    }
}