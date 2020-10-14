using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;

namespace BitMiracle.Docotic.Pdf.Samples
{
    public static class ExcellFileOperations
    {
        
        public static List<ApplicantStudent> GetApplicantStudentsFromExcel()
        {
            Console.WriteLine("Öğrenci isimleri Excel'den okunuyor...");
            // Excell dosyasını aç
            var excel = new ExcelQueryFactory();
            excel.FileName = "Sample Data/gys.xlsx";
            
            //Exceldeki kolonları ApplicantStudent'taki proplarla eşleştir
            excel.AddMapping<ApplicantStudent>(x=> x.Name, "Ad");
            excel.AddMapping<ApplicantStudent>(x=>x.Surname, "Soyad");
            excel.AddMapping<ApplicantStudent>(x=>x.IdentityNo, "Kimlik Numarası");
            excel.AddMapping<ApplicantStudent>(x=>x.FatherName, "Baba Adı");
            excel.AddMapping<ApplicantStudent>(x=>x.MailAddress, "Email");
            excel.AddMapping<ApplicantStudent>(x => x.ApplicationDate, "Timestamp");
            excel.AddMapping<ApplicantStudent>(x=> x.ImageUrl , "Fotoğraf");

            var people = from x in excel.Worksheet<ApplicantStudent>("applicantList") select x;

            //Image dosyalarını okumak için IQueryable'ı List tipine çevir
            List<ApplicantStudent> applicantStudents = people.ToList();

           
            return DownloadAllApplicantStudentsPhoto(applicantStudents);
        }

        public static List<ApplicantStudent> DownloadAllApplicantStudentsPhoto(List<ApplicantStudent> applicantStudents)
        {
            Console.WriteLine($"{applicantStudents.Count} kişi bulundu...");
            Console.WriteLine("Okunan kişilerin fotoğrafları indiriliyor...");

            FileDownloader fileDownloader = new FileDownloader();


            foreach (var applicantStudent in applicantStudents)
            {

                if (applicantStudent.Image is null)
                {
                    // aşağıdaki satıra şimdilik comment attım çünkü gerekli fotoğrafları indirildi. 
                    //fileDownloader.DownloadFile($"{applicantStudent.ImageUrl}", $"Outputs\\Photos\\{applicantStudent.IdentityNo}.jpg"); // google drive'dan foto indir


                    applicantStudent.Image = new Bitmap($"Outputs\\Photos\\{applicantStudent.IdentityNo}.jpg");
                    
                }

            }

            return applicantStudents;
        }
    }
}
