using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BitMiracle.Docotic.Pdf.Samples
{
    public static class ClassroomOperations
    {

        public static List<Classroom> SeedClassroom()
        {
            List<Classroom> classrooms = new List<Classroom>()
            {
                new Classroom() {Name = "AMFİ 1", Building = "Eğitim Fakültesi", Id = 1, ExamType = "MÜDÜRLÜK", ExamDeskCount = 14},
                new Classroom() {Name = "DERSLİK 1", Building = "Eğitim Fakültesi", Id = 2, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 2", Building = "Eğitim Fakültesi", Id = 3, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 3", Building = "Eğitim Fakültesi", Id = 4, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 4", Building = "Eğitim Fakültesi", Id = 5, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 5", Building = "Eğitim Fakültesi", Id = 6, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 6", Building = "Eğitim Fakültesi", Id = 7, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 7", Building = "Eğitim Fakültesi", Id = 8, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 1", Building = "Turizm Fakültesi", Id = 9, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 2", Building = "Turizm Fakültesi", Id = 10, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 3", Building = "Turizm Fakültesi", Id = 11, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 4", Building = "Turizm Fakültesi", Id = 12, ExamType = "ŞEFLİK", ExamDeskCount = 12},
                new Classroom() {Name = "DERSLİK 5", Building = "Turizm Fakültesi", Id = 13, ExamType = "ŞEFLİK", ExamDeskCount = 11},
            };

            return classrooms;
        }

        public static List<Classroom> GetClassesWithApplicants(List<ApplicantStudent> applicantStudents)
        {
            var classrooms = ClassroomOperations.SeedClassroom();

            for (int i = 0; i < classrooms.Count; i++)
            {
                classrooms[i].ApplicantStudents = applicantStudents.Where(p => p.ClassroomId == i + 1).ToList();
            }

            return classrooms;
        }
             
        //public static List<ApplicantStudent> AssignClassroomToApplicantStudents(List<ApplicantStudent> applicantStudents, List<Classroom> classrooms)
        //{
        //    IDictionary<int,int> classroomsWithCapacities = new Dictionary<int, int>(classrooms.Count); // Key: classId , Value: classCapacity
        //    Random random = new Random();
        //    int randClassroom = 0;

        //    for (int i = 1; i <= classrooms.Count; i++)
        //    {
        //        classroomsWithCapacities[i] = classrooms.Where(I => I.Id == i).FirstOrDefault().Capacity;
        //    }

        //    int j = 0;

        //    for (int i = 0; i < applicantStudents.Count; i++)
        //    {
        //        randClassroom = random.Next(0, classrooms.Count-1);

        //        if (classroomsWithCapacities[randClassroom+1] > 0) // ilgili sınıfın kapasitesi dolmadıysa
        //        {
        //            classrooms[randClassroom].ApplicantStudents.Add(applicantStudents[i]);
        //            applicantStudents[i].Classroom = classrooms[randClassroom];
        //            classroomsWithCapacities[randClassroom+1]--;
        //        }
        //        else
        //        {
        //            foreach (var classroom in classroomsWithCapacities)
        //            {
        //                if (classroom.Value>0)
        //                {
        //                    classrooms[classroom.Key-1].ApplicantStudents.Add(applicantStudents[i]);
        //                    applicantStudents[i].Classroom = classrooms[classroom.Key-1];
        //                    classroomsWithCapacities[classroom.Key]--;
        //                    break;
        //                }   
        //            }
        //        }

        //    }

        //    return applicantStudents;

        //}




        //public static List<ApplicantStudent> GetApplicantListOfGivenClass(List<Classroom> classrooms, string classroomName)
        //{
        //    return classrooms.Find(I => I.Name.Equals(classroomName)).ApplicantStudents.OrderBy(I=>I.IdentityNo).ToList();
        //}

    }


}
