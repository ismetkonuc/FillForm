using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BitMiracle.Docotic.Pdf.Samples
{
    public static class ClassroomOperations
    {
        public static List<ApplicantStudent> AssignClassroomToApplicantStudents(List<ApplicantStudent> applicantStudents, List<Classroom> classrooms)
        {
            IDictionary<int,int> classroomsWithCapacities = new Dictionary<int, int>(classrooms.Count); // Key: classId , Value: classCapacity
            Random random = new Random();
            int randClassroom = 0;

            for (int i = 1; i <= classrooms.Count; i++)
            {
                classroomsWithCapacities[i] = classrooms.Where(I => I.Id == i).FirstOrDefault().Capacity;
            }

            int j = 0;

            for (int i = 0; i < applicantStudents.Count; i++)
            {
                randClassroom = random.Next(0, classrooms.Count-1);

                if (classroomsWithCapacities[randClassroom+1] > 0) // ilgili sınıfın kapasitesi dolmadıysa
                {
                    classrooms[randClassroom].ApplicantStudents.Add(applicantStudents[i]);
                    applicantStudents[i].Classroom = classrooms[randClassroom];
                    classroomsWithCapacities[randClassroom+1]--;
                }
                else
                {
                    foreach (var classroom in classroomsWithCapacities)
                    {
                        if (classroom.Value>0)
                        {
                            classrooms[classroom.Key-1].ApplicantStudents.Add(applicantStudents[i]);
                            applicantStudents[i].Classroom = classrooms[classroom.Key-1];
                            classroomsWithCapacities[classroom.Key]--;
                            break;
                        }   
                    }
                }

            }

            return applicantStudents;

        }


        public static List<ApplicantStudent> GetApplicantListOfGivenClass(List<Classroom> classrooms, string classroomName)
        {
            return classrooms.Find(I => I.Name.Equals(classroomName)).ApplicantStudents.OrderBy(I=>I.IdentityNo).ToList();
        }

    }


}
