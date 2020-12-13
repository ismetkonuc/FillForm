using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BitMiracle.Docotic.Pdf.Samples
{
    public class Classroom
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string ExamType { get; set; } // Müdürlük ? Şeflik ?
        public string Building { get; set; }
        public int ExamDeskCount { get; set; }

        public List<ApplicantStudent> ApplicantStudents { get; set; } = new List<ApplicantStudent>();

    }
}
