using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;

namespace BitMiracle.Docotic.Pdf.Samples
{
    public class ApplicantStudent
    {
        public long IdentityNo { get; set; }
        public Guid ApplicationNo { get; set; } = Guid.NewGuid();
        public string Name { get; set; }
        public string Surname { get; set; }
        public string FatherName { get; set; }
        public string MailAddress { get; set; }
        public string ImageUrl { get; set; }
        public string ImagePath { get; set; }
        public Bitmap Image { get; set; }
        public DateTime ApplicationDate { get; set; }

        public string ExamType { get; set; }
        public string ExamBuilding { get; set; }
        public string ExamClass { get; set; }
        public string ExamDeskNo { get; set; }
        public string ExamDate => "19/12/2020";

        public int? ClassroomId { get; set; }

    }
}
