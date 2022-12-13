using System;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc;

namespace RISTExamOnlineProject.Models.db
{
    [ValidateAntiForgeryToken]
    public class vewOperatorAlls
    {


        
        [Key]
        [Required(ErrorMessage = "Please enter Operator No.")]
        [Display(Name = "Operator No.")]
        public string OperatorID { get; set; }
        [Required]
        [DataType(DataType.Password)]
        public string Password { get; set; }

        [Display(Name = "Name Eng")] public string NameEng { get; set; }

        [Display(Name = "Name Th")] public string NameThai { get; set; }

        public string JobTitle { get; set; }
        public string Division { get; set; }
        public string Department { get; set; }
        public string Section { get; set; }
        public string SectionCode { get; set; }
        public string Position { get; set; }
        public string OperatorGroup { get; set; }
        public string GroupName { get; set; }
        public string Email1 { get; set; }
        public string Email2 { get; set; }
        public string RFID { get; set; }
        public string Authority { get; set; }
        public bool Active { get; set; }
        public DateTime AddDate { get; set; }
        public DateTime UpdDate { get; set; }
    }
}