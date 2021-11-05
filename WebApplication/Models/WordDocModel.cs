using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace WebApplication.Models
{
    public class WordDocModel
    {

        public WordDocModel()
        {
            Items = new List<Peresdacha>(Rows);
        }
        [Required]
        [Range(1990, 2022)]
        public int Year { get; set; }

        [Required]
        [Range(1, 10)]
        public int Semester { get; set; }
        
        [Required]
        [Range(1, 5)]
        public int CourseNumber { get; set; }

        public List<Peresdacha> Items{ get; set; }
        
        [DefaultValue(1)]
        public int Rows{ get; set; }
    }

    public class Peresdacha
    {
        
        public DateTime StartDate { get; set; }

        public int Auditory { get; set; }
        public string Tutor { get; set; }
        public string Subject { get; set; }
        public int Group { get; set; }
        public double FailedPercent { get; set; }
    }
}