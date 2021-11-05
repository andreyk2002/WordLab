using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using WebApplication.Models;
using MSWord = Microsoft.Office.Interop.Word;

namespace WebApplication.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        private static Dictionary<string, Func<Peresdacha,object>> sortDict =
            new()
            {
                { "Auditory", item => item.Auditory },
                { "Tutor", item => item.Tutor},
                { "StartDate", item => item.StartDate},
                { "Group", item => item.Group},
                { "FailedPercent", item => item.FailedPercent },
                {
                    "Subject",
                    item => item.Subject
                }
            };

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            ViewBag.Rows = 5;
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public IActionResult CreateWordDocument(WordDocModel model,IFormFile file)
        {
            String filename = "template.dotx";
            var fullPath = Directory.GetCurrentDirectory() + "\\" + filename;
            file.CopyTo(target:new FileStream(fullPath, FileMode.Create));
            MSWord.Application wordApplication = new MSWord.Application();
            var wordDoc = wordApplication.Documents.Add(fullPath);
            wordDoc.Bookmarks["courseNumber"].Select();

            ReplaceBookmark(wordApplication, model.CourseNumber.ToString());

            wordDoc.Bookmarks["semesterNumber"].Select();
            ReplaceBookmark(wordApplication, model.Semester.ToString());

            wordDoc.Bookmarks["year"].Select();
            ReplaceBookmark(wordApplication, model.Year.ToString());

            FillTable(model.Items, wordDoc, wordApplication);
            wordApplication.Visible = true;
            return View("Created");
        }

        

        private void FillTable(IList<Peresdacha> modelItems, MSWord.Document wordDoc,
            MSWord.Application wordApplication)
        {
            wordDoc.Bookmarks["tableStart"].Select();
            var selection = wordApplication.Selection;
            foreach (var item in modelItems)
            {
                selection.TypeText(item.Group.ToString());
                selection.MoveRight();
                selection.TypeText(item.Subject);
                selection.MoveRight();
                selection.TypeText(item.Tutor);
                selection.MoveRight();
                selection.TypeText(item.Auditory.ToString());
                selection.MoveRight();
                selection.TypeText(item.StartDate.ToString(format:"dd.MM.yyyy HH:mm tt"));
                selection.MoveRight();
                selection.TypeText(item.FailedPercent.ToString());
                selection.InsertRowsBelow();
            }
        }

        private static void ReplaceBookmark(MSWord.Application wordApplication, string text)
        {
            var courseNumber = wordApplication.Selection;
            courseNumber.TypeBackspace();
            courseNumber.TypeText(text);
        }

        public IActionResult ChangeRows(int rows)
        {
            ViewBag.Rows = rows;
            return View("Index");
        }

        public IActionResult SortRows(string sort, string isDesc, WordDocModel model)
        {
            bool isDescending = isDesc == "on";
            ViewBag.Rows = model.Items.Count;
            var comparison = sortDict[sort];
            if (isDescending)
            {
                model.Items = new List<Peresdacha>(model.Items.OrderByDescending(comparison));
            }
            else
            {
                model.Items = new List<Peresdacha>(model.Items.OrderBy(comparison));
            }

            return View("Index", model);
        }

        public IActionResult GenerateTable(WordDocModel model, int rowsCount)
        {
            List<string> subjects = new List<string>()
                { "Computer Architecture", "Math analysis", "Web-programming", "Theory of probability" , 
                    "Differential equations", "Algebra"};
            List<string> tutors = new List<string>()
                { "Matveev", "Kash", "Krasnogirok", "Rafeenko" , "Kondratjeva"};
            Random random = new Random();
            ViewBag.Rows = model.Rows = rowsCount;
            List<Peresdacha> newItems = new List<Peresdacha>();
            for (int i = 0; i < rowsCount; i++)
            {
                Peresdacha item = new Peresdacha();
                item.Auditory = random.Next() % 625 + 1;
                item.Group = random.Next() % 15 + 1;
                item.Subject = subjects[random.Next() % subjects.Count];
                item.Tutor = tutors[random.Next() % tutors.Count];
                item.StartDate = DateTime.Parse("1/1/2001 9:00 AM");
                item.StartDate = item.StartDate.AddDays(i).AddHours(i / 2);
                item.FailedPercent = Math.Round(random.NextDouble() * 100, 2);
                newItems.Add(item);
            }

            model.Items = newItems;
            return View("Index", model);
        }
    }
}