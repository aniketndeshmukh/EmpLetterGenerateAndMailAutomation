using Dapper;
using EmpLetterGenerate.Models;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using System.Net;
using System.Net.Mail;

namespace EmpLetterGenerate.Controllers
{
    public class LetterController : Controller
    {
        private readonly string connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;

        // GET: Letter
        public ActionResult Index()
        {
            using (var db = new SqlConnection(connectionString))
            {
                var letterTypes = db.Query<string>("SELECT LetterType FROM LetterTemplate").ToList();
                ViewBag.LetterTypes = letterTypes.Select(type => new SelectListItem { Value = type, Text = type }).ToList();
            }
            return View();
        }

       [HttpPost]
        public ActionResult GenerateLetter(LetterModel model)
        {
            if (string.IsNullOrEmpty(model.EmpNo) || string.IsNullOrEmpty(model.LetterType))
            {
                TempData["Error"] = "Employee No and Letter Type are required.";
                return RedirectToAction("Index");
            }

            using (var db = new SqlConnection(connectionString))
            {
                var empData = db.QueryFirstOrDefault<EmpData>("SELECT * FROM EmpData WHERE EmpId = @EmpId", new { EmpId = model.EmpNo });
                if (empData == null)
                {
                    TempData["Error"] = "Employee not found.";
                    return RedirectToAction("Index");
                }
                var letterTypes = db.Query<string>("SELECT DISTINCT LetterType FROM LetterTemplate").ToList();
                ViewBag.LetterTypes = letterTypes.Any()
                    ? letterTypes.Select(type => new SelectListItem { Value = type, Text = type }).ToList()
                    : new List<SelectListItem> { new SelectListItem { Value = "", Text = "No Letter Types Available" } };
                var letterTemplate = db.QueryFirstOrDefault<LetterTemplate>("SELECT * FROM LetterTemplate WHERE LetterType = @LetterType", new { LetterType = model.LetterType });
                if (letterTemplate == null)
                {
                    TempData["Error"] = "Letter template not found.";
                    return RedirectToAction("Index");
                }

                // Replace placeholders in letter body
                string letterContent = letterTemplate.LetterBody
                    .Replace("[EmpName]", empData.EmpName)
                    .Replace("[EmpEmail]", empData.EmpEmail)
                    .Replace("[Designation]", empData.Designation)
                    .Replace("[EmpAddress]", empData.EmpAddress)
                    .Replace("[CompanyName]", "HGS")
                    .Replace("[StartDate]", DateTime.Now.ToString("yyyy-MM-dd"))
                    .Replace("[OfferAcceptanceDeadline]", DateTime.Now.AddDays(15).ToString("yyyy-MM-dd"))
                    .Replace("[LastWorkingDay]", DateTime.Now.AddMonths(1).ToString("yyyy-MM-dd"))
                    .Replace("[ResignationDate]", DateTime.Now.ToString("yyyy-MM-dd"))
                    .Replace("[AssetReturnDeadline]", DateTime.Now.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd"))
                    .Replace("[CompanyWebsite]", "hgs.com");

                string letterSubject = letterTemplate.LetterSubject
                    .Replace("[EmpName]", empData.EmpName)
                    .Replace("[Designation]", empData.Designation);

                // Generate PDF
                string pdfPath = Server.MapPath("~/GeneratedLetters/");
                if (!Directory.Exists(pdfPath))
                    Directory.CreateDirectory(pdfPath);

                string fileName = $"Letter_{empData.EmpId}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
                string fullPath = Path.Combine(pdfPath, fileName);

                using (FileStream fs = new FileStream(fullPath, FileMode.Create))
                {
                    Document doc = new Document();
                    PdfWriter.GetInstance(doc, fs);
                    doc.Open();
                    doc.Add(new Paragraph(letterSubject, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16)));
                    doc.Add(new Paragraph("\n"));
                    doc.Add(new Paragraph(letterContent, FontFactory.GetFont(FontFactory.HELVETICA, 12)));
                    doc.Close();
                }

                // Send Email with Attachment
                SendEmail(empData.EmpEmail, letterSubject, letterContent, fullPath);

                ViewBag.FileName = fileName;
                ViewBag.LetterContent = letterContent;
                ViewBag.LetterSubject = letterSubject;

                return View("Index", model);
            }
        }

        private void SendEmail(string toEmail, string subject, string body, string attachmentPath)
        {
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("your-email@example.com"); 
                    mail.To.Add(toEmail);
                    mail.Subject = subject;
                    mail.Body = body;
                    mail.IsBodyHtml = true;

                    // Attach the PDF
                    Attachment attachment = new Attachment(attachmentPath);
                    mail.Attachments.Add(attachment);

                    using (SmtpClient smtp = new SmtpClient("smtp.office365.com", 587)) 
                    {
                        smtp.Credentials = new NetworkCredential("mahiprajapat@outlook.com", "Mahi@786"); // 
                        smtp.EnableSsl = true;
                        smtp.Send(mail);
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Error sending email: " + ex.Message;
            }
        }

        public ActionResult DownloadLetter(string fileName)
        {
            string pdfPath = Path.Combine(Server.MapPath("~/GeneratedLetters/"), fileName);

            if (System.IO.File.Exists(pdfPath))
            {
                byte[] fileBytes = System.IO.File.ReadAllBytes(pdfPath);
                return File(fileBytes, "application/pdf", fileName);
            }
            else
            {
                TempData["Error"] = "File not found.";
                return RedirectToAction("Index");
            }
        }
    }
}
