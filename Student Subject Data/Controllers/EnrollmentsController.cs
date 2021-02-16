using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using WebApplication.Context;
using WebApplication.Models;

namespace WebApplication.Controllers
{
    public class EnrollmentsController : Controller
    {
        private readonly MyContext _context;

        public EnrollmentsController(MyContext context)
        {
            _context = context;
        }

        // GET: Enrollments
        public async Task<IActionResult> Index()
        {
            var myContext = _context.enrollments.Include(e => e.student).Include(e => e.subject);
            return View(await myContext.ToListAsync());
        }

        public IActionResult ExportToExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Subject Student");
                ws.Range("A1:E1").Merge();
                ws.Cell(1, 1).Value = "Report";

                ws.Cell(4, 1).Value = "student";
                ws.Cell(4, 2).Value = "subject";
                ws.Cell(4, 3).Value = "TeacherName";
                ws.Range("A4:E4").Style.Fill.BackgroundColor = XLColor.AirForceBlue;

                System.Data.DataTable dt = new System.Data.DataTable();
                SqlConnection con = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=WebAppStudent");
                SqlDataAdapter ad = new SqlDataAdapter("SELECT * from enrollments", con);
                ad.Fill(dt);
                int i = 5;

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    ws.Cell(i, 1).Value = row[1].ToString();
                    ws.Cell(i, 2).Value = row[2].ToString();
                    ws.Cell(i, 3).Value = row[3].ToString();
                    i = i + 1;
                }
                i = i - 1;

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(
                            content,
                            "application/vnd.openxmlformats-officedocument-spreadsheetml.sheet",
                            "Enrollment.xlsx"
                        );
                }
            }

            //List<Enrollment> enrollments = _context.students.Select(x => new Enrollment()).ToList();

            //ExcelPackage excelPackage = new ExcelPackage();
            //ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add("Report");

            //ws.Cells["A1"].Value = "student";
            //ws.Cells["A2"].Value = "subject";
            //ws.Cells["A3"].Value = "TeacherName";

            ////int rowStart = 7;
            ////foreach (var item in enrollments)
            ////{
            ////    if (item.SubjectId < 5)
            ////    {
            ////        ws.Row(rowStart)
            ////    }
            ////}

            //ws.Cells["A:AZ"].AutoFitColumns();
            //Response.Clear();
        }

        // GET: Enrollments/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var enrollment = await _context.enrollments
                .Include(e => e.student)
                .Include(e => e.subject)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (enrollment == null)
            {
                return NotFound();
            }

            return View(enrollment);
        }

        // GET: Enrollments/Create
        public IActionResult Create()
        {
            ViewData["StudentId"] = new SelectList(_context.students, "Id", "Name");
            ViewData["SubjectId"] = new SelectList(_context.subjects, "Id", "SubjectName");
            return View();
        }

        // POST: Enrollments/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,StudentId,SubjectId,TeacherName")] Enrollment enrollment)
        {
            if (ModelState.IsValid)
            {
                _context.Add(enrollment);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            ViewData["StudentId"] = new SelectList(_context.students, "Id", "Name", enrollment.StudentId);
            ViewData["SubjectId"] = new SelectList(_context.subjects, "Id", "SubjectName", enrollment.SubjectId);
            return View(enrollment);
        }

        // GET: Enrollments/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var enrollment = await _context.enrollments.FindAsync(id);
            if (enrollment == null)
            {
                return NotFound();
            }
            ViewData["StudentId"] = new SelectList(_context.students, "Id", "Name", enrollment.StudentId);
            ViewData["SubjectId"] = new SelectList(_context.subjects, "Id", "SubjectName", enrollment.SubjectId);
            return View(enrollment);
        }

        // POST: Enrollments/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("Id,StudentId,SubjectId,TeacherName")] Enrollment enrollment)
        {
            if (id != enrollment.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(enrollment);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!EnrollmentExists(enrollment.Id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            ViewData["StudentId"] = new SelectList(_context.students, "Id", "Name", enrollment.StudentId);
            ViewData["SubjectId"] = new SelectList(_context.subjects, "Id", "SubjectName", enrollment.SubjectId);
            return View(enrollment);
        }

        // GET: Enrollments/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var enrollment = await _context.enrollments
                .Include(e => e.student)
                .Include(e => e.subject)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (enrollment == null)
            {
                return NotFound();
            }

            return View(enrollment);
        }

        // POST: Enrollments/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var enrollment = await _context.enrollments.FindAsync(id);
            _context.enrollments.Remove(enrollment);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool EnrollmentExists(int id)
        {
            return _context.enrollments.Any(e => e.Id == id);
        }
    }
}
