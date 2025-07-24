using Demo_Mvc.Data;
using Demo_Mvc.Models;
using DemoMVC.Models;
using DemoMVC.Models.Process;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Demo_Mvc.Controllers
{
    public class PersonController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ExcelProcess _excelProcess = new ExcelProcess();

        public PersonController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: Person
        public async Task<IActionResult> Index()
        {
            var model = await _context.Persons.ToListAsync();
            return View(model);
        }

        // GET: Person/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Person/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("FullName,Address")] Person person)
        {
            if (ModelState.IsValid)
            {
                person.PersonId = AutoGenerateCode.GeneratePersonId(_context);
                _context.Add(person);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(person);
        }

        // GET: Person/Edit/5
        public async Task<IActionResult> Edit(string? id)
        {
            if (id == null || _context.Persons == null)
                return NotFound();

            var person = await _context.Persons.FirstOrDefaultAsync(m => m.PersonId == id);
            if (person == null)
                return NotFound();

            return View(person);
        }

        // POST: Person/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(string id, [Bind("PersonId,FullName,Address")] Person person)
        {
            if (id != person.PersonId)
                return NotFound();

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(person);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!PersonExists(person.PersonId))
                        return NotFound();
                    else
                        throw;
                }
                return RedirectToAction(nameof(Index));
            }
            return View(person);
        }

        // GET: Person/Delete/5
        public async Task<IActionResult> Delete(string? id)
        {
            if (id == null || _context.Persons == null)
                return NotFound();

            var person = await _context.Persons.FirstOrDefaultAsync(m => m.PersonId == id);
            if (person == null)
                return NotFound();

            return View(person);
        }

        // POST: Person/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(string id)
        {
            var person = await _context.Persons.FindAsync(id);
            if (person != null)
            {
                _context.Persons.Remove(person);
                await _context.SaveChangesAsync();
            }
            return RedirectToAction(nameof(Index));
        }

        private bool PersonExists(string id)
        {
            return (_context.Persons?.Any(e => e.PersonId == id)).GetValueOrDefault();
        }

        // GET: Upload
        public IActionResult Upload()
        {
            return View();
        }

        // POST: Upload Excel
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ModelState.AddModelError("", "Vui lòng chọn một file Excel để upload.");
                return View();
            }

            string fileExtension = Path.GetExtension(file.FileName);
            if (fileExtension != ".xls" && fileExtension != ".xlsx")
            {
                ModelState.AddModelError("", "Chỉ chấp nhận file Excel (.xls, .xlsx).");
                return View();
            }

            string filename = DateTime.Now.ToString("yyyyMMdd_HHmmss") + fileExtension;
            string uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Uploads", "Excels");

            if (System.IO.File.Exists(uploadsFolder))
                System.IO.File.Delete(uploadsFolder);

            Directory.CreateDirectory(uploadsFolder);

            string filePath = Path.Combine(uploadsFolder, filename);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            var dt = _excelProcess.ExcelToDataTable(filePath);
            var importedPersons = new List<Person>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var person = new Person
                {
                    PersonId = dt.Rows[i][0].ToString(),
                    FullName = dt.Rows[i][1].ToString(),
                    Address = dt.Rows[i][2].ToString()
                };

                _context.Persons.Add(person);
                importedPersons.Add(person);
            }

            await _context.SaveChangesAsync();

            ViewBag.ImportedPersons = importedPersons;
            return View();
        }
    }
}
