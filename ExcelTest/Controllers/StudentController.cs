using ClosedXML.Excel;
using ExcelTest.Models;
using ExcelTest.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;

namespace ExcelTest.Controllers
{
    [ApiController]
    [Route("api/[controller]")]

    public class StudentController : ControllerBase
    {
        const int maxSize = 500 * 1024; // 500 KB

        private readonly Data.ApplicationDbContext _context;

        public StudentController(Data.ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(IFormFile file, [FromQuery] bool replaceAll = false)
        {

            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            if (!file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                !file.ContentType.Contains("spreadsheet"))
                return BadRequest("The uploaded file is not a valid Excel (.xlsx) file.");

            if (file.Length > maxSize)
                return BadRequest("File size exceeds 500 KB limit.");

            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            stream.Position = 0;

            try
            {
                using var workbook = new XLWorkbook(stream);

                var schema = new List<ColumnSchema>
                {
                    new() { Name = "Name", Required = true, MaxLength = 100 },
                    new() { Name = "Age", Required = true, MinValue = 5, MaxValue = 100 },
                    new() { Name = "Email", Required = true, MaxLength = 150 },
                    new() { Name = "GraduationYear", Required = true, MinValue = 2000, MaxValue = 2100 }
                };

                var validator = new ExcelValidator(schema);
                var (validRows, errors) = validator.ValidateWorkbook(workbook);

                if (errors.Count > 0)
                    return BadRequest(new { errors });

                var students = validRows.Select(row => new Student
                {
                    Name = row["Name"].ToString(),
                    Age = Convert.ToInt32(row["Age"]),
                    Email = row["Email"].ToString(),
                    GraduationYear = Convert.ToInt32(row["GraduationYear"])
                }).ToList();

                if (replaceAll)
                {
                    _context.Students.RemoveRange(_context.Students);
                }

                _context.Students.AddRange(students);
                await _context.SaveChangesAsync();

                return Ok(new { message = $"Inserted {students.Count} students." });
            }
            catch (Exception ex)
            {
                return BadRequest(new { error = $"Failed to process Excel file: {ex.Message}" });
            }
        }

        [HttpGet("download")]
        public async Task<IActionResult> DownloadExcel()
        {
            var students = await _context.Students.ToListAsync();

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Students");

            // Add headers
            worksheet.Cell(1, 1).Value = "Id";
            worksheet.Cell(1, 2).Value = "Name";
            worksheet.Cell(1, 3).Value = "Age";
            worksheet.Cell(1, 4).Value = "Email";
            worksheet.Cell(1, 5).Value = "GraduationYear";

            // Add data
            for (int i = 0; i < students.Count; i++)
            {
                var student = students[i];
                worksheet.Cell(i + 2, 1).Value = student.Id;
                worksheet.Cell(i + 2, 2).Value = student.Name;
                worksheet.Cell(i + 2, 3).Value = student.Age;
                worksheet.Cell(i + 2, 4).Value = student.Email;
                worksheet.Cell(i + 2, 5).Value = student.GraduationYear;
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray(); // buffer the data
            var fileName = $"students-{DateTime.Now:yyyyMMddHHmmss}.xlsx";

            return File(content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        [HttpPost("upload-upsert")]
        public async Task<IActionResult> UploadExcelUpsert(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            if (!file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                return BadRequest("The uploaded file is not a valid Excel (.xlsx) file. Wrong extension.");

            if (!file.ContentType.Contains("spreadsheet"))
                return BadRequest("The uploaded file is not a valid Excel (.xlsx) file.");

            if (file.Length > maxSize)
                return BadRequest("File size exceeds 500 KB limit.");

            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            stream.Position = 0;

            try
            {
                using var workbook = new XLWorkbook(stream);

                // Define schema: Id is optional, used for upsert
                var schema = new List<ColumnSchema>
                {
                    new() { Name = "Id", Required = false, MinValue = 1 },
                    new() { Name = "Name", Required = true, MaxLength = 100 },
                    new() { Name = "Age", Required = true, MinValue = 5, MaxValue = 100 },
                    new() { Name = "Email", Required = true, MaxLength = 150 },
                    new() { Name = "GraduationYear", Required = true, MinValue = 2000, MaxValue = 2100 }
                };

                var validator = new ExcelValidator(schema);
                var (validRows, errors) = validator.ValidateWorkbook(workbook);

                if (errors.Count > 0)
                    return BadRequest(new { errors });

                var newStudents = new List<Student>();

                foreach (var row in validRows)
                {
                    int id = 0;
                    bool hasId = row.ContainsKey("Id") && int.TryParse(row["Id"].ToString(), out id) && id > 0;

                    string name = row["Name"].ToString();
                    int age = Convert.ToInt32(row["Age"]);
                    string email = row["Email"].ToString();
                    int graduationYear = Convert.ToInt32(row["GraduationYear"]);

                    if (hasId)
                    {
                        var existing = await _context.Students.FindAsync(id);
                        if (existing != null)
                        {
                            // Update existing
                            existing.Name = name;
                            existing.Age = age;
                            existing.Email = email;
                            existing.GraduationYear = graduationYear;
                            continue;
                        }

                        // Insert new with specified ID
                        newStudents.Add(new Student
                        {
                            Id = id,
                            Name = name,
                            Age = age,
                            Email = email,
                            GraduationYear = graduationYear
                        });
                    }
                    else
                    {
                        // Insert new with auto-generated ID
                        newStudents.Add(new Student
                        {
                            Name = name,
                            Age = age,
                            Email = email,
                            GraduationYear = graduationYear
                        });
                    }
                }

                if (newStudents.Any())
                    await _context.Students.AddRangeAsync(newStudents);

                await _context.SaveChangesAsync();

                return Ok(new { message = "Upsert complete.", inserted = newStudents.Count });
            }
            catch (Exception ex)
            {
                return BadRequest(new { error = $"Failed to process Excel file: {ex.Message}" });
            }
        }



    }
}
