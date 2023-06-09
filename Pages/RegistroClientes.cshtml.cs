using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel.DataAnnotations;
using OfficeOpenXml;
using System;
using System.IO;

namespace AppWebPizzas.Pages
{
    public class RegistroClientesModel : PageModel
    {   //declaro variables
        [BindProperty]
        public string? Nombre { get; set; }

        [BindProperty]
        public string? Apellido { get; set; }

        [BindProperty]
        public string? Correo { get; set; }

        [BindProperty]
        public string? Telefono { get; set; }

        [BindProperty]
        public string? Usuario { get; set; }

        [BindProperty]
        public string? Contraseña { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPost()
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return Page();
                }

                string registro = $"{Nombre},{Apellido},{Correo},{Telefono},{Usuario},{Contraseña}";

                // Ruta y nombre de la carpeta compartida
                string rutaCarpetaCompartida = "\\\\Laptop-vscoqd5q\\datos pizzeria\\"; // Reemplaza con la ruta de red de la carpeta compartida

                // Ruta del archivo de Excel local
                //var rutaArchivoExcel = "C:\\Users\\sneider\\OneDrive\\Escritorio\\UNIVERSIDAD\\7 SEMESTRE\\PROGRAMACION 3\\3 CORTE\\PEDIDOS.xlsx";


                // Nombre del archivo de Excel para guardar los registros
                string nombreArchivoExcel = "REGISTROCLIENTES.xlsx";

                // Ruta completa del archivo de Excel
                string rutaArchivoExcel = Path.Combine(rutaCarpetaCompartida, nombreArchivoExcel);

                string nombreHoja = "Hoja1"; // Nombre de la hoja

                string rutaCompleta = rutaArchivoExcel + "#" + nombreHoja;

                // Guardar el registro en el archivo de Excel
                using (var package = new ExcelPackage(new FileInfo(rutaArchivoExcel)))
                {
                    ExcelWorksheet worksheet;
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        worksheet = package.Workbook.Worksheets.Add("Registros");
                        worksheet.Cells[1, 2].Value = "Nombre";
                        worksheet.Cells[1, 3].Value = "Apellido";
                        worksheet.Cells[1, 4].Value = "Correo Electrónico";
                        worksheet.Cells[1, 5].Value = "Teléfono";
                        worksheet.Cells[1, 6].Value = "Usuario";
                        worksheet.Cells[1, 7].Value = "Contraseña";
                        worksheet.Cells[1, 8].Value = "Fecha";
                        worksheet.Cells[1, 9].Value = "Hora";
                    }
                    else
                    {
                        worksheet = package.Workbook.Worksheets[0];
                    }
                    
                    int lastRow = worksheet.Dimension?.End.Row ?? 1;
                    worksheet.Cells[lastRow + 1, 2].Value = Nombre;
                    worksheet.Cells[lastRow + 1, 3].Value = Apellido;
                    worksheet.Cells[lastRow + 1, 4].Value = Correo;
                    worksheet.Cells[lastRow + 1, 5].Value = Telefono;
                    worksheet.Cells[lastRow + 1, 6].Value = Usuario;
                    worksheet.Cells[lastRow + 1, 7].Value = Contraseña;
                    worksheet.Cells[lastRow + 1, 8].Value = DateTime.Now.ToString("dd-MM-yyyy-");
                    worksheet.Cells[lastRow + 1, 9].Value = DateTime.Now.ToString("HH:mm:ss");

                    package.Save();
                }

                //mostrar un mensaje en la pagina Login si el registro se realizo correctamente
                TempData["Error"] = "El registro fue Exitoso" + "---" + "Ahora Inicie sesión";
                return RedirectToPage("/Login");
            }
            catch (NullReferenceException)
            {
                ModelState.AddModelError(string.Empty, "Ha ocurrido un error al registrar los datos.");
                return Page();
            }
        }
    }
}
