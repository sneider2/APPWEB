using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using System.IO;

namespace AppWebPizzas.Pages
{
    public class LoginModel : PageModel
    {
        [BindProperty]
        public string? Usuario { get; set; }

        [BindProperty]
        public string? Contrasena { get; set; }

        [TempData]
        public string? ErrorMessage { get; set; }

        public IActionResult OnPost()
        {   
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

            using (var package = new ExcelPackage(new FileInfo(rutaArchivoExcel)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[nombreHoja];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // Comienza en la fila 2, asumiendo que la primera fila es la cabecera
                {
                    string? storedUsuario = worksheet.Cells[row, 6].Value?.ToString(); // Columna F
                    string? storedContrasena = worksheet.Cells[row, 7].Value?.ToString(); // Columna G

                    if (storedUsuario is not null && storedContrasena is not null && storedUsuario == Usuario && storedContrasena == Contrasena)
                    {
                        // Las credenciales coinciden, el inicio de sesión es exitoso
                        return RedirectToPage("/RealizarPedidos"); // Redirigir a la página de RealizarPedidos después del inicio de sesión exitoso
                    }
                   
                }
            }

             // Realizar la validación de las credenciales en el servidor
            if (Usuario == "admin89" && Contrasena == "12345")
            {
                // Las credenciales son válidas, redirigir a la página de Inventario
            return RedirectToPage("/Inventario", null, new { target = "_blank" });

            }
            


            // No se encontraron coincidencias para las credenciales, establecer el mensaje de error en la variable de estado temporal
            ErrorMessage = "Las credenciales ingresadas son incorrectas";
            return Page();
        }
    }
}
