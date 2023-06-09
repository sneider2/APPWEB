using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using System;
using System.IO;

namespace AppWebPizzas.Pages
{
    public class RealizarPedidosModel : PageModel
    {

        public IActionResult OnPost()
        {
            try
            {
                // Obtener los datos del formulario
                var nombreCliente = Request.Form["NombreCliente"];
                var nombrePizza = Request.Form["NombrePizza"];
                var telefono = Request.Form["Telefono"];
                var cantidadPorciones = Request.Form["CantidadPorciones"];
                var aplicarSalsas = Request.Form["AplicarSalsas"];
                var direccionDomicilio = Request.Form["DireccionDomicilio"];
                var precio = Request.Form["Precio"];

                // Guardar los datos en un archivo de Excel
                GuardarPedidoEnExcel(nombreCliente, nombrePizza, telefono, cantidadPorciones, aplicarSalsas, direccionDomicilio, precio);
                
                //mostrar un mensaje en la pagina index si el pedido se realizo correctamente
                TempData["Error2"] = "Ha realizado el pedido correctamente, en unos minutos nos pondremos en contacto contigo para confirmar el domicilio";
                return RedirectToPage("/Index");
            }
            catch 
            {
                // mostrar un mensaje de error en la página si no se guarda el registro del pedido en excel
                TempData["Error"] = "Ocurrió un error al procesar el pedido. Por favor, inténtalo nuevamente.";
                return RedirectToPage("/RealizarPedidos");
            }
        }

        private void GuardarPedidoEnExcel(string nombreCliente, string nombrePizza, string telefono, string cantidadPorciones, string aplicarSalsas, string direccionDomicilio, string precio)
        {
            // Ruta y nombre de la carpeta compartida
            string rutaCarpetaCompartida = "\\\\Laptop-vscoqd5q\\datos pizzeria\\"; // Reemplaza con la ruta de red de la carpeta compartida

            // Ruta del archivo de Excel local
            //var rutaArchivoExcel = "C:\\Users\\sneider\\OneDrive\\Escritorio\\UNIVERSIDAD\\7 SEMESTRE\\PROGRAMACION 3\\3 CORTE\\PEDIDOS.xlsx";


            // Nombre del archivo de Excel para guardar los registros
            string nombreArchivoExcel = "PEDIDOS.xlsx";

            // Ruta completa del archivo de Excel
            string rutaArchivoExcel = Path.Combine(rutaCarpetaCompartida, nombreArchivoExcel);

            string nombreHoja = "Pedidos"; // Nombre de la hoja

            string rutaCompleta = rutaArchivoExcel + "#" + nombreHoja;

            // Crear o abrir el archivo de Excel
            using (var package = new ExcelPackage(new FileInfo(rutaArchivoExcel)))
            {
                // Obtener la hoja de trabajo
                var hojaTrabajo = package.Workbook.Worksheets["Pedidos"];

                // Obtener la última fila ocupada en la columna A
                var ultimaFila = hojaTrabajo.Dimension?.End.Row ?? 1;

                // Escribir los datos del pedido en la siguiente fila
                hojaTrabajo.Cells[ultimaFila + 1, 1].Value = nombreCliente;
                hojaTrabajo.Cells[ultimaFila + 1, 2].Value = nombrePizza;
                hojaTrabajo.Cells[ultimaFila + 1, 3].Value = telefono;
                hojaTrabajo.Cells[ultimaFila + 1, 4].Value = cantidadPorciones;
                hojaTrabajo.Cells[ultimaFila + 1, 5].Value = aplicarSalsas;
                hojaTrabajo.Cells[ultimaFila + 1, 6].Value = direccionDomicilio;
                hojaTrabajo.Cells[ultimaFila + 1, 7].Value = DateTime.Now.ToString("dd-MM-yyyy-");
                hojaTrabajo.Cells[ultimaFila + 1, 8].Value = DateTime.Now.ToString("HH:mm:ss");

                // Guardar los cambios en el archivo de Excel
                package.Save();
            }
        }
    }
}
