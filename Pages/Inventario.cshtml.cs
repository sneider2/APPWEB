using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Collections.Generic;

namespace AppWebPizzas.Pages
{
    public class InventarioModel : PageModel
    {
        public List<Pizza> Pizzas { get; set; } = new List<Pizza>(); // Lista de objetos Pizza para almacenar los datos de las pizzas

        public void OnGet()
        {
            Pizzas = new List<Pizza> // Inicialización de la lista de pizzas con valores predeterminados
            {
                new Pizza { Nombre = "Margarita", CantidadInventario = 0, CantidadVendida = 0 },
                new Pizza { Nombre = "Cuatro quesos", CantidadInventario = 0, CantidadVendida = 0 },
                new Pizza { Nombre = "Pepperoni", CantidadInventario = 0, CantidadVendida = 0 },
                new Pizza { Nombre = "Cuatro estaciones", CantidadInventario = 0, CantidadVendida = 0 },
                new Pizza { Nombre = "Champiñones", CantidadInventario = 0, CantidadVendida = 0 },
                new Pizza { Nombre = "Hawaiana", CantidadInventario = 0, CantidadVendida = 0 },
                new Pizza { Nombre = "Marinara", CantidadInventario = 0, CantidadVendida = 0 },
                new Pizza { Nombre = "Napolitana", CantidadInventario = 0, CantidadVendida = 0 }
            };

            CalcularCantidadRestante(); // Calcula la cantidad restante de porciones para cada pizza
        }

        public IActionResult OnPost()
        {
            // Procesar los datos enviados por el formulario, si es necesario
            CalcularCantidadRestante(); // Calcula la cantidad restante de porciones para cada pizza después de recibir los datos del formulario
            return Page(); // Retorna la página actual
        }

        private void CalcularCantidadRestante()
        {
            foreach (var pizza in Pizzas)
            {
                pizza.CantidadRestante = pizza.CantidadInventario - pizza.CantidadVendida; // Calcula la cantidad restante de porciones restando la cantidad vendida de la cantidad en inventario
            }
        }
    }

    public class Pizza
    {
        public string? Nombre { get; set; } // Nombre de la pizza
        public int CantidadInventario { get; set; } // Cantidad de porciones en inventario
        public int CantidadVendida { get; set; } // Cantidad de porciones vendidas
        public int CantidadRestante { get; set; } // Cantidad restante de porciones
    }
}
