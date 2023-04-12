using Microsoft.AspNetCore.Mvc;
using ProyectoExcel.Models;
using System.Diagnostics;


using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using ProyectoExcel.Models.ViewModels;
using EFCore.BulkExtensions;

namespace ProyectoExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly DBPRUEBASContext _dbocontext;

        public HomeController(DBPRUEBASContext context)
        {
            _dbocontext = context;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult MostrarDatos([FromForm] IFormFile ArchivoExcel)
        {
            Stream stream = ArchivoExcel.OpenReadStream();

            IWorkbook MiExcel = null;

            if (Path.GetExtension(ArchivoExcel.FileName) == ".xlsx")
            {
                MiExcel = new XSSFWorkbook(stream);
            }
            else
            {
                MiExcel = new HSSFWorkbook(stream);
            }

            ISheet HojaExcel = MiExcel.GetSheetAt(0);

            int cantidadFilas = HojaExcel.LastRowNum;

            List<VMContacto> lista = new List<VMContacto>();

            for (int i = 1; i <= cantidadFilas; i++)
            {

                IRow fila = HojaExcel.GetRow(i);

                lista.Add(new VMContacto
                {
                    nombre = fila.GetCell(0).ToString(),
                    apellido = fila.GetCell(1).ToString(),
                    telefono = fila.GetCell(2).ToString(),
                    correo = fila.GetCell(3).ToString(),

                });
            }

            return StatusCode(StatusCodes.Status200OK, lista);
        }


        [HttpPost]
        public IActionResult EnviarDatos([FromForm] IFormFile ArchivoExcel)
        {
            Stream stream = ArchivoExcel.OpenReadStream();

            IWorkbook MiExcel = null;

            if (Path.GetExtension(ArchivoExcel.FileName) == ".xlsx")
            {
                MiExcel = new XSSFWorkbook(stream);
            }
            else
            {
                MiExcel = new HSSFWorkbook(stream);
            }

            ISheet HojaExcel = MiExcel.GetSheetAt(0);

            int cantidadFilas = HojaExcel.LastRowNum;

            List<Contacto> lista = new List<Contacto>();

            for (int i = 1; i <= cantidadFilas; i++)
            {

                IRow fila = HojaExcel.GetRow(i);

                lista.Add(new Contacto
                {
                    Nombre = fila.GetCell(0).ToString(),
                    Apellido = fila.GetCell(1).ToString(),
                    Telefono = fila.GetCell(2).ToString(),
                    Correo = fila.GetCell(3).ToString(),

                });
            }

            _dbocontext.BulkInsert(lista);

            return StatusCode(StatusCodes.Status200OK, new { mensaje = "ok"});
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
    }
}