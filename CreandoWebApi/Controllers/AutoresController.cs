using CreandoWebApi.Entidades;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using Microsoft.AspNetCore.Hosting;

using Newtonsoft.Json.Linq;
using Syncfusion.EJ2.Navigations;
using Syncfusion.XlsIO;

using System.Dynamic;
using System.IO;
using System.Text;

namespace CreandoWebApi.Controllers
{

    [ApiController]
    [Route("api/autores")]
    public class AutoresController : Controller
    {

        //[HttpGet]
        //public ActionResult<List<Autor>> Get()
        //{
        //    return new List<Autor>(){
        //    new Autor() { Id = 1,Nombre = "Felipe"},
        //new Autor(){Id=2,Nombre="Claudia"}
        //    };



        //}





        [HttpGet]

        public IActionResult Index(string button, string ConvertOptions)
        {
            ViewBag.Sheet1 = new TabHeader { Text = "datos" };
            ViewBag.Sheet2 = new TabHeader { Text = "reactivos" };

            button = "hjjh";
            ConvertOptions = "Workbook";


            if (button == null)
                return View();
            else if (button == "Input Template")
            {

                //Instantiate the spreadsheet creation engine.
                ExcelEngine excelEngine = new ExcelEngine();

                //Instantiate the Excel application object.
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the input Excel file
                FileStream stream = new FileStream("C:\\Users\\Jorge\\Desktop\\cuestionario.xlsx", FileMode.Open, FileAccess.ReadWrite);
                IWorkbook workbook = application.Workbooks.Open(stream);
                stream.Close();

                //Save the input Excel file to a stream
                MemoryStream ms = new MemoryStream();
                workbook.SaveAs(ms);
                ms.Position = 0;

                excelEngine.Dispose();

                string contentType = "Application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string fileName = "Sample.xlsx";

                //Return the input Excel file stream
                return File(ms, contentType, fileName);
            }
            else
            {
                //Instantiate the spreadsheet creation engine.
                ExcelEngine excelEngine = new ExcelEngine();

                //Instantiate the Excel application object.
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the input Excel file
                FileStream stream = new FileStream("C:\\Users\\Jorge\\Desktop\\cuestionario.xlsx", FileMode.Open, FileAccess.ReadWrite);
                IWorkbook book = application.Workbooks.Open(stream);
                stream.Close();

                //Access first worksheet
                IWorksheet worksheet = book.Worksheets[0];

                //Access a range
                IRange range = worksheet.Range["A1:H10"];

                MemoryStream jsonStream = new MemoryStream();

                if (ConvertOptions == "Workbook")
                    book.SaveAsJson(jsonStream); //Save the entire workbook as a JSON stream
                else if (ConvertOptions == "Worksheet")
                    book.SaveAsJson(jsonStream, worksheet); //Save the first worksheet as a JSON stream
                else if (ConvertOptions == "Range")
                    book.SaveAsJson(jsonStream, range); //Save the range as JSON stream

                excelEngine.Dispose();

                byte[] json = new byte[jsonStream.Length];

                //Read the Json stream and convert to a Json object
                jsonStream.Position = 0;
                jsonStream.Read(json, 0, (int)jsonStream.Length);
                string jsonString = Encoding.UTF8.GetString(json);
                JObject jsonObject = JObject.Parse(jsonString);

                //Bind the converted Json object to the DataGrid
                if (ConvertOptions == "Workbook")
                {
                    //The first worksheet in the input document is converted to a JSON object and bind to the DataGrid in the first tab.
                    //ViewBag.Tab1 = ((JArray)(jsonObject["Hoja1"])).ToObject<List<CustomDynamicObject>>();

                    ////The second worksheet in the input document is converted to a JSON object and bind to the DataGrid in the second tab.
                    //ViewBag.Tab2 = ((JArray)(jsonObject["Hoja2"])).ToObject<List<CustomDynamicObject>>();



                    JArray hoja1 = (JArray)jsonObject["datos"];
                    JArray hoja2 = (JArray)jsonObject["reactivos"];


                   

                    if (1==1)
                    {
                        string a = "df";

                    }

                    return View();
                }
                else if (ConvertOptions == "Worksheet" || ConvertOptions == "Range")
                {
                    ViewBag.Tab1 = ((JArray)(jsonObject["Hoja1"])).ToObject<List<CustomDynamicObject>>();
                }

                jsonStream.Position = 0;

                return View();
            }
        }






















    }




    public class CustomDynamicObject : DynamicObject
    {
        /// <summary>
        /// The dictionary property used store the data
        /// </summary>
        internal Dictionary<string, object> properties = new Dictionary<string, object>();
        /// <summary>
        /// Provides the implementation for operations that get member values.
        /// </summary>
        /// <param name="binder">Get Member Binder object</param>
        /// <param name="result">The result of the get operation.</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            result = default(object);

            if (properties.ContainsKey(binder.Name))
            {
                result = properties[binder.Name];
                return true;
            }
            return false;
        }
        /// <summary>
        /// Provides the implementation for operations that set member values.
        /// </summary>
        /// <param name="binder">Set memeber binder object</param>
        /// <param name="value">The value to set to the member</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            properties[binder.Name] = value;
            return true;
        }
        /// <summary>
        /// Return all dynamic member names
        /// </summary>
        /// <returns>the property name list</returns>
        public override IEnumerable<string> GetDynamicMemberNames()
        {
            return properties.Keys;
        }
    }












}
