using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AnalysisServices.AdomdClient;
using System.Text;
using System.Xml;
using System.Net;
using System.Net.Http;
using Newtonsoft.Json;
using System.IO;

namespace ADOMD.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AnalyticsController : ControllerBase
    {
        private string ConnString { get; set; }
        private string DefaultQuery { get; set; }
        private string QuerySecond{ get; set; }
        public AnalyticsController()
        {
            this.ConnString = "Data Source=DESKTOP-7FOU4BG; Catalog=Cube Basics 3";
            
            //This is one big 11second query
            this.DefaultQuery= "select non empty {[Measures].[Sales Amount],[Measures].[Order Quantity],[Measures].[Tax Amt]} on columns,"+

                "( {[Dim Sales Territory 1].[Sales Territory Key].[Sales Territory Key]}, {[Dim Product 1].[Product Key].[Product Key]}," +
                "{[Due Date].[Calendar Year].&[2014]," +
                "[Due Date].[Calendar Year].&[2013],"+
                "[Due Date].[Calendar Year].&[2012],"+
                "[Due Date].[Calendar Year].&[2011]},"+

                "{[Due Date].[English Month Name].&[January]," +
                "[Due Date].[English Month Name].&[February]," +
                "[Due Date].[English Month Name].&[March]," +
                "[Due Date].[English Month Name].&[April]}"+
                ") on rows "+

                "from[Adventure Works DW2016]; ";

            this.QuerySecond = "select {[Measures].[Sales Amount]} on columns,[Due Date].[Calendar Year].[Calendar Year] on rows from[Adventure Works DW2016]";
        }

        //returns the information of all cubes in the server
        //https://localhost:*portnumber*/api/analytics/cubes
        [HttpGet]
        [Route("cubes")]
        public string GetCubeInformation()
        {
            StringBuilder cubeInformation = new StringBuilder();
            
            AdomdConnection conn = new AdomdConnection(ConnString);
            
            conn.Open();

            //Cube objects are CubeDef here
            foreach (CubeDef cube in conn.Cubes)
            {
                if (cube.Name.StartsWith('$'))
                    continue;
                cubeInformation.Append("Cube Name: " + cube.Name + '\n');
                cubeInformation.Append("Cube KPIs: " + cube.Kpis.Count + '\n');
                cubeInformation.Append("Cube Measures: " + cube.Measures.Count + '\n');
                cubeInformation.Append("Updated at " + cube.LastUpdated + '\n' + "Dimensions: " + '\n');
               
                foreach (Dimension dim in cube.Dimensions)
                {
                    cubeInformation.AppendLine(dim.Name);
                }

                cubeInformation.Append("\n\n");
            }

            conn.Close();

            return cubeInformation.ToString();
        }

        //Fetches data using a cellset from the analysis server
        //GET https://localhost:*portnumber*/api/analytics/cellset --is the complete link
        [HttpGet]
        [Route("cellset")]
        public string GetCellSetInJson()
        {
            StringBuilder result = new StringBuilder();

            AdomdConnection conn = new AdomdConnection(ConnString);
            conn.Open();

            string commandText = DefaultQuery;

            AdomdCommand adomdCommand = new AdomdCommand(commandText, conn);

            CellSet cs = adomdCommand.ExecuteCellSet();

            //insert a tab into the document
            result.Append('\t');

            //these are the tuples  in the columns i.e first row
            TupleCollection tupleColumns = cs.Axes[0].Set.Tuples;

            //foreach tuple cycle through and append  to result string
            foreach (Microsoft.AnalysisServices.AdomdClient.Tuple colValue in tupleColumns)
            {

                //get the string value of the tuple
                result.Append(colValue.Members[0].Caption + '\t');
            }

            //add blank line after the first row in the string
            result.AppendLine();

            //for each of the rows
            TupleCollection tupleRows = cs.Axes[1].Set.Tuples;

            //take each row
            for (int row = 0; row < tupleRows.Count; row++)
            {
                //add the  caption like before
                result.Append(tupleRows[row].Members[0].Caption+'\t');
                result.Append(tupleRows[row].Members[1].Caption + '\t');
                result.Append(tupleRows[row].Members[2].Caption + '\t');
                result.Append(tupleRows[row].Members[3].Caption + '\t');

                //foreach col in the row append the result
                for (int col = 0; col < tupleColumns.Count; col++)
                {
                    if (cs.Cells[col, row].FormattedValue != null || cs.Cells[col, row].FormattedValue!="")
                    {
                        result.Append(cs.Cells[col, row].FormattedValue + '\t');
                    }
                }
                result.AppendLine();
            }

            conn.Close();

            return result.ToString();

            //return new JsonResult(new { result=result.ToString() });
        }


        //NEW
        //Fetches data using a cellset from the analysis server
        //GET https://localhost:*portnumber*/api/analytics/cstojson --is the complete link
        [HttpGet]
        [Route("cstojson")]
        public IActionResult ConvertToJson()
        {
            AdomdConnection conn = new AdomdConnection(ConnString);
            conn.Open();

            string commandText = "select {[Measures].[Sales Amount]} on columns,[Due Date].[Calendar Year].[Calendar Year] on rows from[Adventure Works DW2016]";

            AdomdCommand adomdCommand = new AdomdCommand(commandText, conn);

            CellSet cs = adomdCommand.ExecuteCellSet();

            var contractResolver = new CellSetContractResolver();
            // we want Axes and Cells to be serialized from the CellSet
            contractResolver.AddInclude("CellSet", new List<string>() {
                    "Axes",
                    "Cells"
                });

            //In the Asix lets Serialize Set and Name properties
            contractResolver.AddInclude("Axis", new List<string>() {
                    "Set",
                    "Name"
                });

            //... and so on, whatever we need to include in the serialized JSON
            var settings = new JsonSerializerSettings()
            {
                ContractResolver = contractResolver,
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore
            };

            settings.Converters.Add(new CellSetJsonConverter
                         (cs.Axes[0].Set.Tuples.Count, cs.Axes[1].Set.Tuples.Count));

            string output = JsonConvert.SerializeObject
                                   (cs, Newtonsoft.Json.Formatting.Indented, settings);
           

            conn.Close();

            return new JsonResult(output);
        }

        [HttpGet]
        [Route("xmlreader")]
        public string UseXmlReader()
        {
            AdomdConnection conn = new AdomdConnection(ConnString);
            conn.Open();

            string commandText = DefaultQuery;

            AdomdCommand adomdCommand = new AdomdCommand(commandText, conn);

            //Instantiate the xmlreader
            XmlReader reader = adomdCommand.ExecuteXmlReader();
            string result = reader.ReadOuterXml();
            conn.Close();

            return result;
        }

        //API not working
        //Fetches data using a AdomdDataReader from the analysis server
        //https://localhost:*portnumber*/api/analytics/adomdreader --is the complete link
        [HttpGet]
        [Route("adomdreader")]
        public List<string> UseAdoReader()
        {
            AdomdConnection conn = new AdomdConnection(ConnString);
            conn.Open();

            string commandText = "select {[Measures].[Sales Amount]} on columns,[Due Date].[Calendar Year].[Calendar Year] on rows from[Adventure Works DW2016]";

            AdomdCommand adomdCommand = new AdomdCommand(commandText, conn);

            AdomdDataReader dataReader = adomdCommand.ExecuteReader();

            var str = new List<string>();

            while (dataReader.Read())
            {
                str.Add(dataReader[0].ToString() + "    " + dataReader[1].ToString());
            }

            dataReader.Close();
            conn.Close();

            return str;
        }
    }
}
