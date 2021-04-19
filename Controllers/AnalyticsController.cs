using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AnalysisServices.AdomdClient;
using System.Text;
using System.Xml;

namespace ADOMD.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AnalyticsController : ControllerBase
    {
        private string connString { get; set; }
        private string defaultQuery { get; set; }

        public AnalyticsController()
        {
            this.connString = "Data Source=DESKTOP-7FOU4BG; Catalog=Cube Basics 3";
            
            this.defaultQuery= "select non empty({[Measures].[Total Product Cost],[Measures].[Order Quantity],[Measures].[Sales Amount]}) on  columns," +
                                "({[Due Date].[Calendar Year].&[2011]:null}," +
                                "{[Due Date].[English Month Name].[English Month Name]}) on rows " +
                                 "from[Adventure Works DW2016]";
        }

        //returns the information of all cubes in the server
        //https://localhost:*portnumber*/api/analytics/cubes
        [HttpGet]
        [Route("cubes")]
        public string GetCubeInformation()
        {
            StringBuilder cubeInformation = new StringBuilder();
            
            AdomdConnection conn = new AdomdConnection(connString);
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
        //https://localhost:*portnumber*/api/analytics/cellset --is the complete link
        [HttpGet]
        [Route("cellset")]
        public string UseCellSet()
        {
            StringBuilder result = new StringBuilder();

            AdomdConnection conn = new AdomdConnection(connString);
            conn.Open();

            string commandText = defaultQuery;

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

                //foreach col in the row append the result
                for(int col = 0; col < tupleColumns.Count; col++)
                {
                    result.Append(cs.Cells[col, row].FormattedValue + '\t');
                }
                result.AppendLine();
            }

            conn.Close();

            return result.ToString();

            //return new JsonResult(new { result=result.ToString() });
        }

        [HttpGet]
        [Route("xmlreader")]
        public string UseXmlReader()
        {
            AdomdConnection conn = new AdomdConnection(connString);
            conn.Open();

            string commandText = defaultQuery;

            AdomdCommand adomdCommand = new AdomdCommand(commandText, conn);

            //Instantiate the xmlreader
            XmlReader reader = adomdCommand.ExecuteXmlReader();
            string result = reader.ReadOuterXml();
            conn.Close();

            return result;
        }

        //API  not  working
        //Fetches data using a AdomdDataReader from the analysis server
        //https://localhost:*portnumber*/api/analytics/adomdreader --is the complete link
        [HttpGet]
        [Route("adomdreader")]
        public IActionResult UseAdoReader()
        {
            AdomdConnection conn = new AdomdConnection(connString);
            conn.Open();

            string commandText = defaultQuery;

            AdomdCommand adomdCommand = new AdomdCommand(commandText, conn);

            AdomdDataReader dr = adomdCommand.ExecuteReader();

            StringBuilder str = new StringBuilder();

            if (dr == null)
            {
                return new JsonResult(new { message = "This is not working." });
            }

            while (dr.Read())
            {
                str.Append(dr.GetString(0) + "," + dr.GetString(1) + dr.GetString(2) + dr.GetString(3) + dr.GetString(4) + '\n'); //Error here
            }

            dr.Close();
            conn.Close();

            return new JsonResult(new { str });
        }
    }
}
