using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using SolidEdgeConstants;
using System.Runtime.InteropServices;


namespace Solid
{
    class Saving
    {
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeFramework.Documents documents = null;
            SolidEdgePart.PartDocument partDocument = null;
            Type type = null;


            try
            {

                type = Type.GetTypeFromProgID("SolidEdge.Application");

                application = (SolidEdgeFramework.Application)
                Activator.CreateInstance(type);

                application.Visible = true;

                documents = application.Documents;
                partDocument = (PartDocument)documents.Add("SolidEdge.PartDocument");


            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }

            finally
            { 
                application = null;
                documents = null;
                partDocument = null;
            }


        }
    
    }
    
}