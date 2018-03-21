using System;
using SolidEdgeFramework;
using SolidEdgeDraft;
using SolidEdgeAssembly;
using SolidEdgePart;
using System.Linq;
using System.Reflection;

namespace Blueprint
{
    class TestClass
    {
        public Application application;  //pointer for SE application
        private SETools SEtools;  //solid edge toolbox

        public TestClass()
        {
            application = null;
            SEtools = new SETools();
        }

        public void run_test()
        {
            Console.WriteLine("test button clicked");

            try
            {
                OleMessageFilter.Register();

                // Connect to a running instance of Solid Edge
                application = SEtools.startup_SE();
                if (application != null)
                {
                    Console.WriteLine("connected to SE");

                    //    //application.Visible = false;

                    //    documents = application.Documents;

                    //    draft = (DraftDocument)documents.Open(Filename: draftFileName);

                    //    draftPropertySets = (PropertySets)draft.Properties;

                    //    properties = draftPropertySets.Item("Custom");  //grabs draft file properties

                    //    //property = properties.Item("WIDTH");  //example of grabbing specific custom file property

                    //    sections = draft.Sections;

                    //    section = sections.WorkingSection;

                    //    sectionSheets = section.Sheets;

                    //    sheets = draft.Sheets;

                    //    sheet = sectionSheets.Item(1);  //grabs sheet 1

                    //    drawingViews = sheet.DrawingViews;  //grabs every drawing view on sheet

                    //    drawingView = drawingViews.Item(1);  //main draft view

                    //    modelLink = drawingView.ModelLink;  //link to main assembly

                    //    assemblyFileName = modelLink.FileName;  //identifies main assembly

                    //    modify_assembly(assemblyFileName);  //does stuff to assembly

                    //    foreach (DrawingView dwgView in drawingViews) dwgView.Update();  //updates all drawing views
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                OleMessageFilter.Revoke();
            }
        }
    }
}
