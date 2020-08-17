using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Packaging; // dotnet add package System.IO.Packaging
using System.Diagnostics;
using System.Xml;
using System.Xml.Linq;

using packages;

namespace packages
{

    class Program
    {
        static string outputDirectory = "./testoutput/";
        static void Main(string[] args)
        {
            //string fileNameOrig = "./data/60489.vsdx";
            //string fileNameOrig = "data/hello-world-unsigned.docx";
            //string fileNameOrig = "./data/Visio Package.vsdx";
            string fileNameOrig = "./data/PID12235.vsdm";
            string fileNameCopy = outputDirectory + Path.GetFileName(fileNameOrig);
            try
            {
                OOXMLPackage.CreateDirectory(outputDirectory);
                File.Copy(fileNameOrig, fileNameCopy, true);
                Console.WriteLine("Opening the Package in file \"{0}\" ...", fileNameCopy);

                // We're not going to do any more than open
                // and read the list of parts in the package, although
                // we can create a package or read/write what's inside.
                using (Package fPackage = Package.Open(fileNameCopy, FileMode.Open, FileAccess.ReadWrite))
                {
                    // we need this only to get info or for debugging / testing
                    //IteratePackageParts(fPackage);
                    OOXMLPackage.UnpackPackage(fPackage, fileNameCopy+".unpacked");

                    // Get a reference to the Visio Document part contained in the file package.
                    PackagePart documentPart = OOXMLPackage.GetPackagePart(fPackage, OOXMLPackage.VISIO_CORE_DOCUMENT);
                    if (documentPart != null)
                    {
                        // Get a reference to the collection of pages in the document, 
                        // and then to the first page in the document.
                        PackagePart pagesPart = OOXMLPackage.GetPackagePart(fPackage, documentPart, OOXMLPackage.VISIO_PAGES);
                        if (pagesPart != null)
                        {
                            PackagePart page1Part = OOXMLPackage.GetPackagePart(fPackage, pagesPart, OOXMLPackage.VISIO_PAGE);
                            if (page1Part != null)
                            {
                                // Open the XML from the Page Contents part.
                                XDocument page1XML = OOXMLPackage.GetXMLFromPart(page1Part);
                                page1XML.Save(outputDirectory + "page1_orig.xml");
                                // Get all of the shapes from the page by getting
                                // all of the Shape elements from the pageXML document.
                                IEnumerable<XElement> shapesXML = OOXMLPackage.GetXElementsByName(page1XML, "Shape");
                                if (shapesXML != null)
                                {
                                    // Select a Shape element from the shapes on the page by 
                                    // its name. You can modify this code to select elements
                                    // by other attributes and their values.
                                    XElement startEndShapeXML =
                                        OOXMLPackage.GetXElementByAttribute(shapesXML, "NameU", "Start/End");
                                    if (startEndShapeXML != null)
                                    {
                                        Console.WriteLine("Found shape named \"Start/End\"");
                                        // Query the XML for the shape to get the Text element, and
                                        // return the first Text element node.
                                        IEnumerable<XElement> textElements = from element in startEndShapeXML.Elements()
                                                                             where element.Name.LocalName == "Text"
                                                                             select element;
                                        XElement textElement = textElements.ElementAt(0);
                                        // Change the shape text, leaving the <cp> element alone.
                                        textElement.LastNode.ReplaceWith("Start process\n");
                                        /* CAUTION
                                            In the previous code example, the existing shape text and the string used to replace it have the same 
                                            number of characters. Also note that the LINQ query changes the value of the last child node of the returned 
                                            element (which, in this case, is a text node). This is done to avoid changing the settings of the cp element 
                                            that is a child of the Text element. It is possible to cause file instability if you alter shape text 
                                            programmatically by overwriting all children of the Text element. As in the example above, the text 
                                            formatting is represented by cp elements under the Text element in the file. The definition of the formatting 
                                            is stored in the parent Section element. If these two pieces of information become inconsistent, then the file 
                                            may not behave as expected. Visio heals many inconsistencies, but it is better to ensure that any programmatic 
                                            changes are consistent so that you are controlling the ultimate behavior of the file.
                                        */

                                        // Insert a new Cell element in the Start/End shape that adds an arbitrary
                                        // local ThemeIndex value. This code assumes that the shape does not 
                                        // already have a local ThemeIndex cell.
                                        startEndShapeXML.Add(new XElement("Cell",
                                            new XAttribute("N", "ThemeIndex"),
                                            new XAttribute("V", "25"),
                                            new XProcessingInstruction("NewValue", "V")));
                                        // Change the shape's horizontal position on the page 
                                        // by getting a reference to the Cell element for the PinY 
                                        // ShapeSheet cell and changing the value of its V attribute.
                                        XElement pinYCellXML = OOXMLPackage.GetXElementByAttribute(
                                            startEndShapeXML.Elements(), "N", "PinY");
                                        pinYCellXML.SetAttributeValue("V", "2");
                                        // Add instructions to Visio to recalculate the entire document
                                        // when it is next opened.
                                        OOXMLPackage.RecalcDocument(fPackage);

                                        // Save the XML back to the Page Contents part.
                                        OOXMLPackage.SaveXDocumentToPart(page1Part, page1XML);
                                        fPackage.Close();
                                        Console.WriteLine("Closed modified package in file \"{0}\"", fileNameCopy);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Couldn't find shape named \"Start/End\"");
                                    }
                                }
                                // save the XML document representing the first page as XML file
                                OOXMLPackage.CreateDirectory(outputDirectory);
                                page1XML.Save(outputDirectory + "page1_possibly_modified.xml");
                            } // if (page1Part != null )
                            else
                            {
                                Console.WriteLine("Couldn't find Visio page1Part.");
                            }
                        } // if (pagesPart != null)
                        else
                        {
                            Console.WriteLine("Couldn't find Visio pagesPart.");
                        }
                    } // if (documentPart != null)
                    else
                    {
                        Console.WriteLine("Couldn't find Visio documentPart. Trying to get Office document ...");
                        documentPart = OOXMLPackage.GetPackagePart(fPackage, OOXMLPackage.CORE_DOCUMENT);
                        if (documentPart != null)
                        {
                            Console.WriteLine("Found Office documentPart with URI \"{0}\"", documentPart.Uri);
                        }
                        else
                        {
                            Console.WriteLine("Couldn't find Office documentPart.");
                        }
                    }
                } // using (Package fPackage = Package.Open(fileNameCopy, FileMode.Open, FileAccess.Read))

                // using (Package fPackage = Package.Open(fileNameCopy, FileMode.Open, FileAccess.Read))
                // {
                //     string targetDir = outputDirectory + Path.GetFileName(fileNameCopy) + ".unpacked";
                //     UnpackPackage(fPackage, targetDir);
                // }

            }
            catch (Exception err)
            {
                Console.WriteLine("Error: {0}", err.Message);
                Console.WriteLine(err.StackTrace);
            }
            finally
            {
                // Console.Write("\nPress any key to continue ...");
                // Console.Write("\nBye!\n");
                // Console.ReadKey();
            }
        } // static void Main(string[] args)

    } // class Program
} // namespace packages
