using System.Xml;
using Microsoft.Office.Interop.Excel;
using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;

namespace PriceListToExcel
{
    class Program
    {
        private static string xmlFile = "IWE OFDACatalog.xml";
        private static string excelfile = "IWE-DATABASE-AUGUST-2016-v1";

        //This struct will hold the <Product> node info
        public struct ProductData
        {
            public ProductData(string code, string description, string price, string features)
                : this()
            {
                CodeData = code;
                DescriptionData = description;
                PriceData = price;
                FeatureData = features;
            }

            public string CodeData { get; private set; }

            public string DescriptionData { get; private set; }

            public string PriceData { get; private set; }

            public string FeatureData { get; private set; }

        }

        //This struct will hold the <Option> node info
        public struct OptionData
        {
            public OptionData(string featureCode, string description, string prices, string optionDesc, string nestedFeature)
                : this()
            {
                FeatureCodeData = featureCode;
                DescriptionData = description;
                PriceData = prices;
                OptionDescriptionCodeData = optionDesc;
                NestedFeatureCodeData = nestedFeature;
            }
            public string DescriptionData { get; set; }

            public string PriceData { get; private set; }

            public string FeatureCodeData { get; private set; }

            public string OptionDescriptionCodeData { get; private set; }

            public string NestedFeatureCodeData { get; private set; }

        }
        //Hold the product xml info
        private static List<ProductData> prodList = new List<ProductData>();
        //Hold the option xml info
        private static List<OptionData> optionList = new List<OptionData>();
        static void Main(string[] args)
        {
            GetProductInfo();
            GetOptionInfo();
            PopulateExcelSheet();
        }
        //Get Product info from xml file and store it into a list
        private static void GetProductInfo()
        {
            string code = "", description = "", feature = "", price = "";
            int i = 0, f = 0;
            using (XmlReader oldCatReader = XmlReader.Create(xmlFile))
            {
                while (oldCatReader.Read())
                {
                    //Read the <Product> node
                    if (oldCatReader.Name == "Product")
                    {
                        i = 0;
                        f = 0;
                        using (var subtree = oldCatReader.ReadSubtree())
                        {
                            //Get the Code and all Features associated with that product
                            while (subtree.Read())
                            {
                                if (subtree.Name == "Code")
                                {
                                    code = subtree.ReadElementContentAsString();

                                }
                                if (subtree.Name == "Description")
                                {
                                    description = subtree.ReadElementContentAsString();
                                }
                                if (subtree.Name == "Price")
                                {
                                    using (var subtree2 = oldCatReader.ReadSubtree())
                                    {
                                        while (subtree2.Read())
                                        {
                                            if (subtree2.Name == "Value")
                                            {
                                                //Get all of the different prices (different currencies and year)
                                                price = subtree2.ReadElementContentAsString();
                                                i++;
                                            }
                                        }
                                    }
                                }
                                if (subtree.Name == "Features")
                                {
                                    using (var subtree2 = oldCatReader.ReadSubtree())
                                    {
                                        while (subtree2.Read())
                                        {
                                            if (subtree2.Name == "FeatureRef")
                                            {
                                                //Get the Feature value
                                                feature = subtree2.ReadElementContentAsString();
                                                f++;
                                            }
                                        }
                                    }
                                }
                            }
                            prodList.Add(new ProductData(code, description, price, feature));
                        }
                    }
                }
                oldCatReader.Close();
            }
            prodList = prodList.Distinct().ToList();
        }

        private static void GetOptionInfo()
        {
            string code = "", price = "";
            string descriptions = "", nestedFeature = "", optionDesc = "";
            //Check for options
            using (XmlReader oldCatReader = XmlReader.Create(xmlFile))
            {
                while (oldCatReader.Read())
                {
                    //Read the <Option> node
                    if (oldCatReader.Name == "Feature")
                    {
                        using (var subtree = oldCatReader.ReadSubtree())
                        {
                            //Get the Code and price associated with that Option
                            while (subtree.Read())
                            {
                                if (subtree.Name == "Code")
                                {
                                    code = subtree.ReadElementContentAsString();
                                }
                                if (subtree.Name == "Description")
                                {
                                    descriptions = subtree.ReadElementContentAsString();
                                }
                                if (subtree.Name == "Option")
                                {
                                    using (var subtree2 = oldCatReader.ReadSubtree())
                                    {
                                        while (subtree2.Read())
                                        {
                                            if (subtree2.Name == "Description")
                                            {
                                                optionDesc = subtree2.ReadElementContentAsString();
                                            }
                                            if (subtree2.Name == "Features")
                                            {
                                                using (var subtree3 = subtree2.ReadSubtree())
                                                {
                                                    while (subtree3.Read())
                                                    {
                                                        if (subtree3.Name == "FeatureRef")
                                                        {
                                                            nestedFeature = subtree2.ReadElementContentAsString();
                                                        }
                                                    }
                                                }
                                            }
                                            if (subtree2.Name == "OptionPrice")
                                            {
                                                using (var subtree3 = subtree2.ReadSubtree())
                                                {
                                                    while (subtree3.Read())
                                                    {
                                                        if (subtree3.Name == "Value")
                                                        {
                                                            //Get the first price
                                                            price = subtree3.ReadElementContentAsString();
                                                            if (price != "0")
                                                            {
                                                                //descriptions = descriptions.Distinct().ToList();
                                                                optionList.Add(new OptionData(code, descriptions, price, optionDesc, nestedFeature));
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            //descriptions.Clear();
                        }
                    }
                }
                oldCatReader.Close();
            }
            optionList = optionList.Distinct().ToList();
        }
        //Once we have all the product and option info populate the excel sheet
        private static void PopulateExcelSheet()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = true;
            string workbookPath = (@"D:\Documents\visual studio 2015\Projects\PriceListToExcel\PriceListToExcel\bin\Debug\IWE-DATABASE-AUGUST-2016-v1.xlsx");
            Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Open(workbookPath,
                0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = true;

            Worksheet ws = (Worksheet)wb.Worksheets[1];
            ws.Name = excelfile;
            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            var range = ws.get_Range("A2", "A2");

            if (range == null)
            {
                Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            }

            List<string> descs = new List<string>();
            foreach (OptionData option in optionList)
            {
                if(option.NestedFeatureCodeData.Any(x => char.IsLetterOrDigit(x)) && option.DescriptionData.Any(x => char.IsLetterOrDigit(x)))
                {
                    descs.Add(option.OptionDescriptionCodeData);
                }
                else
                {
                    descs.Add(option.DescriptionData);
                }
            }

            int column = 11;// K column for options

            //Delete duplicate options and populate excel columns
            descs = descs.Distinct().ToList();

            foreach (string item in descs)
            {
                //Place option into an excel column
                ws.Cells[1, column] = item;
                column++;
            }

            //now the list
            string productCell, descriptionCell, listPriceCell;
            int row = 2;
            var rows = ws.get_Range("K1", "HW1");
            bool priceSet = false;
            string lastDescription = "", lastOptionDescription = "";

            foreach (ProductData product in prodList)
            {
                //Product
                productCell = "A" + row.ToString();
                range = ws.get_Range(productCell, productCell);
                range.Value2 = product.CodeData;
                //Description
                descriptionCell = "E" + row.ToString();
                range = ws.get_Range(descriptionCell, descriptionCell);
                range.Value2 = product.DescriptionData;
                //List Price
                listPriceCell = "J" + row.ToString();
                range = ws.get_Range(listPriceCell, listPriceCell);
                range.Value2 = product.PriceData;
                optionList = optionList.Distinct().ToList();
                lastDescription = "";
                lastOptionDescription = "";
                foreach (OptionData option in optionList)
                {
                    if (product.FeatureData == option.FeatureCodeData)
                    {
                        try
                        {
                            if (float.Parse(option.PriceData) > 0)
                            {
                                foreach (Range c in rows)
                                {

                                    foreach (string desc in descs)
                                    {
                                        if (c.Value2 == desc && c.Value2 != lastDescription && c.Value2 == option.DescriptionData)
                                        {
                                            ws.Cells[row, c.Column] = option.PriceData;
                                            lastDescription = c.Value2;
                                            break;
                                        }
                                        else if (c.Value2 == desc && c.Value2 != lastOptionDescription && c.Value2 == option.OptionDescriptionCodeData)
                                        {
                                            ws.Cells[row, c.Column] = option.PriceData;
                                            lastOptionDescription = c.Value2;
                                            break;
                                        }

                                    }
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }
                        catch (FormatException f)
                        {
                            continue;
                        }
                    }
                }
                ++row;//Next row
            }//end for each product*/
        }
        //Get the feature code by description
        private static List<string> GetFeatureCodesByDesc(List<OptionData> aList, string optionDesc)
        {
            List<string> codes = new List<string>();
            foreach (OptionData option in aList)
            {
                if (option.DescriptionData == optionDesc)
                {
                    codes.Add(option.FeatureCodeData);
                }
            }
            return codes;
        }
    }
}
