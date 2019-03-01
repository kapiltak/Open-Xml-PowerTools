// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace OpenXmlPowerTools
{
    class WmlComparer01
    {
        static void Main(string[] args)
        {
            MemoryStream msOriginal = new MemoryStream();
            using (FileStream fs = File.OpenRead(@"C:\Users\KapilTak\Downloads\Run test (2)\Run test\First run from drafter.docx"))
                fs.CopyTo(msOriginal);
            MemoryStream msDoc1 = new MemoryStream();
            using (FileStream fs = File.OpenRead(@"C:\Users\KapilTak\Downloads\Run test (2)\Run test\Manual edits.docx"))
                fs.CopyTo(msDoc1);
            MemoryStream msDoc2 = new MemoryStream();
            using (FileStream fs = File.OpenRead(@"C:\Users\KapilTak\Downloads\Run test (2)\Run test\Second run from drafter.docx"))
                fs.CopyTo(msDoc2);

            try
            {
                var comparedFile =
                    OpenXmlPowerTools.WmlComparer.TriangularCompare(msOriginal, msDoc1, msDoc2, "Green Meadow");
                comparedFile.SaveAs(@"C:\Users\KapilTak\Downloads\Run test (2)\Run test\result.docx");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }

            //var n = DateTime.Now;
            //var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            //tempDi.Create();

            //WmlComparerSettings settings = new WmlComparerSettings();
            //WmlDocument result = WmlComparer.Compare(
            //    new WmlDocument("../../Source1.docx"),
            //    new WmlDocument("../../Source2.docx"),
            //    settings);
            //result.SaveAs(Path.Combine(tempDi.FullName, "Compared.docx"));

            //var revisions = WmlComparer.GetRevisions(result, settings);
            //foreach (var rev in revisions)
            //{
            //    Console.WriteLine("Author: " + rev.Author);
            //    Console.WriteLine("Revision type: " + rev.RevisionType);
            //    Console.WriteLine("Revision text: " + rev.Text);
            //    Console.WriteLine();
            //}
        }
    }
}
