// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// It is possible to optimize DescendantContentAtoms

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Currently, the unid is set at the beginning of the algorithm.  It is used by the code that establishes correlation
// based on first rejecting// tracked revisions, then correlating paragraphs/tables.  It is requred for this algorithm
// - after finding a correlated sequence in the document with rejected revisions, it uses the unid to find the same
// paragraph in the document without rejected revisions, then sets the correlated sha1 hash in that document.
//
// But then when accepting tracked revisions, for certain paragraphs (where there are deleted paragraph marks) it is
// going to lose the unids.  But this isn't a problem because when paragraph marks are deleted, the correlation is
// definitely no longer possible.  Any paragraphs that are in a range of paragraphs that are coalesced can't be
// correlated to paragraphs in the other document via their hash.  At that point we no longer care what their unids
// are.
//
// But after that it is only used to reconstruct the tree.  It is also used in the debugging code that
// prints the various correlated sequences and comparison units - this is display for debugging purposes only.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// The key idea here is that a given paragraph will always have the same ancestors, and it doesn't matter whether the
// content was deleted from the old document, inserted into the new document, or set as equal.  At this point, we
// identify a paragraph as a sequential list of content atoms, terminated by a paragraph mark.  This entire list will
// for a single paragraph, regardless of whether the paragraph is a child of the body, or if the paragraph is in a cell
// in a table, or if the paragraph is in a text box.  The list of ancestors, from the paragraph to the root of the XML
// tree will be the same for all content atoms in the paragraph.
//
// Therefore:
//
// Iterate through the list of content atoms backwards.  When the loop sees a paragraph mark, it gets the ancestor
// unids from the paragraph mark to the top of the tree, and sets this as the same for all content atoms in the
// paragraph.  For descendants of the paragraph mark, it doesn't really matter if content is put into separate runs
// or what not.  We don't need to be concerned about what the unids are for descendants of the paragraph.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        public static WmlDocument Compare(WmlDocument source1, WmlDocument source2, WmlComparerSettings settings)
        {
            return CompareInternal(source1, source2, settings, true);
        }

        public static WmlDocument TriangularCompare(MemoryStream original, MemoryStream updatedDocumentFirst,
            MemoryStream updatedDocumentSecond, string author, bool acceptDocumentFirstChanges = true)
        {
            setBookmarks(original);
            setBookmarks(updatedDocumentFirst);
            setBookmarks(updatedDocumentSecond);

            var docOriginal = new WmlDocument("original.docx", original);
            var docNegotiated = new WmlDocument("doc1.docx", updatedDocumentFirst);
            var docUpdated = new WmlDocument("doc2.docx", updatedDocumentSecond);

            var settings = new WmlComparerSettings();
            settings.AuthorForRevisions = author;

            var result = WmlComparer.Compare(docNegotiated, docUpdated, settings);

            if (acceptDocumentFirstChanges)
            {
                var manualEdits = WmlComparer.Compare(docOriginal, docNegotiated, settings);
                var manualRevisions = WmlComparer.GetRevisions(manualEdits, new WmlComparerSettings(){AuthorForRevisions = author + "_Utility"});
                var manualRevisionRemovalList = new List<WmlComparer.WmlComparerRevision>();
                var manualRevisionAdditionList = new List<WmlComparer.WmlComparerRevision>();

                foreach (var revision in manualRevisions)
                {
                    if (revision.Text.Contains("\n"))
                    {
                        var textArray = revision.Text.Split('\n');
                        foreach (var text in textArray)
                        {
                            WmlComparer.WmlComparerRevision newRevision = new WmlComparer.WmlComparerRevision()
                            {
                                Author = revision.Author,
                                ContentXElement = revision.ContentXElement,
                                Date = revision.Date,
                                PartContentType = revision.PartContentType,
                                PartUri = revision.PartUri,
                                RevisionType = revision.RevisionType,
                                RevisionXElement = revision.RevisionXElement,
                                Text = text
                            };
                            manualRevisionAdditionList.Add(newRevision);
                        }

                        manualRevisionRemovalList.Add(revision);
                    }
                }

                foreach (var revision in manualRevisionRemovalList)
                {
                    manualRevisions.Remove(revision);
                }

                manualRevisions.AddRange(manualRevisionAdditionList);

                var qUpdates = WmlComparer.Compare(docOriginal, docUpdated, settings);
                var updateRevisions = WmlComparer.GetRevisions(qUpdates, settings);

                result = acceptManualEditsAfterTriangularCompare(result, manualRevisions, updateRevisions);
            }
            
            return result;
        }

        //private static void setBookmarks(MemoryStream stream)
        //{
        //    using (var doc1 = WordprocessingDocument.Open(stream, true))
        //    {

        //    }
        //}

        private static void setBookmarks(MemoryStream stream)
        {
            using (var doc1 = WordprocessingDocument.Open(stream, true))
            {
                //foreach (var fc in doc1.MainDocumentPart.Document.Body.Descendants<FieldChar>())
                //{
                //    fc.Dirty = false;
                //    if (fc.FieldCharType == "begin")
                //    {
                //        var run = fc.Parent;
                //        while (true)
                //        {
                //            run = run.NextSibling();
                //            if (run == null)
                //                break;

                //            foreach (var text in run.Descendants<Text>())
                //            {
                //                text.Text = "{Link}";
                //            }

                //            if (run.Descendants<FieldChar>().Count() <= 0)
                //                continue;

                //            var nextfc = run.Descendants<FieldChar>().First();

                //            if (nextfc.FieldCharType == "end")
                //                break;
                //        }
                //    }
                //}

                //foreach (BookmarkStart bkmStart in doc1.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
                //{
                //    var previous = bkmStart.PreviousSibling();

                //    while (previous.GetType() == typeof(BookmarkEnd))
                //    {
                //        var id = (previous as BookmarkEnd).Id.Value;

                //        if (id == bkmStart.Id.Value)
                //        {
                //            bkmStart.Parent.AppendChild(previous.CloneNode(true));
                //            previous.Remove();
                //            break;
                //        }

                //        previous = previous.PreviousSibling();
                //    }
                //}

                try
                {
                    List<OpenXmlElement> bkmToremove = new List<OpenXmlElement>();
                    List<OpenXmlElement> bkmToAdd = new List<OpenXmlElement>();
                    string id = "";

                    foreach (var bkmStart in doc1.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
                    {
                        id = bkmStart.Id.Value;
                        var bkmEnd = doc1.MainDocumentPart.Document.Body.Descendants<BookmarkEnd>().Where(x => x.Id.Value == id)
                            .FirstOrDefault();
                        if(bkmEnd == null)
                            bkmToremove.Add(bkmStart);
                    }

                    foreach (var bkmEnd in doc1.MainDocumentPart.Document.Body.Descendants<BookmarkEnd>())
                    {
                        id = bkmEnd.Id.Value;
                        var bkmStart = doc1.MainDocumentPart.Document.Body.Descendants<BookmarkStart>()
                            .Where(x => x.Id.Value == id).FirstOrDefault();
                        if(bkmStart == null)
                            bkmToremove.Add(bkmEnd);
                    }

                    foreach (var bkm in bkmToremove)
                        bkm.Remove();

                    bkmToremove = new List<OpenXmlElement>();

                    foreach (var element in doc1.MainDocumentPart.Document.Body.Descendants<OpenXmlElement>())
                    {
                        if (bkmToAdd.Count > 0 && element.GetType() == typeof(Paragraph))
                        {
                            var paraProps = element.Descendants<ParagraphProperties>().FirstOrDefault();
                            if (paraProps != null)
                            {
                                for (int i = bkmToAdd.Count - 1; i >= 0; i--)
                                {
                                    paraProps.InsertAfterSelf(bkmToAdd[i].CloneNode(true));
                                }

                                bkmToAdd = new List<OpenXmlElement>();
                            }
                        }

                        if (element.GetType() != typeof(BookmarkStart))
                            continue;

                        var parent = element.Parent;
                        var hasParentPara = false;
                        while (parent.GetType() != typeof(Body))
                        {
                            if (parent.GetType() == typeof(Paragraph))
                            {
                                hasParentPara = true;
                                break;
                            }

                            parent = parent.Parent;
                        }

                        if (!hasParentPara)
                        {
                            bkmToAdd.Add(element.CloneNode(true));
                            bkmToremove.Add(element);
                        }
                    }

                    foreach (var bkm in bkmToremove)
                        bkm.Remove();


                    bkmToremove = new List<OpenXmlElement>();
                    OpenXmlElement lastPara = null;
                    foreach (var element in doc1.MainDocumentPart.Document.Body.Descendants<OpenXmlElement>())
                    {
                        if (element.GetType() == typeof(Paragraph))
                            lastPara = element;

                        if (element.GetType() == typeof(BookmarkEnd))
                        {
                            id = (element as BookmarkEnd).Id.Value;
                            var bkmStart = doc1.MainDocumentPart.Document.Body.Descendants<BookmarkStart>()
                                .Where(x => x.Id.Value == id).FirstOrDefault();

                            if (bkmStart != null) //&& element.Parent != bkmStart.Parent
                            {
                                var nextElement = bkmStart.NextSibling();
                                while (nextElement != null)
                                {
                                    if (nextElement.GetType() != typeof(BookmarkStart))
                                    {
                                        nextElement.InsertBeforeSelf(element.CloneNode(true));
                                        bkmToremove.Add(element);
                                        break;
                                    }

                                    nextElement = nextElement.NextSibling();
                                }
                            }
                        }
                    }

                    foreach (var bkm in bkmToremove)
                        bkm.Remove();


                    foreach (DocumentFormat.OpenXml.Wordprocessing.Table tbl in doc1.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>())
                    {
                        var maxCellCount = 0;
                        foreach (DocumentFormat.OpenXml.Wordprocessing.TableRow tr in tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>())
                        {
                            if (maxCellCount < tr.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count())
                                maxCellCount = tr.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count();
                        }

                        foreach (DocumentFormat.OpenXml.Wordprocessing.TableRow tr in tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>())
                        {
                            if (tr.TableRowProperties != null && maxCellCount == tr.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count() &&
                                tr.TableRowProperties.Descendants<GridAfter>().Any())
                            {
                                tr.TableRowProperties.Descendants<GridAfter>().FirstOrDefault().Remove();
                                foreach (var gs in tr.Descendants<GridSpan>())
                                    gs.Remove();
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }

                doc1.Save();
            }
        }

        private static WmlDocument acceptManualEditsAfterTriangularCompare(WmlDocument result, List<WmlComparerRevision> manualRevisions, List<WmlComparerRevision> updateRevisions)
        {
            WmlDocument returnDoc = result;

            try
            {
                using (var ms = new MemoryStream())
                {
                    ms.Write(result.DocumentByteArray, 0, result.DocumentByteArray.Length);
                    using (DocumentFormat.OpenXml.Packaging.WordprocessingDocument wdDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, true))
                    {
                        //// Handle the formatting changes.
                        //List<ParagraphPropertiesChange> changes = wdDoc.MainDocumentPart.Document.Body.Descendants<ParagraphPropertiesChange>().ToList();

                        //foreach (var change in changes)
                        //{
                        //    if (!manualRevisions.Any(x => x.Date == change.Date))
                        //        continue;
                        //    change.Remove();
                        //}

                        try
                        {
                            //Correct the paragraph sequence. 
                            foreach (var insertion in wdDoc.MainDocumentPart.Document.Body.Descendants<InsertedRun>())
                            {
                                var text = insertion.InnerText;
                                var parent = insertion.Parent;
                                while (parent.GetType() != typeof(Paragraph))
                                {
                                    if (parent.Parent == null)
                                        break;
                                    parent = parent.Parent;
                                }

                                if (parent.GetType() != typeof(Paragraph))
                                    continue;

                                var parentPara = parent as Paragraph;
                                if (parentPara.ParagraphProperties == null || parentPara.ParagraphProperties.ParagraphStyleId == null)
                                    continue;

                                string styleId = parentPara.ParagraphProperties.ParagraphStyleId.Val;
                                var hasFoundDeletions = false;
                                while (true)
                                {
                                    parent = parent.PreviousSibling();
                                    if (parent == null || parent.GetType() != typeof(Paragraph) || (parent as Paragraph).ParagraphProperties == null)
                                        break;

                                    var para = new Paragraph(parent.OuterXml);
                                    foreach (var dr in para.Descendants<DeletedRun>())
                                        dr.Remove();

                                    if (para.InnerText == "")
                                    {
                                        if (para.ParagraphProperties == null || para.ParagraphProperties.ParagraphStyleId == null)
                                            continue;

                                        if(para.ParagraphProperties.ParagraphStyleId.Val != styleId)
                                            hasFoundDeletions = true;
                                        continue;
                                    }

                                    if ((parent as Paragraph).ParagraphProperties == null ||
                                        (parent as Paragraph).ParagraphProperties.ParagraphStyleId == null)
                                        break;

                                    var styleIdofPrevious = (parent as Paragraph).ParagraphProperties.ParagraphStyleId.Val;

                                    if (hasFoundDeletions && styleId == styleIdofPrevious)
                                    {
                                        var newPara = new Paragraph(parentPara.OuterXml);
                                        var firstIns = newPara.Descendants<InsertedRun>().FirstOrDefault();
                                        if (firstIns != null && newPara.ParagraphProperties != null)
                                        {
                                            var mrp = newPara.ParagraphProperties.ParagraphMarkRunProperties;
                                            if (mrp == null)
                                            {
                                                mrp = new ParagraphMarkRunProperties();
                                                newPara.ParagraphProperties.AppendChild(mrp);
                                            }

                                            if (mrp.Descendants<Inserted>().Count() < 1)
                                            {
                                                var inserted = new Inserted();
                                                inserted.Author = firstIns.Author;
                                                inserted.Id = firstIns.Id;
                                                inserted.Date = firstIns.Date;
                                                mrp.AppendChild(inserted);
                                            }
                                        }

                                        parent.InsertAfterSelf(newPara);
                                        parentPara.Remove();
                                    }

                                    break;
                                }
                            }
                        }
                        catch (Exception exc)
                        {
                            var msg = exc.Message;
                            throw exc;
                        }

                        var deletedRuns = wdDoc.MainDocumentPart.Document.Body.Descendants<DeletedRun>().ToList();
                        var insertedRuns = wdDoc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.InsertedRun>().ToList();

                        // Handle the deletions.
                        //var deletedItems = wdDoc.MainDocumentPart.Document.Body.Descendants<Deleted>().ToList();
                        //foreach (var deletion in deletedItems)
                        //{
                        //    if (manualRevisions.Any(x =>
                        //        x.Text.Trim() == deletion.InnerText.Trim() && x.Date == deletion.Date &&
                        //        x.RevisionType == WmlComparer.WmlComparerRevisionType.Inserted))
                        //    {
                        //        deletion.Remove();
                        //    }
                        //}


                        foreach (var deletion in deletedRuns)
                        {
                            if (manualRevisions.Any(x =>
                                x.Text.Trim() == deletion.InnerText.Trim() && x.Date == deletion.Date &&
                                x.RevisionType == WmlComparer.WmlComparerRevisionType.Inserted))
                            {
                                var fcStarted = false;
                                foreach (var run in deletion.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
                                {
                                    if (run.Descendants<FieldChar>().Count() > 0)
                                    {
                                        var fc = run.Descendants<FieldChar>().FirstOrDefault();
                                        if (fc.FieldCharType == FieldCharValues.Begin)
                                            fcStarted = true;
                                        else if (fc.FieldCharType == FieldCharValues.End)
                                            fcStarted = false;
                                    }

                                    var outerXML = run.OuterXml.Replace("w:delText", "w:t");
                                    if(!fcStarted)
                                        outerXML = outerXML.Replace("w:delInstrText", "w:InstrText");
                                    else
                                        outerXML = outerXML.Replace("w:delInstrText", "w:instrText");
                                    DocumentFormat.OpenXml.Wordprocessing.Run newRun = new DocumentFormat.OpenXml.Wordprocessing.Run(outerXML);
                                    deletion.InsertBeforeSelf(newRun);
                                    //if (run == deletion.FirstChild)
                                    //{
                                    //    deletion.InsertAfterSelf(newRun);
                                    //}
                                    //else
                                    //{
                                    //    deletion.NextSibling().InsertAfterSelf(newRun);
                                    //}
                                }

                                var parentPara = deletion.Parent;
                                var deleted = parentPara.Descendants<Deleted>().FirstOrDefault();
                                if (deleted != null)
                                    deleted.Remove();

                                deletion.RemoveAttribute("rsidR",
                                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                                deletion.RemoveAttribute("rsidRPr",
                                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                                var parent1 = deletion.Parent;
                                var xml1 = parent1.OuterXml;
                                deletion.RemoveAllChildren();
                                deletion.Remove();

                                continue;
                            }
                        }

                        var deletedMathControls = wdDoc.MainDocumentPart.Document.Body.Descendants<DeletedMathControl>().ToList();
                        foreach (var deletion in deletedMathControls)
                        {
                            if (manualRevisions.Any(x =>
                                x.Text.Trim() == deletion.InnerText.Trim() && x.Date == deletion.Date &&
                                x.RevisionType == WmlComparer.WmlComparerRevisionType.Inserted))
                            {
                                foreach (var run in deletion.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
                                {
                                    DocumentFormat.OpenXml.Wordprocessing.Run newRun = new DocumentFormat.OpenXml.Wordprocessing.Run(run.OuterXml.Replace("w:delText", "w:t"));

                                    if (run == deletion.FirstChild)
                                    {
                                        deletion.InsertAfterSelf(newRun);
                                    }
                                    else
                                    {
                                        deletion.NextSibling().InsertAfterSelf(newRun);
                                    }
                                }

                                var parentPara = deletion.Parent;
                                var deleted = parentPara.Descendants<Deleted>().FirstOrDefault();
                                if (deleted != null)
                                    deleted.Remove();

                                deletion.RemoveAttribute("rsidR",
                                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                                deletion.RemoveAttribute("rsidRPr",
                                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                                deletion.Remove();
                                continue;
                            }
                        }

                        //Insertions
                        //var insertedItems = wdDoc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Inserted>().ToList();
                        //foreach (var insertion in insertedItems)
                        //{
                        //    if (manualRevisions.Any(x =>
                        //        x.Text.Trim() == insertion.InnerText.Trim() && x.Date == insertion.Date &&
                        //        x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted))
                        //    {
                        //        manualRevisions.Remove(manualRevisions.Where(x => x.Text.Trim() == insertion.InnerText.Trim() && x.Date == insertion.Date && x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted).First());
                        //        var parentPara = insertion.Parent;
                        //        insertion.Remove();
                        //        if (parentPara.InnerText == string.Empty)
                        //            parentPara.Remove();
                        //    }
                        //}

                        List<InsertedRun> combinedRuns = new List<InsertedRun>();
                        List<WmlComparer.WmlComparerRevision> combinedRevisions = new List<WmlComparer.WmlComparerRevision>();
                        foreach (var insertion in insertedRuns)
                        {
                            var fcStarted = false;
                            List<string> textToCompare = new List<string>();

                            foreach (var child in insertion.Descendants())
                            {
                                if (child.GetType() == typeof(FieldCode))
                                    continue;

                                if (child.GetType() == typeof(FieldChar) &&
                                    (child as FieldChar).FieldCharType == FieldCharValues.Begin)
                                {
                                    fcStarted = true;
                                    continue;
                                }

                                if (child.GetType() == typeof(FieldChar) &&
                                    (child as FieldChar).FieldCharType == FieldCharValues.End)
                                {
                                    fcStarted = false;
                                    continue;
                                }

                                if (child.GetType() == typeof(Text) && fcStarted == false)
                                {
                                    textToCompare.Add(child.InnerText);
                                }
                            }

                            if (updateRevisions.Any(x =>
                                x.Text.Trim() == insertion.InnerText.Trim() &&
                                x.RevisionType == WmlComparer.WmlComparerRevisionType.Inserted))
                            {
                                updateRevisions.Remove(updateRevisions.Where(x => x.Text.Trim() == insertion.InnerText.Trim() && x.RevisionType == WmlComparer.WmlComparerRevisionType.Inserted).First());
                                continue;
                            }

                            //var parentPara = insertion.Parent;
                            //insertion.Remove();
                            //if (parentPara.InnerText == string.Empty)
                            //    parentPara.Remove();

                            if (manualRevisions.Any(x =>
                                x.Text.Trim() == string.Join("", textToCompare) && x.Date == insertion.Date &&
                                x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted))
                            {
                                combinedRuns = new List<InsertedRun>();
                                combinedRevisions = new List<WmlComparer.WmlComparerRevision>();

                                var parentPara = insertion.Parent;
                                insertion.Remove();

                                if (parentPara != null && parentPara.Parent != null && parentPara.InnerText == string.Empty)
                                {
                                    var parentCell = parentPara.Parent;
                                    parentPara.Remove();

                                    if (parentCell != null && parentCell.InnerText == string.Empty)
                                    {
                                        var parentRow = parentCell.Parent;
                                        if (parentRow != null && parentRow.InnerText == string.Empty)
                                            parentRow.Remove();
                                    }
                                }

                                //manualRevisions.Remove(manualRevisions.Where(x => x.Text.Trim() == string.Join("", textToCompare) && x.Date == insertion.Date && x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted).First());
                            }
                            else if (manualRevisions.Any(x => x.Text.Trim().Contains(string.Join("", textToCompare)) && x.Date == insertion.Date && x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted))
                            {
                                combinedRuns.Add(insertion);
                                combinedRevisions.Add(manualRevisions.Where(x => x.Text.Trim().Contains(string.Join("", textToCompare)) && x.Date == insertion.Date && x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted).First());
                                var combinedText = "";
                                foreach (DocumentFormat.OpenXml.Wordprocessing.InsertedRun combinedInsertion in combinedRuns)
                                {
                                    combinedText += combinedInsertion.InnerText;
                                }

                                if (manualRevisions.Any(x =>
                                    x.Text.Trim() == combinedText && x.Date == insertion.Date &&
                                    x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted))
                                {
                                    foreach (DocumentFormat.OpenXml.Wordprocessing.InsertedRun combinedInsertion in combinedRuns)
                                    {
                                        var parentPara = combinedInsertion.Parent;
                                        combinedInsertion.Remove();
                                        if (parentPara.InnerText == string.Empty)
                                            parentPara.Remove();
                                    }

                                    //foreach (var rev in combinedRevisions)
                                    //    manualRevisions.Remove(rev);

                                    combinedRuns = new List<InsertedRun>();
                                }

                                //manualRevisions.Remove(manualRevisions.Where(x => x.Text.Trim().Contains(string.Join("", textToCompare)) && x.Date == insertion.Date && x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted).First());
                            }
                            else
                            {
                                foreach (WmlComparerRevision mr in manualRevisions.Where(x => x.Date == insertion.Date && x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted))
                                {
                                    var mrText = mr.Text;
                                    var containsText = true;
                                    foreach (var text in textToCompare)
                                    {
                                        if (mrText.Contains(text))
                                        {
                                            mrText = mrText.Remove(mrText.IndexOf(text), text.Length);
                                        }
                                        else
                                        {
                                            containsText = false;
                                            break;
                                        }
                                    }

                                    if (containsText)
                                    {
                                        var parentPara = insertion.Parent;
                                        insertion.Remove();

                                        if (parentPara != null && parentPara.Parent != null && parentPara.InnerText == string.Empty)
                                        {
                                            var parentCell = parentPara.Parent;
                                            parentPara.Remove();

                                            if (parentCell != null && parentCell.InnerText == string.Empty)
                                            {
                                                var parentRow = parentCell.Parent;
                                                if (parentRow != null && parentRow.InnerText == string.Empty)
                                                    parentRow.Remove();
                                            }
                                        }

                                        break;
                                    }
                                }
                            }
                        }

                        var insertedMathControls = wdDoc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.InsertedMathControl>().ToList();
                        foreach (var insertion in insertedMathControls)
                        {
                            if (manualRevisions.Any(x =>
                                x.Text.Trim() == insertion.InnerText.Trim() && x.Date == insertion.Date &&
                                x.RevisionType == WmlComparer.WmlComparerRevisionType.Deleted))
                            {
                                var parentPara = insertion.Parent;
                                insertion.Remove();
                                if (parentPara.InnerText == string.Empty)
                                    parentPara.Remove();
                            }
                        }

                        foreach (FieldChar fc in wdDoc.MainDocumentPart.Document.Body.Descendants<FieldChar>())
                        {
                            fc.Dirty = true;
                        }

                        foreach (var bkmStart in wdDoc.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
                        {
                            var bkmId = bkmStart.Id.Value;
                            var bkmEnd = wdDoc.MainDocumentPart.Document.Body.Descendants<BookmarkEnd>()
                                .Where(x => x.Id.Value == bkmId).FirstOrDefault();
                        }
                    }

                    returnDoc = new WmlDocument("Dummy.docx", ms.ToArray());
                }
            }
            catch (Exception ex)
            {
                var msg = ex.Message;
            }

            return returnDoc;

            //setBookmarksAfterCompare(path);
        }

        //private static WmlDocument acceptManualEditsAfterTriangularCompare2(WmlDocument result, List<WmlComparerRevision> manualRevisions, List<WmlComparerRevision> updateRevisions)
        //{
        //    WmlDocument returnDoc = result;
        //    try
        //    {
        //        using (var ms = new MemoryStream())
        //        {
        //            ms.Write(result.DocumentByteArray, 0, result.DocumentByteArray.Length);
        //            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(ms, true))
        //            {
        //                //Handle Deletions
        //                var deletedItems = wdDoc.MainDocumentPart.Document.Body.Descendants<Deleted>().ToList();
        //                foreach (Deleted deletedItem in deletedItems)
        //                {
        //                    if (updateRevisions.Any(x =>
        //                        x.Text.Trim() == deletedItem.InnerText.Trim() && x.Date == deletedItem.Date &&
        //                        x.RevisionType == WmlComparerRevisionType.Inserted))
        //                    {

        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        var msg = ex.Message;
        //    }

        //    return returnDoc;
        //}

        private static WmlDocument CompareInternal(
            WmlDocument source1,
            WmlDocument source2,
            WmlComparerSettings settings,
            bool preProcessMarkupInOriginal)
        {
            if (preProcessMarkupInOriginal)
            {
                source1 = PreProcessMarkup(source1, settings.StartingIdForFootnotesEndnotes + 1000);
            }

            source2 = PreProcessMarkup(source2, settings.StartingIdForFootnotesEndnotes + 2000);

            SaveDocumentIfDesired(source1, "Source1-Step1-PreProcess.docx", settings);
            SaveDocumentIfDesired(source2, "Source2-Step1-PreProcess.docx", settings);

            //source1 = RemoveBookmarks(source1);
            //source2 = RemoveBookmarks(source2);

            // at this point, both source1 and source2 have unid on every element.  These are the values that will
            // enable reassembly of the XML tree.  But we need other values.

            // In source1:
            // - accept tracked revisions
            // - determine hash code for every block-level element
            // - save as attribute on every element

            // - accept tracked revisions and reject tracked revisions leave the unids alone, where possible.
            // - after accepting and calculating the hash, then can use the unids to find the right block-level
            //   element in the unmodified source1, and install the hash

            // In source2:
            // - reject tracked revisions
            // - determine hash code for every block-level element
            // - save as an attribute on every element

            // - after rejecting and calculating the hash, then can use the unids to find the right block-level element
            //   in the unmodified source2, and install the hash

            // - sometimes after accepting or rejecting tracked revisions, several paragraphs will get coalesced into a
            //   single paragraph due to paragraph marks being inserted / deleted.
            // - in this case, some paragraphs will not get a hash injected onto them.
            // - if a paragraph doesn't have a hash, then it will never correspond to another paragraph, and such
            //   issues will need to be resolved in the normal execution of the LCS algorithm.
            // - note that when we do propagate the unid through for the first paragraph.

            // Establish correlation between the two.
            // Find the longest common sequence of block-level elements where hash codes are the same.
            // this sometimes will be every block level element in the document.  Or sometimes will be just a fair
            // number of them.

            // at the start of doing the LCS algorithm, we will match up content, and put them in corresponding unknown
            // correlated comparison units.  Those paragraphs will only ever be matched to their corresponding paragraph.
            // then the algorithm can proceed as usual.

            // need to call ChangeFootnoteEndnoteReferencesToUniqueRange before creating the wmlResult document, so that
            // the same GUID ids are used for footnote and endnote references in both the 'after' document, and in the
            // result document.

            WmlDocument source1AfterAccepting = RevisionProcessor.AcceptRevisions(source1);
            WmlDocument source2AfterRejecting = RevisionProcessor.RejectRevisions(source2);

            SaveDocumentIfDesired(source1AfterAccepting, "Source1-Step2-AfterAccepting.docx", settings);
            SaveDocumentIfDesired(source2AfterRejecting, "Source2-Step2-AfterRejecting.docx", settings);

            // this creates the correlated hash codes that enable us to match up ranges of paragraphs based on
            // accepting in source1, rejecting in source2
            source1 = HashBlockLevelContent(source1, source1AfterAccepting, settings);
            source2 = HashBlockLevelContent(source2, source2AfterRejecting, settings);

            SaveDocumentIfDesired(source1, "Source1-Step3-AfterHashing.docx", settings);
            SaveDocumentIfDesired(source2, "Source2-Step3-AfterHashing.docx", settings);

            // Accept revisions in before, and after
            source1 = RevisionProcessor.AcceptRevisions(source1);
            source2 = RevisionProcessor.AcceptRevisions(source2);

            SaveDocumentIfDesired(source1, "Source1-Step4-AfterAccepting.docx", settings);
            SaveDocumentIfDesired(source2, "Source2-Step4-AfterAccepting.docx", settings);

            // after accepting revisions, some unids may have been removed by revision accepter, along with the
            // correlatedSHA1Hash codes, this is as it should be.
            // but need to go back in and add guids to paragraphs that have had them removed.

            using (var ms = new MemoryStream())
            {
                ms.Write(source2.DocumentByteArray, 0, source2.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    AddUnidsToMarkupInContentParts(wDoc);
                }
            }

            var wmlResult = new WmlDocument(source1);
            using (var ms1 = new MemoryStream())
            using (var ms2 = new MemoryStream())
            {
                ms1.Write(source1.DocumentByteArray, 0, source1.DocumentByteArray.Length);
                ms2.Write(source2.DocumentByteArray, 0, source2.DocumentByteArray.Length);
                WmlDocument producedDocument;

                using (WordprocessingDocument wDoc1 = WordprocessingDocument.Open(ms1, true))
                using (WordprocessingDocument wDoc2 = WordprocessingDocument.Open(ms2, true))
                {
                    producedDocument = ProduceDocumentWithTrackedRevisions(settings, wmlResult, wDoc1, wDoc2);
                }

                SaveDocumentsAfterProducingDocument(ms1, ms2, settings);
                SaveCleanedDocuments(source1, producedDocument, settings);

                return producedDocument;
            }
        }

        private static List<BkmInfo> bkmInfos = new List<BkmInfo>();
        private static WmlDocument RemoveBookmarks(WmlDocument document)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(document.DocumentByteArray,0,document.DocumentByteArray.Length);
                using (DocumentFormat.OpenXml.Packaging.WordprocessingDocument doc1 =
                    DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, true))
                {
                    List<OpenXmlElement> bkmToremove = new List<OpenXmlElement>();
                    string id = "";

                    foreach (var bkmStart in doc1.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
                    {
                        id = bkmStart.Id.Value;
                        var bkmEnd = doc1.MainDocumentPart.Document.Body.Descendants<BookmarkEnd>().Where(x => x.Id.Value == id)
                            .FirstOrDefault();
                        if (bkmEnd == null)
                            bkmToremove.Add(bkmStart);
                    }

                    foreach (var bkmEnd in doc1.MainDocumentPart.Document.Body.Descendants<BookmarkEnd>())
                    {
                        id = bkmEnd.Id.Value;
                        var bkmStart = doc1.MainDocumentPart.Document.Body.Descendants<BookmarkStart>()
                            .Where(x => x.Id.Value == id).FirstOrDefault();
                        if (bkmStart == null)
                            bkmToremove.Add(bkmEnd);
                    }

                    foreach (var bkm in bkmToremove)
                        bkm.Remove();

                    foreach (var bkm in doc1.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
                    {
                        var nextSibling = bkm.NextSibling();
                        while (nextSibling != null)
                        {
                            if (nextSibling.GetType() != typeof(BookmarkStart) && nextSibling.GetType() != typeof(BookmarkEnd))
                            {
                                BkmInfo info = new BkmInfo();
                                info.elementXML = bkm.OuterXml;
                                info.UnidofNextElement = nextSibling.GetAttribute("Unid", "pt14").Value;
                                bkmInfos.Add(info);
                                break;
                            }
                        }
                    }

                    var bkmRemovedDocument = new WmlDocument("bkmRemoved.docx", ms.ToArray());
                    return bkmRemovedDocument;
                }
            }
        }

        private static void SaveDocumentIfDesired(WmlDocument source, string name, WmlComparerSettings settings)
        {
            if (SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                var fileInfo = new FileInfo(Path.Combine(settings.DebugTempFileDi.FullName, name));
                source.SaveAs(fileInfo.FullName);
            }
        }

        private static void SaveDocumentsAfterProducingDocument(MemoryStream ms1, MemoryStream ms2, WmlComparerSettings settings)
        {
            if (SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                SaveDocumentIfDesired(new WmlDocument("after1.docx", ms1), "Source1-Step5-AfterProducingDocument.docx", settings);
                SaveDocumentIfDesired(new WmlDocument("after2.docx", ms2), "Source2-Step5-AfterProducingDocument.docx", settings);
            }
        }

        private static void SaveCleanedDocuments(WmlDocument source1, WmlDocument producedDocument, WmlComparerSettings settings)
        {
            if (SaveIntermediateFilesForDebugging && settings.DebugTempFileDi != null)
            {
                WmlDocument cleanedSource = CleanPowerToolsAndRsid(source1);
                SaveDocumentIfDesired(cleanedSource, "Cleaned-Source.docx", settings);

                WmlDocument cleanedProduced = CleanPowerToolsAndRsid(producedDocument);
                SaveDocumentIfDesired(cleanedProduced, "Cleaned-Produced.docx", settings);
            }
        }

        private static WmlDocument CleanPowerToolsAndRsid(WmlDocument producedDocument)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(producedDocument.DocumentByteArray, 0, producedDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    foreach (OpenXmlPart cp in wDoc.ContentParts())
                    {
                        XDocument xd = cp.GetXDocument();
                        object newRoot = CleanPartTransform(xd.Root);
                        xd.Root?.ReplaceWith(newRoot);
                        cp.PutXDocument();
                    }
                }

                var cleaned = new WmlDocument("cleaned.docx", ms.ToArray());
                return cleaned;
            }
        }

        private static object CleanPartTransform(XNode node)
        {
            if (node is XElement element)
            {
                return new XElement(element.Name,
                    element.Attributes().Where(a => a.Name.Namespace != PtOpenXml.pt &&
                                                    !a.Name.LocalName.ToLower().Contains("rsid")),
                    element.Nodes().Select(CleanPartTransform));
            }

            return node;
        }
    }

    public class BkmInfo
    {
        public string elementXML;
        public string UnidofNextElement;
    }
}
