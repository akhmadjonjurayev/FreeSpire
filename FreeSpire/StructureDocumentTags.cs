using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace FreeSpire
{
    public class StructureDocumentTags
    {
        public List<StructureDocumentTagInline> TagInlines = new List<StructureDocumentTagInline>();
        public List<StructureDocumentTag> Tags = new List<StructureDocumentTag>();

        public void LoadAllTags(Document document)
        {
            foreach (Section section in document.Sections)
            {
                foreach (DocumentObject obj in section.Body.ChildObjects)
                {
                    GetSDTInDocumentObject(obj);
                }
            }
        }

        public void GetSDTInDocumentObject(DocumentObject documentObject)
        {
            if (documentObject.DocumentObjectType == DocumentObjectType.TextBox)
            {
                foreach (DocumentObject pobj in (documentObject as TextBox).ChildObjects)
                {
                    GetSDTInDocumentObject(pobj);
                }
            }

            if (documentObject.DocumentObjectType == DocumentObjectType.Paragraph)
            {
                foreach (DocumentObject pobj in (documentObject as Paragraph).ChildObjects)
                {
                    if (pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline)
                    {
                        TagInlines.Add(pobj as StructureDocumentTagInline);
                    }
                    if (pobj.DocumentObjectType == DocumentObjectType.ShapeGroup)
                    {
                        foreach (DocumentObject pobjInner in pobj.ChildObjects)
                        {
                            GetSDTInDocumentObject(pobjInner);
                        }
                    }
                    if (pobj.DocumentObjectType == DocumentObjectType.TextBox)
                    {
                        foreach (DocumentObject pobjInner in pobj.ChildObjects)
                        {
                            GetSDTInDocumentObject(pobjInner);
                        }
                    }
                }
            }
            else if (documentObject is StructureDocumentTag)
            {
                Tags.Add(documentObject as StructureDocumentTag);
            }

            else if (documentObject.DocumentObjectType == DocumentObjectType.Table)
            {
                foreach (Spire.Doc.TableRow row in (documentObject as Spire.Doc.Table).Rows)
                {
                    foreach (Spire.Doc.TableCell cell in row.Cells)
                    {
                        foreach (DocumentObject cellChild in cell.ChildObjects)
                        {
                            if (cellChild.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
                            {
                                Tags.Add(cellChild as StructureDocumentTag);
                            }
                            else if (cellChild.DocumentObjectType == DocumentObjectType.Paragraph)
                            {
                                foreach (DocumentObject pobj in (cellChild as Paragraph).ChildObjects)
                                {
                                    if (pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline)
                                    {
                                        TagInlines.Add(pobj as StructureDocumentTagInline);
                                    }
                                }
                            }
                        }
                    }
                }

            }
        }

        public string GetSDTText(string tagName)
        {
            string result = "";
            try
            {
                if (string.IsNullOrEmpty(result))
                {
                    StructureDocumentTagInline inlineTag = GetTagInlines(tagName).FirstOrDefault();
                    if (inlineTag != null)
                    {
                        result = inlineTag.SDTContent.Text;
                    }
                }

                if (string.IsNullOrEmpty(result))
                {
                    StructureDocumentTag tag = GetTag(tagName);
                    if (tag != null)
                    {
                        List<string> lines = new List<string>();
                        foreach (Paragraph paragraph in tag.SDTContent.Paragraphs)
                        {
                            lines.Add(paragraph.Text);
                        }
                        result = String.Join(Environment.NewLine, lines.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return result;
        }

        public List<TextRange> GetTagInlineTextRanges(string tagName)
        {
            List<TextRange> resultTextRanges = new List<TextRange>();
            try
            {
                List<StructureDocumentTagInline> tagInlines = GetTagInlines(tagName);
                if (tagInlines != null && tagInlines.Count > 0)
                {
                    foreach (StructureDocumentTagInline tagInline in tagInlines)
                    {
                        if (tagInline.ChildObjects.Count > 0 && (tagInline.ChildObjects[0] as TextRange) != null)
                        {
                            TextRange orginalTextRange = tagInline.ChildObjects[0] as TextRange;
                            if (orginalTextRange != null)
                            {
                                orginalTextRange = orginalTextRange.Clone() as TextRange;
                            }
                            tagInline.ChildObjects.Clear();
                            tagInline.ChildObjects.Add(orginalTextRange);
                            resultTextRanges.Add(orginalTextRange);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return resultTextRanges;
        }

        public List<DocPicture> GetTagInlineDocPictures(string tagName)
        {
            List<DocPicture> resultDocPictures = new List<DocPicture>();
            try
            {
                List<StructureDocumentTagInline> tagInlines = GetTagInlines(tagName);
                if (tagInlines != null && tagInlines.Count > 0)
                {
                    foreach (StructureDocumentTagInline tagInline in tagInlines)
                    {
                        if (tagInline != null && tagInline.ChildObjects.Count > 0 && (tagInline.ChildObjects[0] as DocPicture != null))
                        {
                            resultDocPictures.Add(tagInline.ChildObjects[0] as DocPicture);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return resultDocPictures;
        }

        public List<StructureDocumentTagInline> GetTagInlines(string tagName)
        {
            return TagInlines.Where(s => s.SDTProperties.Tag.ToUpper() == tagName.ToUpper()).ToList();
        }

        public StructureDocumentTag GetTag(string tagName)
        {
            return Tags.FirstOrDefault(s => s.SDTProperties.Tag.ToUpper() == tagName.ToUpper());
        }

        public void SetTagValue(string tagName, string value)
        {
            var targetTag = Tags.FirstOrDefault(tag => tag.SDTProperties.Tag.ToLower() == tagName.ToLower());
            if(targetTag != null)
            {
                (targetTag.FirstChild as Paragraph).Text = value;
            }
        }

        public void SetInlineTagValue(string inlineTagName, string value)
        {
            var inlineTag = TagInlines.FirstOrDefault(l=>l.SDTProperties.Tag.ToLower() == inlineTagName.ToLower());
            if(inlineTag != null)
            {
                (inlineTag.FirstChild as TextRange).Text = value;
            }
        }
    }
}
