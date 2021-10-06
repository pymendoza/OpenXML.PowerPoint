using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace OpenXML.PPT
{
    class Program
    {
        static void Main(string[] args)
        {
            var replaceName = "Leader X";

            var path = @"E:\Downloads\Upwork part I.pptx";
            
            using (PresentationDocument wd = PresentationDocument.Open(path, true))
            {
                var presentationPart = wd.PresentationPart.GetPartById("rId2");

                foreach (var part in presentationPart.Parts)
                {
                    if (part.RelationshipId == "rId2")
                    {
                        foreach (var innerPart in part.OpenXmlPart.Parts)
                        {
                            if (part.RelationshipId == "rId2") 
                            {
                                OpenXmlElement replace = null;
                                OpenXmlElement old = null;

                                var slide = ((SlidePart)innerPart.OpenXmlPart).Slide;

                                foreach (var childElemet in slide.ChildElements)
                                {
                                    if (childElemet.LocalName == "cSld")
                                    {
                                        old = childElemet;
                                        replace = childElemet.Clone() as OpenXmlElement;
                                        break;
                                    }
                                }

                                replace.InnerXml = replace.InnerXml.Replace("{{name}}", replaceName);
                                
                                slide.ReplaceChild(replace, old);

                                var dir = Path.GetDirectoryName(path);
                                var filename = Path.GetFileNameWithoutExtension(path);
                                var ext = Path.GetExtension(path);

                                var newFilename = Path.Combine(dir, $"{filename}-copy{ext}");
                                
                                wd.SaveAs(newFilename);
                                break;
                            }
                        }
                    }
                }
            }
        }
    }
}