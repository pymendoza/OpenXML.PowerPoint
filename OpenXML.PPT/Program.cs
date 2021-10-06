using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace OpenXML.PPT
{
    class Program
    {
        static void Main(string[] args)
        {
            // The absoluete path to the PowerPoint file
            var path = @"E:\Downloads\Upwork part I.pptx";

            // The text that will replace the {{name}} placeholder
            var replaceName = "Leader X";

            var dir = Path.GetDirectoryName(path);
            var filename = Path.GetFileNameWithoutExtension(path);
            var ext = Path.GetExtension(path);

            // The path to the new generated file with "-generated" appended to the filename of the file in the same directory
            var newPath = Path.Combine(dir, $"{filename}-generated{ext}");

            File.Copy(path, newPath);

            using (PresentationDocument wd = PresentationDocument.Open(newPath, true))
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

                                wd.Save();

                                break;
                            }
                        }
                    }
                }
            }
        }
    }
}