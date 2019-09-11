using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.Net;
using System.Drawing;

namespace ConvertionOfPPTWeb.Business
{
    public class PresentationService
    {
        public Presentation LoadPresentation(string url)
        {
            Presentation ppt = new Presentation();

            ppt.LoadFromFile(url);

            //ppt.LoadFromStream(GetPresentationStream(url), FileFormat.Pptx2010);

            return ppt;
            
        }

        private Stream GetPresentationStream(string url)
        {
            WebRequest request = WebRequest.Create(url);

            //request.Credentials = CredentialCache.DefaultCredentials;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            Stream dataStream = response.GetResponseStream();

            dataStream.Close();

            response.Close();

            return dataStream;
        }

        public string SplitPresentation(Presentation ppt, int slideIndex, string saveDir)
        {
            if (ppt == null)
                return null;

            if (slideIndex >= ppt.Slides.Count)
            {
                return null;
            }

            Presentation newPPT = new Presentation();

            newPPT.Slides.RemoveAt(0);

            newPPT.Slides.Append(ppt.Slides[slideIndex]);

            string savePath = Path.Combine(saveDir, ppt.DocumentProperty.Title + "_" + slideIndex + ".pptx");

            newPPT.SaveToFile(savePath, FileFormat.Pptx2010);

            return savePath;
        }

        public string ConvertPPTToHTML(Presentation ppt, string saveDir)
        {
            if (ppt == null)
                return null;

            string savePath = Path.Combine(saveDir, ppt.DocumentProperty.Title + "a.html");

            ppt.SaveToFile(savePath, FileFormat.Html);

            return readFromFile(savePath);
        }

        private string readFromFile(string path)
        {
            string dataInFile = "";
            string line;
            using (StreamReader reader = new StreamReader(path))
            {
                while ((line = reader.ReadLine()) != null)
                {
                    dataInFile += line;
                }
            }

            return dataInFile;
        }

        public List<string> GetShapes(Presentation ppt, int slideIndex, string saveDir)
        {
            if (ppt == null)
                return null;

            if (slideIndex >= ppt.Slides.Count)
            {
                return null;
            }

            var allShapes = ppt.Slides[slideIndex].Shapes;

            List<string> list = new List<string>();

            for(var i = 0;i < allShapes.Count;i ++)
            {
                Image image = allShapes.SaveAsImage(i);

                string savePath = Path.Combine(saveDir, "pic_"+i+".png");

                list.Add("pic_" + i + ".png");

                image.Save(savePath, System.Drawing.Imaging.ImageFormat.Png);
            }
            return list;
        }
    }
}