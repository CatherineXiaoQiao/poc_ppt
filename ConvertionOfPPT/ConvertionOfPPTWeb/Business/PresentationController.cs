using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Web.Http.Results;

namespace ConvertionOfPPTWeb.Business
{
    public class PresentationController : ApiController
    {
        
        public string GetConvertSlideToHtml()
        {//string pptUrl, int slideIndex
         string pptUrl = HttpContext.Current.Server.MapPath("/SGEN-ATLAS_Simple_ForTest.pptx");
            //string pptUrl = "https://canvizconsultingllc.sharepoint.com/:p:/r/sites/SeattleGen/_layouts/15/Doc.aspx?sourcedoc=%7B4AC5C6A1-F40C-491E-85D8-CD9EB4684E53%7D&file=SGEN-ATLAS_Simple_ForTest.pptx&action=edit&mobileredirect=true";
            int slideIndex = 1;
            
            PresentationService presentationService = new PresentationService();

            var ppt = presentationService.LoadPresentation(pptUrl);

            string splitedPPTPath = presentationService.SplitPresentation(ppt, slideIndex, HttpContext.Current.Server.MapPath("/cf"));

            var splitedPPT = presentationService.LoadPresentation(splitedPPTPath);

            string convertedPPTPath = presentationService.ConvertPPTToHTML(splitedPPT, HttpContext.Current.Server.MapPath("/sf"));

            return "The html file:"+convertedPPTPath;
        }
        // GET api/<controller>/5
        public List<string> GetShapes(int id)
        {
            string pptUrl = HttpContext.Current.Server.MapPath("/SGEN-ATLAS_Simple_ForTest.pptx");
            //string pptUrl = "https://canvizconsultingllc.sharepoint.com/:p:/r/sites/SeattleGen/_layouts/15/Doc.aspx?sourcedoc=%7B4AC5C6A1-F40C-491E-85D8-CD9EB4684E53%7D&file=SGEN-ATLAS_Simple_ForTest.pptx&action=edit&mobileredirect=true";
            int slideIndex = 1;

            PresentationService presentationService = new PresentationService();

            var ppt = presentationService.LoadPresentation(pptUrl);

            return presentationService.GetShapes(ppt, slideIndex, HttpContext.Current.Server.MapPath("/pf"));
        }

        // POST api/<controller>
        public void Post([FromBody]string value)
        {
            
        }

        // PUT api/<controller>/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/<controller>/5
        public void Delete(int id)
        {
        }
    }
}