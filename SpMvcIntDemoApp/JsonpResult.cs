using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace SpMvcIntDemoApp
{
    public class JsonpResult : JsonResult
    {
        public Encoding ContentEncoding
        {
            get;
            set;
        }

        public string ContentType
        {
            get;
            set;
        }

        public object Data
        {
            get;
            set;
        }

        public string JsonCallback { get; set; }

        public override void ExecuteResult(ControllerContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            this.JsonCallback = context.HttpContext.Request["jsoncallback"];

            if (string.IsNullOrEmpty(this.JsonCallback))
                this.JsonCallback = context.HttpContext.Request["callback"];

            if (string.IsNullOrEmpty(this.JsonCallback))
                throw new ArgumentNullException("JsonCallback required for JSONP response.");

            HttpResponseBase response = context.HttpContext.Response;

            if (!String.IsNullOrEmpty(ContentType))
            {
                response.ContentType = ContentType;
            }
            else
            {
                response.ContentType = "application/json";
            }
            if (ContentEncoding != null)
            {
                response.ContentEncoding = ContentEncoding;
            }
            if (Data != null)
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                response.Write(string.Format("{0}({1});", this.JsonCallback, serializer.Serialize(Data)));
            }
        }
    }
}
