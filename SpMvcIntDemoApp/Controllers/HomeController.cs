using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Microsoft.SharePoint.Client;
using SpMvcIntDemoApp.Models;

namespace SpMvcIntDemoApp.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult Documents()
        {
            ClientContext ctx = new ClientContext("{repalce}");
            List documents = ctx.Web.Lists.GetByTitle("Documents");
            CamlQuery query = new CamlQuery() { ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/>" + "<Value Type='Number'>10</Value></Geq></Where></Query><RowLimit>100</RowLimit></View>" };

            ListItemCollection items = documents.GetItems(query);
            ctx.Load(items);
            ctx.ExecuteQuery();
            List<DocumentModel> docs = new List<DocumentModel>();

            foreach (var item in items)
                docs.Add(new DocumentModel() { Id = item.Id, Title = (item["Title"] != null ? item["Title"].ToString() : "") });

            return View(docs);
        }
        public JsonpResult DocumentsJsonP()
        {
            ClientContext ctx = new ClientContext("{replace}");
            List documents = ctx.Web.Lists.GetByTitle("Documents");
            CamlQuery query = new CamlQuery() { ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/>" + "<Value Type='Number'>10</Value></Geq></Where></Query><RowLimit>100</RowLimit></View>" };

            ListItemCollection items = documents.GetItems(query);
            ctx.Load(items);
            ctx.ExecuteQuery();
            List<DocumentModel> docs = new List<DocumentModel>();

            foreach (var item in items)
                docs.Add(new DocumentModel() { Id = item.Id, Title = (item["Title"] != null ? item["Title"].ToString() : "") });

            return new JsonpResult
            {
                Data = new
                {
                    success = true,
                    message = "Documents retrieved successfully",
                    view = this.RenderPartialView("Documents", docs)
                },
                JsonRequestBehavior = JsonRequestBehavior.AllowGet
            };
        }
    }
}