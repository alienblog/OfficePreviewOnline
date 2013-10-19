using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Brume.Common.Office;
using MongoDB.Bson;
using MongoDB.Repository.GridFs;

namespace Brume.PreviewOnline.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            var docs = Models.DocModel.FindAll();
            return View(docs);
        }

        public ActionResult Add()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Add(Models.DocModel model)
        {
            try
            {
                var doc = Request.Files[0];
                model.Doc = new File {FileName = doc.FileName, Data = new byte[doc.InputStream.Length]};

                doc.InputStream.Read(model.Doc.Data, 0, model.Doc.Data.Length);

                model.Doc.Save();
                model.Save();

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message;
            }
            return View(model);
        }

        public ActionResult Delete(string id)
        {
            var doc = Models.DocModel.Find(id);
            return View(doc);
        }

        [HttpPost]
        [ActionName("Delete")]
        public ActionResult ConfirmDelete(string id)
        {
            var doc = Models.DocModel.Find(id);
            doc.Doc.Remove();
            doc.Remove();

            return RedirectToAction("Index");
        }

        public ActionResult Preview(string id)
        {
            var convertor = new OfficeConvertor();
            var url = convertor.GetTempDocumentUrl(id);

            ViewBag.Url = url;

            return View();
        }
    }
}
