using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using File = MongoDB.Repository.GridFs.File;

namespace Brume.Common.Office
{
    public class MongoDbFileHelper : IFileHelper
    {
        public string GetDocumentPath(string id)
        {
            var file = new File { Id = id };

            if (file.Data.Length == 0)
                throw new FileNotFoundException(id);
            var dir = HttpContext.Current.Request.MapPath("~/temp/");

            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            ClearTempDocument(dir);

            var filename = string.Format("{0}{1}", id, GetExtention(file.FileName));

            using (var fs = new FileStream(dir + filename, FileMode.Create, FileAccess.ReadWrite))
            {
                fs.Write(file.Data, 0, file.Data.Length);
                fs.Flush();
            }
            return dir + filename;
        }

        public void ClearTempDocument(string dir)
        {
            var dirInfo = new DirectoryInfo(dir);
            FileInfo[] fileInfos = dirInfo.GetFiles();
            foreach (
                FileInfo info in
                    from info in fileInfos
                    let exTime = DateTime.Now.AddSeconds(-30)
                    where info.CreationTime < exTime
                    select info)
            {
                try
                {
                    System.IO.File.Delete(info.FullName);
                }
                catch
                {
                }
            }
        }

        private static string GetExtention(string filename)
        {
            int s = filename.LastIndexOf('.');
            int l = filename.Length - s;
            return filename.Substring(s, l);
        }
    }
}
