using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Brume.Common.Office
{
    public interface IFileHelper
    {
        string GetDocumentPath(string id);

        void ClearTempDocument(string dir);
    }
}
