using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using MongoDB.Repository;
using MongoDB.Repository.GridFs;

namespace Brume.PreviewOnline.Models
{
    public class DocModel:Entity<DocModel>
    {
        public string Name { get; set; }

        public File Doc { get; set; }
    }
}