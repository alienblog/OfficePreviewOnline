using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using MongoDB.Repository;

namespace Brume.PreviewOnline
{
    public class DocDbContext:MongoDBContext
    {
        public DocDbContext() : base("docContext")
        {

        }

        public override void OnRegisterModel(ITypeRegistration registration)
        {
            registration.RegisterType<Models.DocModel>();
        }
    }
}