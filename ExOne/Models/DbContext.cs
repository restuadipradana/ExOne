using LiteDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExOne.Models
{
    public static class DbContext
    {
        static LiteDatabase _instance;
        public static void Init(string fileName)
        {
            _instance = new LiteDatabase(fileName);
        }
        public static LiteDatabase GetInstance()
        {
            return _instance;
        }
    }
}
