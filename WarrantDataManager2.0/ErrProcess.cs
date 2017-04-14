using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WarrantDataManager2._0
{
    class ErrProcess
    {

        
    }

    

    public class ErrObj
    {
        public DateTime errTime;
        public int errCode = 0;
        public string errMessage = "";

        public ErrObj(int errCode, string errMessage)
        {
            this.errTime = DateTime.Now;
            this.errCode = errCode;
            this.errMessage = errMessage;
        }
    }
}
