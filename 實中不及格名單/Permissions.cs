using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JH.IBSH.FailPeriodicalExamList
{
    class Permissions
    {
        public static string 實中班級不及格名單 { get { return "JH.IBSH.FailPeriodicalExamList.Class.cs"; } }
        public static bool 實中班級不及格名單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[實中班級不及格名單].Executable;
            }
        }
        public static string 實中全部不及格名單 { get { return "JH.IBSH.FailPeriodicalExamList.All.cs"; } }
        public static bool 實中全部不及格名單權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[實中全部不及格名單].Executable;
            }
        }
    }
}
