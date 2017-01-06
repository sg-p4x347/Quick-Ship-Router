using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quick_Ship_Router
{
    enum Method
    {
        StartsWith,
        IsEqualTo,
        EndsWith,
        Has
    }
    class BlacklistItem
    {
        public BlacklistItem(string itemCode)
        {
            m_method = Method.IsEqualTo;
            m_itemCode = itemCode;
        }
        public BlacklistItem(Method method, string itemCode)
        {
            m_method = method;
            m_itemCode = itemCode;
        }
        public bool StartsWith(string s)
        {
            return s.Substring(0, m_itemCode.Length) == m_itemCode;
        }
        
        public static bool operator ==(string item,BlacklistItem blacklist)
        {
            if (blacklist.m_method == Method.IsEqualTo)
            {
                return item == blacklist.m_itemCode;
            }
            else if (blacklist.m_method == Method.StartsWith)
            {
                return item.Substring(0, blacklist.m_itemCode.Length) == blacklist.m_itemCode;
            }
            else if (blacklist.m_method == Method.EndsWith)
            {
                return item.Substring((item.Length - 1) - blacklist.m_itemCode.Length, item.Length-1) == blacklist.m_itemCode;
            } else if (blacklist.m_method == Method.Has)
            {
                return item.IndexOf(blacklist.m_itemCode) != -1;
            } else
            {
                return false;
            }
        }
        public static bool operator !=(string item, BlacklistItem blacklist)
        {
            if (blacklist.m_method == Method.IsEqualTo)
            {
                return !(item == blacklist.m_itemCode);
            }
            else if (blacklist.m_method == Method.StartsWith)
            {
                return !(item.Substring(0, blacklist.m_itemCode.Length) == blacklist.m_itemCode);
            }
            else if (blacklist.m_method == Method.EndsWith)
            {
                return !(item.Substring((item.Length - 1) - blacklist.m_itemCode.Length, item.Length - 1) == blacklist.m_itemCode);
            }
            else if (blacklist.m_method == Method.Has)
            {
                return !(item.IndexOf(blacklist.m_itemCode) != -1);
            } else
            {
                return false;
            }
        }
        private Method m_method = Method.IsEqualTo;
        private string m_itemCode = "";

    }
}
