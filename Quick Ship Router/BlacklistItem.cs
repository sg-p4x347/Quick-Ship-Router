using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quick_Ship_Router
{
    enum Blacklist
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
            m_method = Blacklist.IsEqualTo;
            m_itemCode = itemCode;
        }
        public BlacklistItem(Blacklist method, string itemCode)
        {
            m_method = method;
            m_itemCode = itemCode;
        }
        
        public static bool operator ==(string item,BlacklistItem blacklist)
        {
            if (blacklist.m_method == Blacklist.IsEqualTo)
            {
                return item == blacklist.m_itemCode;
            }
            else if (blacklist.m_method == Blacklist.StartsWith)
            {
                return item.Substring(0, blacklist.m_itemCode.Length) == blacklist.m_itemCode;
            }
            else if (blacklist.m_method == Blacklist.EndsWith)
            {
                return item.Substring((item.Length - 1) - blacklist.m_itemCode.Length, item.Length-1) == blacklist.m_itemCode;
            } else if (blacklist.m_method == Blacklist.Has)
            {
                return item.IndexOf(blacklist.m_itemCode) != -1;
            } else
            {
                return false;
            }
        }
        public static bool operator !=(string item, BlacklistItem blacklist)
        {
            if (blacklist.m_method == Blacklist.IsEqualTo)
            {
                return !(item == blacklist.m_itemCode);
            }
            else if (blacklist.m_method == Blacklist.StartsWith)
            {
                return !(item.Substring(0, blacklist.m_itemCode.Length) == blacklist.m_itemCode);
            }
            else if (blacklist.m_method == Blacklist.EndsWith)
            {
                return !(item.Substring((item.Length - 1) - blacklist.m_itemCode.Length, item.Length - 1) == blacklist.m_itemCode);
            }
            else if (blacklist.m_method == Blacklist.Has)
            {
                return !(item.IndexOf(blacklist.m_itemCode) != -1);
            } else
            {
                return false;
            }
        }
        private Blacklist m_method = Blacklist.IsEqualTo;
        private string m_itemCode = "";

    }
}
