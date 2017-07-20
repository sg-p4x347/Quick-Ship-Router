using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace EATS
{
    public class BackupManager
    {
        #region Public Methods
        static public void Initialize(string rootDir = null)
        {
            
            //CreateBackupDir();

            string[] backupPaths = System.IO.Directory.GetDirectories(System.IO.Path.Combine(rootDir != null ? rootDir : RootDir, "backup\\"));
            m_backupDates = new List<DateTime>();
            foreach (string path in backupPaths)
            {
                m_backupDates.Add(StringToDate(Path.GetFileName(path)));
            }
            // sort descending 
            m_backupDates.Sort((x, y) => y.CompareTo(x));
        }
       
        // Standardized conversion from dateTime to string
        static public string DateToString(DateTime date)
        {
            return date.ToString("MM-dd-yyyy");
        }
        // Standardized conversion from string to dateTime
        static public DateTime StringToDate(string date)
        {
            return DateTime.Parse(date);
        }
        // gets the most recent past backup
        static public DateTime GetMostRecent()
        {
            return m_backupDates.First(x => x.Date < DateTime.Today.Date);
        }
        // returns true if a current backup for today exists
        static public bool CurrentBackupExists(string file)
        {
            return  (m_backupDates.Exists(x => x == DateTime.Today.Date)
                && File.Exists(Path.Combine(RootDir, "backup", DateToString(DateTime.Today.Date),file)));
        }
        // returns the requested file from current day backup
        static public string Import(string filename,DateTime? d = null)
        {
            DateTime date = (d == null ? DateTime.Today.Date : d.Value);
            // if there is a backup for today
            if (m_backupDates.Exists(x => x.Date == date))
            {
                // if the file exists
                if (File.Exists(Path.Combine(RootDir, "backup", DateToString(date), filename)))
                {
                    // return the file text
                    return File.ReadAllText(Path.Combine(RootDir, "backup", DateToString(date), filename));
                }
            }
            return "";
        }

        static public T ImportDerived<T>(string json)
        {
            Dictionary<string, string> obj = (new StringStream(json)).ParseJSON();
            T derived = default(T);
            if (obj["type"] != "")
            {
                Type type = Type.GetType(obj["type"]);
                derived = (T)Activator.CreateInstance(type, json);
            }
            return derived;
        }
        #endregion
        #region Static Properties
        private static List<DateTime> m_backupDates = new List<DateTime>();
        public static string RootDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        public static List<DateTime> BackupDates
        {
            get
            {
                return m_backupDates;
            }
        }
        #endregion
    }
}
