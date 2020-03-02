using luval.word.navigator.terminal.Entities;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace luval.word.navigator.terminal
{
    public class ResultProcessor
    {
        private Database _db;
        private ResolverCache<string, DTP> _dtpResolver;
        private ResolverCache<string, Applications> _appResolver;
        private ResolverCache<string, Transactions> _tranResolver;
        public ResultProcessor()
        {
            _db = new Database(() => { return new SqlConnection("Server=.;Database=BPO;Trusted_Connection=True;"); });
            _dtpResolver = new ResolverCache<string, DTP>();
            _appResolver = new ResolverCache<string, Applications>();
            _tranResolver = new ResolverCache<string, Transactions>();
        }
        public void ImportToDb(IEnumerable<DocumentData> items)
        {
            foreach (var item in items)
                ProcessItem(item);
        }

        public void ProcessItem(DocumentData item)
        {
            var dtp = GetOrCreateDtp(item.FileName);
            var apps = GetApps(item.Systems);
            var trans = GetTrans(item.SAPTransactionCodes);
            apps.ForEach(i => MarryApp(dtp, i));
            trans.ForEach(i => MarryTran(dtp, i));
        }

        public void MarryTran(DTP dtp, Transactions tran)
        {
            _db.ExecuteNonQuery(string.Format("INSERT INTO DTP_Trans VALUES ('{0}','{1}','{2}')", Guid.NewGuid(), dtp.Id, tran.Id));
        }

        public void MarryApp(DTP dtp, Applications apps)
        {
            _db.ExecuteNonQuery(string.Format("INSERT INTO DTP_Apps VALUES ('{0}','{1}','{2}')", Guid.NewGuid(), dtp.Id, apps.Id));
        }

        public DTP GetOrCreateDtp(string name)
        {
            return GetOrCreate<DTP>(name, _dtpResolver);
        }

        public Applications GetOrCreateApp(string name)
        {
            return GetOrCreate<Applications>(name, _appResolver);
        }

        public Transactions GetOrCreateTrans(string name)
        {
            return GetOrCreate<Transactions>(name, _tranResolver);
        }

        public TValue GetOrCreate<TValue>(string name, ResolverCache<string, TValue> resolverCache) where TValue : BaseIdEntity
        {
            return resolverCache.Get(name, 
                (k) => { return GetValue<TValue>(k); }, 
                (k) => { return CreateEntity<TValue>(k); });
        }

        public TValue GetValue<TValue>(string key) where TValue : BaseIdEntity
        {
            var entity = Activator.CreateInstance<TValue>();
            entity.Name = key;
            var record = _db.ExecuteToDictionaryList(entity.ToSqlSelect()).FirstOrDefault();
            if (record == null)
                return default;
            entity.Name = Convert.ToString(record["Name"]);
            entity.Id = (Guid)record["Id"];
            return (TValue)entity;
        }

        public T CreateEntity<T>(string name) where T : BaseIdEntity
        {
            var entity = Activator.CreateInstance<T>();
            entity.Name = name;
            _db.ExecuteNonQuery(entity.ToSqlInsert());
            return entity;
        }

        

        public List<Applications> GetApps(string apps)
        {
            var res = new List<Applications>();
            var appNames = GetAppNames(apps);
            foreach(var appName in appNames)
            {
                res.Add(GetOrCreateApp(appName));
            }
            return res;
        }

        public List<Transactions> GetTrans(string trans)
        {
            var res = new List<Transactions>();
            var tranNames = GetTransNames(trans);
            foreach (var tran in tranNames)
            {
                res.Add(GetOrCreateTrans(tran));
            }
            return res;
        }

        private List<string> GetAppNames(string apps)
        {
            if (string.IsNullOrWhiteSpace(apps)) return new List<string>();
            var list = apps.Split(";".ToCharArray()).ToList();
            var res = new List<string>();
            foreach(var app in list)
            {
                if (!string.IsNullOrEmpty(app) && (app.Contains("SAP") || app.Contains("ECC")))
                    res.Add("SAP");
                else
                    res.Add(app);
            }
            return res.Distinct().ToList();
        }

        private List<string> GetTransNames(string tans)
        {
            return tans.Split(";".ToCharArray()).ToList();
        }
    }
}
