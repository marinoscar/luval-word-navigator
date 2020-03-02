using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.word.navigator.terminal.Entities
{
    public abstract class BaseIdEntity
    {

        protected string TableName { get; }
        public BaseIdEntity(string tableName)
        {
            Id = Guid.NewGuid();
            TableName = tableName;
        }
        public Guid Id { get; set; }
        public string Name { get; set; }

        public virtual string ToSqlInsert()
        {
            return ToSqlInsert(TableName);
        }
        
        protected virtual string ToSqlInsert(string tableName)
        {
            return string.Format("INSERT INTO {0} (Id, Name) VALUES ('{1}', '{2}')", tableName, Id, Name.Replace("'", ""));
        }

        public virtual string ToSqlSelect()
        {
            return string.Format("SELECT * FROM {0} WHERE Name = '{1}'", TableName, Name.Replace("'", ""));
        }
    }
}
