using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.word.navigator.terminal.Entities
{
    public class DTP_Apps
    {
        public Guid Id { get; set; }
        public Guid DTPId { get; set; }
        public Guid AppId { get; set; }

        public string ToSqlInsert()
        {
            return string.Format("INSERT INTO DTP_Apps VALUES ('{0}','{1}','{2}')", Id, DTPId, AppId);
        }
    }
}
