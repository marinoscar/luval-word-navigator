using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.word.navigator.terminal.Entities
{
    public class DTP_Trans
    {
        public Guid Id { get; set; }
        public Guid DTPId { get; set; }
        public Guid TransId { get; set; }

        public string ToSqlInsert()
        {
            return string.Format("INSERT INTO DTP_Trans VALUES ('{0}','{1}','{2}')", Id,DTPId, TransId);
        }
    }
}
