using luval.word.navigator.terminal.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.word.navigator.terminal
{
    public class ResolverCache<TKey, TValue> where TValue: BaseIdEntity
    {
        private static Dictionary<TKey, TValue> _chache = new Dictionary<TKey, TValue>();

        public TValue Get(TKey key, Func<TKey, TValue> get, Func<TKey, TValue> create)
        {
            if (!_chache.ContainsKey(key))
            {
                _chache[key] = get(key);
                if(_chache[key] == null)
                    _chache[key] = create(key);
            }
            return _chache[key];
        }
    }
}
