using System.Collections.Generic;

namespace Bekk.ExcelBuilder.Contracts
{
    public interface IEntityCollection<out T, in TKey> : IEnumerable<T>
    {
        T this[TKey key] { get; }
        T First();
    }
}