using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mimeo.MimeoConnect
{
    [Flags]

    public enum StoreItemLevelOfDetail
    {
        Default = 0,
        IncludeItemData = 1,
        IncludeItemDetails = 2,
        IncludeDescription = 4,
        IncludeSystemDetails = 8,
        IncludeCreatedByUserDetails = 16,
        IncludeFolder = 32,
        IncludeLastModifiedByUserDetails = 64,
        All = 2147483647,
    }
}
