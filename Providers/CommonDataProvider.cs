using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Priem
{
    public static class CommonDataProvider
    {
        public static int GetRegionIdForCountryId(int countryId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                return context.Region.Where(x => x.CountryId == countryId).Select(x => x.Id).FirstOrDefault();
            }
        }

        public static List<KeyValuePair<string, string>> GetRegionListForCountryId(int countryId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                return context.Region.Where(x => x.CountryId == countryId)
                    .Select(x => new { x.Id, x.Name, x.Distance, x.RegionNumber }).OrderBy(x => x.Distance).ThenBy(x => x.RegionNumber)
                    .ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), "[" + (x.RegionNumber ?? "") + "] " + x.Name)).ToList();
            }
        }
    }
}
