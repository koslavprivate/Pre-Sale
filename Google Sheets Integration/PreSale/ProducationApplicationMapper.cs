using PreSale.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PreSale
{
    public static class ProducationApplicationMapper
    {
        public static List<cr254_ProducationApplication> MapFromRangeData(IList<IList<object>> values)
        {
            var items = new List<cr254_ProducationApplication>();

            foreach (var value in values)
            {
                var item = new cr254_ProducationApplication()
                {
                    cr254_Name = value[0].ToString(),
                    cr254_Field1 = value[1].ToString(),
                    cr254_Field2 = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(value[2].ToString())),
                    cr254_Field3 = value[3].ToString(),
                    Id = new Guid(value[4].ToString())
                };

                items.Add(item);
            }

            return items;
        }

        public static IList<IList<object>> MapToRangeData(cr254_ProducationApplication application)
        {
            var objectList = new List<object>() 
            { 
                application.cr254_Name, 
                application.cr254_Field1, 
                application.cr254_Field2, 
                application.cr254_Field3, 
                application.Id
            };
            var rangeData = new List<IList<object>> { objectList };
            return rangeData;
        }
    }
}
