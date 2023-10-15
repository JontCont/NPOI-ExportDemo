using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace App.Model.ExportDtos
{
    public class Demo
    {
        [DisplayName("名稱")]
        public string name { get; set; }

        [DisplayName("日期")]
        public DateTime date { get; set; }
    }
}
