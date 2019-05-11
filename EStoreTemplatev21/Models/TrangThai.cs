using System;
using System.Collections.Generic;

namespace EStoreTemplatev21.Models
{
    public partial class TrangThai
    {
        public TrangThai()
        {
            HoaDon = new HashSet<HoaDon>();
        }

        public int MaTrangThai { get; set; }
        public string TenTrangThai { get; set; }
        public string MoTa { get; set; }

        public ICollection<HoaDon> HoaDon { get; set; }
    }
}
