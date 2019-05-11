using System;
using System.Collections.Generic;

namespace EStoreTemplatev21.Models
{
    public partial class PhongBan
    {
        public PhongBan()
        {
            PhanCong = new HashSet<PhanCong>();
            PhanQuyen = new HashSet<PhanQuyen>();
        }

        public string MaPb { get; set; }
        public string TenPb { get; set; }
        public string ThongTin { get; set; }

        public ICollection<PhanCong> PhanCong { get; set; }
        public ICollection<PhanQuyen> PhanQuyen { get; set; }
    }
}
