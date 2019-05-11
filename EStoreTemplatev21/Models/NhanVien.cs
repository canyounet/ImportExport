using System;
using System.Collections.Generic;

namespace EStoreTemplatev21.Models
{
    public partial class NhanVien
    {
        public NhanVien()
        {
            ChuDe = new HashSet<ChuDe>();
            HoaDon = new HashSet<HoaDon>();
            HoiDap = new HashSet<HoiDap>();
            PhanCong = new HashSet<PhanCong>();
        }

        public string MaNv { get; set; }
        public string HoTen { get; set; }
        public string Email { get; set; }
        public string MatKhau { get; set; }

        public ICollection<ChuDe> ChuDe { get; set; }
        public ICollection<HoaDon> HoaDon { get; set; }
        public ICollection<HoiDap> HoiDap { get; set; }
        public ICollection<PhanCong> PhanCong { get; set; }
    }
}
