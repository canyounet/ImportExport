using System;
using System.Collections.Generic;

namespace EStoreTemplatev21.Models
{
    public partial class ChiTietHd
    {
        public int MaCt { get; set; }
        public int MaHd { get; set; }
        public int MaHh { get; set; }
        public double DonGia { get; set; }
        public int SoLuong { get; set; }
        public double GiamGia { get; set; }

        public HoaDon MaHdNavigation { get; set; }
        public HangHoa MaHhNavigation { get; set; }
    }
}
