using System;
using System.Collections.Generic;

namespace EStoreTemplatev21.Models
{
    public partial class BanBe
    {
        public int MaBb { get; set; }
        public string MaKh { get; set; }
        public int MaHh { get; set; }
        public string HoTen { get; set; }
        public string Email { get; set; }
        public DateTime NgayGui { get; set; }
        public string GhiChu { get; set; }

        public HangHoa MaHhNavigation { get; set; }
        public KhachHang MaKhNavigation { get; set; }
    }
}
