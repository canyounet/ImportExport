using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EStoreTemplatev21.Models
{
    public static class TemplateGenerator
    {
        public static string GetHTMLString()
        {
            var ctx = new MyeStoreContext();
            var employees = ctx.Loai.ToList();

            var sb = new StringBuilder();
            sb.Append(@"
                        <html>
                            <head>
                            </head>
                            <body>
                                <div class='header'><h1>DANH SÁCH LOẠI</h1></div>
                                <table align='center'>
                                    <tr>
                                        <th>Mã loại</th>
                                        <th>Tên loại</th>
                                        <th>Mô tả</th>
                                        <th>Hình</th>
                                    </tr>");

            foreach (var emp in employees)
            {
                sb.AppendFormat(@"<tr>
                                    <td>{0}</td>
                                    <td>{1}</td>
                                    <td>{2}</td>
                                    <td>{3}</td>
                                  </tr>", emp.MaLoai, emp.TenLoai, emp.MoTa, emp.Hinh);
            }

            sb.Append(@"
                                </table>
                            </body>
                        </html>");

            return sb.ToString();
        }
    }
}
