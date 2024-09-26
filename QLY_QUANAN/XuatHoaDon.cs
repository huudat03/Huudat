using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QLY_QUANAN
{
    internal class XuatHoaDon
    {
        public static bool xuatHoaDon(string content, System.Data.DataTable dataTable, string billId, string customerName, string orderDate, float fullValue, string tenNv)
        {
            try
            {
                //Tạo các đối tượng Excel cần thiết để thao tác với file Excel: Application, Workbooks, Sheets, Workbook, và Worksheet
                Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbooks oBooks;
                Microsoft.Office.Interop.Excel.Sheets oSheets;
                Microsoft.Office.Interop.Excel.Workbook oBook;
                Microsoft.Office.Interop.Excel.Worksheet oSheet;

                //Tạo mới một Excel WorkBook với một sheet
                //Đặt tên sheet theo content.
                oExcel.Visible = true;//Thiết lập thuộc tính Visible của đối tượng Excel để làm cho cửa sổ Excel hiển thị lên màn hình. Khi đặt giá trị là true, cửa sổ Excel sẽ hiện ra cho người dùng thấy.
                oExcel.DisplayAlerts = false;//Thiết lập thuộc tính DisplayAlerts để tắt các thông báo cảnh báo trong Excel. Khi đặt giá trị là false, Excel sẽ không hiển thị các hộp thoại cảnh báo khi thực hiện các thao tác.
                oExcel.Application.SheetsInNewWorkbook = 1;//Đặt số lượng trang tính mặc định trong một sổ làm việc mới là 1.
                oBooks = oExcel.Workbooks;//Lấy tập hợp các sổ làm việc (Workbooks) hiện tại từ đối tượng Excel.
                oBook = (Microsoft.Office.Interop.Excel.Workbook)(oExcel.Workbooks.Add(Type.Missing));
                //Tạo một sổ làm việc mới và gán nó vào biến oBook. Add(Type.Missing) được sử dụng để thêm sổ làm việc mới với các tham số mặc định.
                oSheets = oBook.Worksheets;//Lấy tập hợp các trang tính (Worksheets) từ sổ làm việc vừa tạo.
                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oSheets.get_Item(1);
                //Lấy trang tính đầu tiên từ tập hợp các trang tính của sổ làm việc và gán nó vào biến oSheet.
                string sheetName = content;
                string title = "Đơn hàng " + content;

                oSheet.Name = sheetName;//Đặt tên cho trang tính hiện tại thành giá trị của biến sheetName.

                //Tiêu đề định dạng từ cột A1 đến E1
                Microsoft.Office.Interop.Excel.Range head = oSheet.get_Range("A1", "E1");//khai báo 1 biến head kiểu range gán nó bằng phạm vi các ô từ "A1" đến "E1" trên trang tính oSheet.
                head.MergeCells = true;//Thiết lập thuộc tính MergeCells của phạm vi head để hợp nhất các ô trong phạm vi này thành một ô duy nhất.
                head.Value2 = title;//Đặt giá trị cho ô đã hợp nhất (phạm vi head) bằng giá trị của biến title
                head.Font.Bold = true;//in đậm văn bản.
                head.Font.Name = "Times New Roman";
                head.Font.Size = "20";
                head.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //căn chỉnh văn bản trong phạm vi head theo chiều ngang ở giữa ô (center).
                
                //Điền các thông tin về mã hóa đơn, tên khách hàng, ngày đặt, tổng giá trị hóa đơn, và tên nhân viên vào các ô tương ứng trong bảng Excel.
                int billIdRow = 1;//hàng
                Microsoft.Office.Interop.Excel.Range billIdLabelCell = oSheet.get_Range("A" + billIdRow, "A" + billIdRow);//gán gán nó bằng phạm vi ô  A1
                billIdLabelCell.Value2 = "Mã hóa đơn: " + billId;

                //Thông tin
                int customerRow = 2;
                Microsoft.Office.Interop.Excel.Range customerNameLabelCell = oSheet.get_Range("A" + customerRow, "A" + customerRow);
                customerNameLabelCell.Value2 = "Tên khách hàng:";
                Microsoft.Office.Interop.Excel.Range customerNameValueCell = oSheet.get_Range("B" + customerRow, "B" + customerRow);//gán nó bằng phạm vi ô 
                customerNameValueCell.Value2 = customerName;

                int orderDateRow = customerRow + 1;
                Microsoft.Office.Interop.Excel.Range orderDateLabelCell = oSheet.get_Range("A" + orderDateRow, "A" + orderDateRow);
                orderDateLabelCell.Value2 = "Ngày đặt:";// Đặt giá trị cho ô orderDateLabelCell là chuỗi "Ngày đặt:".
                Microsoft.Office.Interop.Excel.Range orderDateValueCell = oSheet.get_Range("B" + orderDateRow, "B" + orderDateRow);
                orderDateValueCell.Value2 = orderDate;//Đặt giá trị cho ô "B3" là giá trị của biến orderDate.

                int totalBillValueRow = orderDateRow + 1;//nằm ngay dưới hàng chứa thông tin ngày đặt hàng.
                Microsoft.Office.Interop.Excel.Range totalBillValueLabelCell = oSheet.get_Range("A" + totalBillValueRow, "A" + totalBillValueRow);
                totalBillValueLabelCell.Value2 = "Tổng giá trị hóa đơn:";
                Microsoft.Office.Interop.Excel.Range totalBillValueValueCell = oSheet.get_Range("B" + totalBillValueRow, "B" + totalBillValueRow);
                totalBillValueValueCell.Value2 = fullValue.ToString();
                //Đặt giá trị cho ô totalBillValueValueCell bằng giá trị của biến fullValue sau khi chuyển đổi sang chuỗi (ToString())
                int tenNhanVienRow = totalBillValueRow + 1;
                Microsoft.Office.Interop.Excel.Range tenNhanVienLabelCell = oSheet.get_Range("A" + tenNhanVienRow, "A" + tenNhanVienRow);
                tenNhanVienLabelCell.Value2 = "Người tạo bill:";
                Microsoft.Office.Interop.Excel.Range tenNhanVienValueCell = oSheet.get_Range("B" + tenNhanVienRow, "B" + tenNhanVienRow);
                tenNhanVienValueCell.Value2 = tenNv;
                tenNhanVienValueCell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                //Thiết lập thuộc tính HorizontalAlignment để căn chỉnh văn bản trong ô tenNhanVienValueCell theo chiều ngang ở bên phải (right)

                //Thiết lập các tiêu đề cột cho bảng chi tiết hóa đơn (thứ tự bill, tên món ăn, đơn giá, số lượng, tổng giá) và căn giữa các tiêu đề
                int columnRow = tenNhanVienRow + 1;

                //Tiêu đề bảng


                Microsoft.Office.Interop.Excel.Range cl1 = oSheet.get_Range("A" + columnRow, "A" + columnRow);//gán nó bằng phạm vi ô 
                cl1.Value2 = "Thứ tự bill";// Đặt giá trị cho ô cl1 là chuỗi "Thứ tự bill".
                cl1.ColumnWidth = 12;//Thiết lập độ rộng của cột chứa ô cl1 là 12 đơn vị

                Microsoft.Office.Interop.Excel.Range cl2 = oSheet.get_Range("B" + columnRow, "B" + columnRow);
                cl2.Value2 = "Tên món ăn";
                cl2.ColumnWidth = 30.29;

                Microsoft.Office.Interop.Excel.Range cl3 = oSheet.get_Range("C" + columnRow, "C" + columnRow);
                cl3.Value2 = "Đơn giá";
                cl3.ColumnWidth = 14;

                Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D" + columnRow, "D" + columnRow);
                cl4.Value2 = "Số lượng";
                cl4.ColumnWidth = 23.71;

                Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("E" + columnRow, "E" + columnRow);
                cl5.Value2 = "Tổng giá";
                cl5.ColumnWidth = 10.71;

                Microsoft.Office.Interop.Excel.Range rowHead = oSheet.get_Range("A" + columnRow, "E" + columnRow);


                // Thiết lập màu nền
                int size = dataTable.Columns.Count;//số lượng cột trong dataTable.


                rowHead.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //căn chỉnh ngang của dải ô rowHead là giữa trên trang tính Excel

                // Tạo một mảng hai chiều để chứa dữ liệu từ DataTable.

                string[,] arr = new string[dataTable.Rows.Count, dataTable.Columns.Count];

                //Chuyển dữ liệu từ DataTable vào mảng đối tượng

                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    DataRow dataRow = dataTable.Rows[row]; //Lấy ra DataRow tương ứng với hàng hiện tại trong dataTable.

                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        arr[row, col] = dataRow[col].ToString();//Gán giá trị của ô tại hàng row và cột col trong mảng arr,ép kiểu trả về là tosing
                    }
                }

                //Thiết lập vùng điền dữ liệu

                int rowStart = 7;//Xác định hàng bắt đầu của vùng dữ liệu trên trang tính Excel

                int columnStart = 1;

                int rowEnd = rowStart + dataTable.Rows.Count - 1;//Xác định hàng kết thúc của vùng dữ liệu

                int columnEnd = dataTable.Columns.Count;

                // Ô bắt đầu điền dữ liệu

                Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowStart, columnStart];

                // Ô kết thúc điền dữ liệu

                Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnEnd];

                // Lấy về vùng điền dữ liệu

                Microsoft.Office.Interop.Excel.Range range = oSheet.get_Range(c1, c2);//Tạo ra một đối tượng Range để đại diện cho một vùng dữ liệu trên trang tính Excel.

                //Điền dữ liệu vào vùng đã thiết lập

                range.Value2 = arr;//Dòng mã này gán giá trị của mảng arr vào vùng dữ liệu trên trang tính Excel được đại diện bởi đối tượng Range range.

                // Kẻ viền

                range.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;
                //Đây là giá trị hằng số (constant) trong Excel, đại diện cho kiểu đường viền liền. Khi được gán cho thuộc tính LineStyle,
                //các đường viền của vùng dữ liệu sẽ được hiển thị dưới dạng đường viền liền.

                // Căn giữa cột mã nhân viên

                Microsoft.Office.Interop.Excel.Range c3 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnStart];

                Microsoft.Office.Interop.Excel.Range c4 = oSheet.get_Range(c1, c3);

                //Căn giữa cả bảng 
                oSheet.get_Range(c1, c2).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;

        }
    }
}
