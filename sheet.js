const NamHienTai = 2025;
let MySheeet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1J3vCj4RPBWNUvJeR6_llmIhhOWTinNGLRmhhQOCHdFA/edit#gid=0");
let SaoHanSheet = MySheeet.getSheetByName("SaoHan");

function formatFirstLetterUppercase(input) {
  return input.toLowerCase().replace(/(^|\s)\S/g, function (firstLetter) {
    return firstLetter.toLocaleUpperCase();
  });
}

function doPost(e) {
  let SaoHan = e.parameters;
  let HoVaTen = SaoHan.ho_va_ten.map(name => formatFirstLetterUppercase(name));
  let DaiDien = formatFirstLetterUppercase(SaoHan.dai_dien[0]);
  let DiaChi = SaoHan.dia_chi[0];
  let dc1 = SaoHan.dc_1[0];
  // let dc1 = SaoHan.dc_1[0] ? `'${SaoHan.dc_1[0]}` : 'Khác';
  let dc2 = SaoHan.dc_2[0];
  // let dc2 = SaoHan.dc_2[0] ? `'${SaoHan.dc_2[0]}` : 'none';
  let dc3 = SaoHan.dc_3[0];
  // let dc3 = SaoHan.dc_3[0] ? `'${SaoHan.dc_3[0]}` : 'none';

  if (dc1 === "Cần Đước" || dc1 === "Long An" || dc1 === "TP.HCM") {
    dc2 = dc2
    dc3 = dc3
  } else if (dc1 === "Khác") {
    dc2 = "none"
    dc3 = "none"
  } else {
    dc2 = "none"
    dc3 = "none"
  }

  if (dc1 === "Khác") {
    DiaChi = formatFirstLetterUppercase(DiaChi);
  }
  
  let SoDienThoai = SaoHan.so_dien_thoai[0] ? `'${SaoHan.so_dien_thoai[0]}` : 'none';
  let ThemMoi = SaoHan.ThemMoi;

  if (ThemMoi == 'Y') {
    // Lấy tất cả các mã số hiện tại từ cột A
    let table = SaoHanSheet.getRange("A:A").getValues().map(row => row[0]).filter(value => value);

    // Phân loại dữ liệu
    let numbersWithoutPrefix = []; // Các số không có tiền tố
    let dataWithPrefix = [];       // Các mã số có tiền tố

    // Phân loại dữ liệu thành 2 nhóm
    table.forEach(value => {
        if (typeof value === "string" && /^[A-Z]+/.test(value)) {
            // Nếu bắt đầu bằng chữ cái, đưa vào nhóm có tiền tố
            dataWithPrefix.push(value);
        } else if (!isNaN(value)) {
            // Nếu là số (không có tiền tố)
            numbersWithoutPrefix.push(Number(value));
        }
    });

    // Xác định tiền tố dựa trên dc1
    let prefix = "N/A"; // Mặc định là "K" nếu không thuộc các trường hợp cụ thể
    if (dc1 === "Cần Đước") {
        // Xác định tiền tố theo dc2
        switch (dc2) {
            case "TT.Cần Đước": prefix = "TCĐ"; break;
            case "Xã Long Trạch": prefix = "XLT"; break;
            case "Xã Long Khê": prefix = "XLK"; break;
            case "Xã Long Định": prefix = "XLĐ"; break;
            case "Xã Phước Vân": prefix = "XPV"; break;
            case "Xã Long Hòa": prefix = "XLH"; break;
            case "Xã Long Cang": prefix = "XLC"; break;
            case "Xã Long Sơn": prefix = "XLS"; break;
            case "Xã Tân Trạch": prefix = "XTT"; break;
            case "Xã Mỹ Lệ": prefix = "XML"; break;
            case "Xã Tân Lân": prefix = "XTL"; break;
            case "Xã Phước Tuy": prefix = "XPT"; break;
            case "Xã Long Hựu Đông": prefix = "XLHĐ"; break;
            case "Xã Tân Ân": prefix = "XTÂ"; break;
            case "Xã Phước Đông": prefix = "XPĐ"; break;
            case "Xã Long Hựu Tây": prefix = "XLHT"; break;
            case "Xã Tân Chánh": prefix = "XTC"; break;
        }
    } else if (dc1 === "Long An") {
        prefix = "LA";
    } else if (dc1 === "TP.HCM") {
        prefix = "HCM";
    }

    // Lọc các mã số có tiền tố tương ứng
    let filteredTable = dataWithPrefix.filter(maSo => maSo.startsWith(prefix));

    // Tìm số lớn nhất cho mã số có tiền tố
    let maxNumberWithPrefix = 0;
    if (filteredTable.length > 0) {
        maxNumberWithPrefix = Math.max(
            ...filteredTable.map(maSo => {
                let numberPart = maSo.replace(prefix, "");
                return isNaN(parseInt(numberPart)) ? 0 : parseInt(numberPart);
            })
        );
    } else {
        // Nếu tiền tố chưa tồn tại, bắt đầu với mã số đầu tiên
        maxNumberWithPrefix = 0;
    }

    // Tìm số lớn nhất cho các số không có tiền tố
    let maxNumberWithoutPrefix = numbersWithoutPrefix.length > 0 ? Math.max(...numbersWithoutPrefix) : 0;

    // Tạo mã số mới
    let MaSo;
    if (dc1 === "Không tiền tố") {
        // Nếu muốn thêm số không có tiền tố
        MaSo = (maxNumberWithoutPrefix + 1).toString();
    } else {
        // Nếu là mã số có tiền tố
        MaSo = prefix + String(maxNumberWithPrefix + 1).padStart(3, '0');
    }

    // Thêm dòng chính
    SaoHanSheet.appendRow([MaSo, DaiDien, DiaChi, dc1, dc2, dc3, SoDienThoai, "Chưa in"]);

    // Thêm các thông tin chi tiết
    for (let index = 1; index < HoVaTen.length; index++) {
        let HoVaTenValue = HoVaTen[index];
        let GioiTinhValue = SaoHan.gioi_tinh[index];
        let NamSinhValue = 0;
        let NguoiSanhValue = SaoHan.nguoi_sanh[index];

        if (SaoHan.tuoi[index] !== "") {
            NamSinhValue = NamHienTai - SaoHan.tuoi[index] + 1;
        } else {
            NamSinhValue = "";
            NguoiSanhValue = "";
        }

        if (HoVaTenValue !== "" && GioiTinhValue !== "") {
            SaoHanSheet.appendRow(["", "", "", "", "", "", "", "", HoVaTenValue, GioiTinhValue, NamSinhValue, NguoiSanhValue]);
        }
    }
}

  else {
    SaoHanSheet.getRange(SaoHan.StartRow, 1, 1, 8).setValues([[SaoHan.ma_so[0], DaiDien, DiaChi, dc1, dc2, dc3, SoDienThoai, "Chưa in"]]);
    let NextRow = SaoHan.StartRow;
    let DeleteRow = +SaoHan.StartRow + 1;
    SaoHanSheet.deleteRows(DeleteRow, SaoHan.RowCount-1);
    for (let index = 1; index < HoVaTen.length; index++) {
      let HoVaTenValue = HoVaTen[index];
      let GioiTinhValue = SaoHan.gioi_tinh[index];
      let NamSinhValue = 0;
      let NguoiSanhValue = SaoHan.nguoi_sanh[index];
      if (SaoHan.tuoi[index] !== "") {
        NamSinhValue = NamHienTai - SaoHan.tuoi[index] + 1;
      } else {
        NamSinhValue = "";
        NguoiSanhValue = "";
      }
      if (HoVaTenValue !== "" && GioiTinhValue !== "") {
        NextRow++;
        SaoHanSheet.insertRows(NextRow, 1);
        SaoHanSheet.getRange(NextRow, 1, 1, 12).setValues([["", "", "", "", "", "", "", "", HoVaTenValue, GioiTinhValue , NamSinhValue, NguoiSanhValue]]);
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ message: "Đã lưu thành công!" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  let page = e.parameter.page;
  
  // Trường hợp không có trang hoặc trang không xác định
  if (page == null || page == undefined) {                      
    let table = SaoHanSheet.getRange("A:H").getValues().filter(r=>r.every(Boolean));
    let str = JSON.stringify(table);
    return ContentService.createTextOutput(str);
  }
  // Trường hợp trang 'max': trả về giá trị mã số lớn nhất
  else if (page == 'max') {
    let table = SaoHanSheet.getRange("A:A").getValues().filter(Number);
    let myMax = Math.max(...table);
    return ContentService.createTextOutput(myMax);
  }

  if (page == "print") {
    // Tìm dòng chứa mã số và kiểm tra giá trị hiện tại của cột thứ 8
    let no = e.parameter.no;
    let rowToUpdate = SaoHanSheet.getRange("A:A").createTextFinder(no).matchEntireCell(true).findNext();
    
    if (rowToUpdate) {
        let rowToUpdateIndex = rowToUpdate.getRow();
        let currentValue = SaoHanSheet.getRange(rowToUpdateIndex, 8).getValue();

        // Kiểm tra giá trị hiện tại và thay đổi nó
        if (currentValue === "Chưa in") {
            SaoHanSheet.getRange(rowToUpdateIndex, 8).setValue("R");
        } else if (currentValue === "R") {
            SaoHanSheet.getRange(rowToUpdateIndex, 8).setValue("Chưa in");
        } else {
            SaoHanSheet.getRange(rowToUpdateIndex, 8).setValue("R");
        }
    }
  }
  
  if (page == 'search') {
    // Tìm kiếm dữ liệu dựa trên số
    let no = e.parameter.no;

    let ReturnData = SaoHanSheet.getRange("A:A").createTextFinder(no).matchEntireCell(true).findAll();
    let StartRow = 0;
    let EndRow = 0;

    if (ReturnData.length > 0) {
      StartRow = ReturnData[0].getRow();

      // Tìm dòng cuối cùng chứa dữ liệu liên quan
      let lastRowWithData = SaoHanSheet.getLastRow();
      for (var i = StartRow + 1; i <= lastRowWithData; i++) {
        let val = SaoHanSheet.getRange(i, 9).getValue();
        if (val == "") {
          EndRow = i - 1;
          break;
        }
      }

      // Nếu không tìm thấy EndRow, tức là có hơn 10 dòng liên tiếp chứa dữ liệu
      if (EndRow === 0) {
        EndRow = lastRowWithData;
      }

      // Lấy dữ liệu từ bảng
      let table = SaoHanSheet.getRange("A" + StartRow + ":L" + EndRow).getValues();
      let cnt = EndRow - StartRow + 1;
      let str = JSON.stringify({ record: table, SR: StartRow, CNT: cnt });

      return ContentService.createTextOutput(str);
    } else {
      let str = JSON.stringify("NOT FOUND");
      return ContentService.createTextOutput(str);
    }
  }
  // else if (page == 'all') {
  //      // Lấy toàn bộ dữ liệu từ bảng
  //      let table = SaoHanSheet.getRange("A:H").getValues().filter(r=>r.every(Boolean));
  //      let str = JSON.stringify(table);
  //      return ContentService.createTextOutput(str); 
  // }
  else if (page == 'all') {
    // Lấy toàn bộ dữ liệu từ bảng
    let data = SaoHanSheet.getRange("A:L").getValues();

    // Mảng để lưu kết quả
    let result = [];
    let currentMain = null;

    // Duyệt qua từng dòng
    data.forEach(row => {
        if (row[0] !== "" && row[0] !== null) {
            // Nếu dòng có mã số, khởi tạo một đối tượng mới
            currentMain = {
                maSo: row[0], // Mã số
                daiDien: row[1], // Đại diện
                diaChi: row[2], // Địa chỉ
                dc1: row[3],
                dc2: row[4],
                dc3: row[5],
                soDienThoai: row[6],
                trangThai: row[7],
                thanhVien: [] // Danh sách thành viên
            };
            result.push(currentMain); // Thêm vào mảng kết quả
        } else if (currentMain && row[8] !== "" && row[8] !== null) {
            // Nếu dòng không có mã số, thêm vào danh sách thành viên
            currentMain.thanhVien.push({
                hoVaTen: row[8], // Họ và tên
                gioiTinh: row[9], // Giới tính
                namSinh: row[10], // Năm sinh
                nguoiSinh: row[11] // Ngươi sanh
            });
        }
    });

    // Chuyển mảng thành JSON và trả về
    let str = JSON.stringify(result);
    return ContentService.createTextOutput(str);
  } else if (page == "set") {
    // Lấy toàn bộ dữ liệu từ bảng (cột A đến cột F)
    let data = SaoHanSheet.getRange("A:F").getValues();

    // Biến đếm cho mỗi tiền tố
    let prefixCounters = {};

    // Duyệt qua từng dòng để tạo mã số mới
    for (let i = 1; i < data.length; i++) { // Bắt đầu từ dòng 2, bỏ qua tiêu đề
        let maSo = data[i][0]; // Cột A - Mã số
        let dc1 = data[i][3]; // Cột D - dc1
        let dc2 = data[i][4]; // Cột E - dc2

        // Kiểm tra nếu cột A trống, bỏ qua dòng này
        if (!maSo || maSo === "") {
            continue;
        }

        // Xác định tiền tố dựa trên dc1 và dc2
        let prefix = "N/A"; // Mặc định nếu không thuộc trường hợp cụ thể

        if (dc1 === "Cần Đước") {
            switch (dc2) {
                case "TT.Cần Đước": prefix = "TCĐ"; break;
                case "Xã Long Trạch": prefix = "XLT"; break;
                case "Xã Long Khê": prefix = "XLK"; break;
                case "Xã Long Định": prefix = "XLĐ"; break;
                case "Xã Phước Vân": prefix = "XPV"; break;
                case "Xã Long Hòa": prefix = "XLH"; break;
                case "Xã Long Cang": prefix = "XLC"; break;
                case "Xã Long Sơn": prefix = "XLS"; break;
                case "Xã Tân Trạch": prefix = "XTT"; break;
                case "Xã Mỹ Lệ": prefix = "XML"; break;
                case "Xã Tân Lân": prefix = "XTL"; break;
                case "Xã Phước Tuy": prefix = "XPT"; break;
                case "Xã Long Hựu Đông": prefix = "XLHĐ"; break;
                case "Xã Tân Ân": prefix = "XTÂ"; break;
                case "Xã Phước Đông": prefix = "XPĐ"; break;
                case "Xã Long Hựu Tây": prefix = "XLHT"; break;
                case "Xã Tân Chánh": prefix = "XTC"; break;
                default: prefix = "K"; // Nếu không khớp, mặc định là "K"
            }
        } else if (dc1 === "Long An") {
            prefix = "LA";
        } else if (dc1 === "TP.HCM") {
            prefix = "HCM";
        }

        // Khởi tạo bộ đếm cho tiền tố nếu chưa có
        if (!prefixCounters[prefix]) {
            prefixCounters[prefix] = 0;
        }

        // Tăng bộ đếm và tạo mã số mới
        prefixCounters[prefix]++;
        let newMaSo = prefix + String(prefixCounters[prefix]).padStart(3, '0');

        // Ghi mã số mới vào cột A
        SaoHanSheet.getRange(i + 1, 1).setValue(newMaSo);
    }

    // Trả về kết quả
    return ContentService.createTextOutput(
        JSON.stringify({ message: "Cập nhật mã số thành công!" })
    ).setMimeType(ContentService.MimeType.JSON);
}


}
