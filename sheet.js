const NamHienTai = 2024;
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
  let dc2 = SaoHan.dc_2[0];
  let dc3 = SaoHan.dc_3[0];

  if (dc1 === "Khác") {
    DiaChi = formatFirstLetterUppercase(DiaChi);
  }
  
  let SoDienThoai = SaoHan.so_dien_thoai[0] ? `'${SaoHan.so_dien_thoai[0]}` : 'none';
  let ThemMoi = SaoHan.ThemMoi;

  if (ThemMoi == 'Y') {
    let table = SaoHanSheet.getRange("A:A").getValues().filter(Number);
    let MaSo = Math.max(...table) + 1 ;
    SaoHanSheet.appendRow([MaSo, DaiDien, DiaChi, dc1, dc2, dc3, SoDienThoai, "Chưa in"]);
    for (let index = 1; index < HoVaTen.length; index++) {
      let HoVaTenValue = HoVaTen[index];
      let GioiTinhValue = SaoHan.gioi_tinh[index];
      let NamSinhValue = NamHienTai - SaoHan.tuoi[index] + 1;
      let NguoiSanhValue = SaoHan.nguoi_sanh[index];
      if (HoVaTenValue !== "" && GioiTinhValue !== "" && NamSinhValue !== "" && NguoiSanhValue !== "") {
        SaoHanSheet.appendRow(["", "", "", "", "", "", "", "", HoVaTenValue, GioiTinhValue , NamSinhValue, NguoiSanhValue]);
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
      let NamSinhValue = NamHienTai - SaoHan.tuoi[index] + 1;
      let NguoiSanhValue = SaoHan.nguoi_sanh[index];
      if (HoVaTenValue !== "" && GioiTinhValue !== "" && NamSinhValue !== "" && NguoiSanhValue !== "") {
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
  else if (page == 'all') {
       // Lấy toàn bộ dữ liệu từ bảng
       let table = SaoHanSheet.getRange("A:H").getValues().filter(r=>r.every(Boolean));
       let str = JSON.stringify(table);
       return ContentService.createTextOutput(str); 
  }
}
