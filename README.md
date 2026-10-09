// Hàm tiện ích để loại bỏ hậu tố "-1" khỏi chuỗi
function removeMinusOne(str) {
  if (typeof str !== 'string') return str.toString().trim();
  if (str.endsWith("-1")) {
    return str.slice(0, -2).trim(); // Loại bỏ "-1" và khoảng trắng
  }
  return str.trim();
}

function checkAllData() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Không thể truy cập spreadsheet hiện tại.');
    }

    var ui = SpreadsheetApp.getUi();
    if (!ui) {
      throw new Error('Không thể khởi tạo giao diện người dùng.');
    }

    // Hiển thị giao diện "Đang kiểm tra"
    var checkingHtml = HtmlService.createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">' +
      '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">' +
      '<style>' +
      'body { font-family: "Roboto", sans-serif; background: linear-gradient(135deg, #e6f3ff 0%, #d1e8ff 100%); }' +
      '.progress-bar { width: 100%; background-color: #e0e0e0; height: 10px; border-radius: 5px; overflow: hidden; margin-top: 20px; }' +
      '.progress { width: 0; height: 100%; background-color: #10b981; animation: progress 2s infinite; }' +
      '@keyframes progress { 0% { width: 0; } 100% { width: 100%; } }' +
      '</style>' +
      '</head>' +
      '<body class="flex items-center justify-center min-h-screen">' +
      '<div class="bg-white p-6 rounded-lg shadow-xl max-w-md w-full text-center">' +
      '<h2 class="text-2xl font-bold text-gray-800 mb-4">Đang kiểm tra...</h2>' +
      '<h3 class="text-xl font-semibold text-blue-800 mb-4">CẤU TẠO MÃ KHUÔN LINH KIỆN</h3>' +
      '<p class="text-gray-600 mt-4">Vui lòng chờ trong khi dữ liệu được kiểm tra.</p>' +
      '<div class="progress-bar"><div class="progress"></div></div>' +
      '<audio id="startSound" autoplay><source src="https://www.soundjay.com/buttons/beep-02.mp3" type="audio/mpeg"></audio>' +
      '<script>' +
      'document.getElementById("startSound").play().catch(function(error) { console.log("Lỗi phát âm thanh: ", error); });' +
      '</script>' +
      '</div>' +
      '</body>' +
      '</html>'
    ).setWidth(400).setHeight(300);
    ui.showModalDialog(checkingHtml, 'Đang kiểm tra');

    // Đợi 1.5 giây để hiển thị giao diện "Đang kiểm tra"
    Utilities.sleep(1500);

    // Định nghĩa các sheet cần kiểm tra
    var sheetsToCheck = {
      'VA001-B': { checkFunction: checkSheet8794B },
      'VA015-A': { checkFunction: checkSheet7040GA },
      'VA015-B': { checkFunction: checkSheet7040GB },
      'VA015-E': { checkFunction: checkSheet7040GE },
      'VS001-A': { checkFunction: checkSheetVS001A },
      'VS001-B': { checkFunction: checkSheetVS001B },
      'VS001-C': { checkFunction: checkSheetVS001C },
      'VA019-A': { checkFunction: checkSheetVA019A },
      'VA019-B': { checkFunction: checkSheetVA019B },
      'VA019-C': { checkFunction: checkSheetVA019C },
      'VA003': { checkFunction: checkSheetVA003 },
      'VS051': { checkFunction: checkSheetVS051 },
      'VA016-A': { checkFunction: checkSheet7041GA },
      'VA016-B': { checkFunction: checkSheet7041GB },
    };

    // Tạo hoặc lấy sheet log và tự động điều chỉnh cột
    var logSheet = spreadsheet.getSheetByName('CheckLog') || spreadsheet.insertSheet('CheckLog');
    if (!logSheet) {
      throw new Error('Không thể tạo hoặc truy cập sheet CheckLog.');
    }
    logSheet.clear();
    logSheet.appendRow(['Thời gian', 'Tên Sheet', 'Trạng thái', 'Số lỗi', 'Thông báo']);

    // Thu thập kết quả và đánh dấu lỗi
    var results = {};
    var startTime = new Date();
    var startTimeFormatted = Utilities.formatDate(startTime, 'GMT+7', 'HH:mm:ss');
    var erroredSheet = null;
    var hasErrors = false;
    for (var sheetName in sheetsToCheck) {
      try {
        var sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
          results[sheetName] = { valid: false, errorCount: 0, message: 'Sheet không tìm thấy', details: [] };
          logSheet.appendRow([new Date(), sheetName, 'LỖI', 0, 'Sheet không tìm thấy']);
          hasErrors = true;
          continue;
        }
        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();
        if (lastRow < 7) {
          results[sheetName] = { valid: false, errorCount: 0, message: 'Không có dữ liệu từ hàng 7', details: [] };
          logSheet.appendRow([new Date(), sheetName, 'LỖI', 0, 'Không có dữ liệu từ hàng 7']);
          hasErrors = true;
          continue;
        }
       

        var result = sheetsToCheck[sheetName].checkFunction(sheet);
        results[sheetName] = result;
        logSheet.appendRow([new Date(), sheetName, result.valid ? 'OK' : 'LỖI', result.errorCount, result.message]);
        if (!result.valid && !erroredSheet) {
          erroredSheet = sheet;
          hasErrors = true;
        }
      } catch (e) {
        results[sheetName] = { valid: false, errorCount: 0, message: 'Lỗi xử lý sheet: ' + e.message, details: [] };
        logSheet.appendRow([new Date(), sheetName, 'LỖI', 0, 'Lỗi xử lý sheet: ' + e.message]);
        hasErrors = true;
      }
    }
    var endTime = new Date();
    var endTimeFormatted = Utilities.formatDate(endTime, 'GMT+7', 'HH:mm:ss');
    var duration = (endTime.getTime() - startTime.getTime()) / 1000;
    Logger.log('Thời gian thực thi: ' + duration + ' giây');

    var lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      logSheet.autoResizeColumns(1, 5);
      for (var i = 1; i <= lastRow; i++) {
        logSheet.setRowHeight(i, 21);
      }
    }

            var htmlOutput = HtmlService.createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">' +
      '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">' +
      '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">' +
      '<style>' +
      'body { font-family: "Roboto", sans-serif; ' +
      (hasErrors 
        ? 'background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%); }'
        : 'background: linear-gradient(135deg, #e6f3ff 0%, #d1e8ff 100%); }') +
      '.time-badge { background-color: #e2e8f0; color: #2c3e50; padding: 8px 12px; border-radius: 20px; display: inline-block; margin: 4px; }' +
      'table { width: 100%; border-collapse: separate; border-spacing: 0; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }' +
      'th { background-color: ' + (hasErrors ? '#dc2626' : '#10b981') + '; color: white; padding: 12px; text-align: left; }' +
      'td { padding: 12px; text-align: left; border: 1px solid #ddd; vertical-align: top; }' +
      'tr:nth-child(even) { background-color: #f8fafc; }' +
      'tr:hover { background-color: ' + (hasErrors ? '#fca5a5' : '#e6fffa') + '; transition: background-color 0.3s; }' +
      '.error-row { background-color: #dc2626; color: white; border: 2px solid #991b1b; }' +
      '.error-row:hover { background-color: #b91c1c; border-color: #7f1d1d; }' +
      '.status-icon { margin-right: 8px; }' +
      '.error-detail { font-size: 0.95em; line-height: 1.6; }' +
      '.error-detail ul { margin: 8px 0; padding-left: 22px; list-style-type: disc; }' +
      '.error-item { margin-bottom: 10px; padding: 10px; background-color: rgba(220, 38, 38, 0.15); border-left: 5px solid #dc2626; border-radius: 6px; }' +
      '.error-pos { font-weight: bold; color: #991b1b; }' +
      '.error-current { font-family: monospace; background: #ffffff; padding: 3px 8px; border-radius: 4px; font-weight: bold; color: #000000; border: 1px solid #999999; }' +
      '.error-expected { font-family: monospace; background: #ffffcc; padding: 3px 8px; border-radius: 4px; font-weight: bold; color: #000000; border: 1px solid #999999; }' +
      '.no-detail { color: #666; font-style: italic; }' +
      '</style>' +
      '</head>' +
      '<body class="flex items-center justify-center min-h-screen" onload="playSoundIfErrors(\'' + (hasErrors ? 'true' : 'false') + '\')">' +
      '<div class="bg-white p-8 rounded-lg shadow-2xl max-w-6xl w-full overflow-x-auto">' +
      '<h3 class="text-2xl font-semibold ' + (hasErrors ? 'text-red-700' : 'text-blue-800') + ' text-center mb-4">CẤU TẠO MÃ KHUÔN LINH KIỆN</h3>' +
      '<h2 class="text-3xl font-bold text-gray-800 text-center mb-6">Kết Quả Kiểm Tra</h2>' +
      '<div class="flex justify-center space-x-3 mb-6">' +
      '<span class="time-badge">Bắt đầu: ' + startTimeFormatted + '</span>' +
      '<span class="time-badge">Kết thúc: ' + endTimeFormatted + '</span>' +
      '<span class="time-badge">Thời gian: ' + duration.toFixed(2) + 's</span>' +
      '</div>' +
      '<div class="overflow-x-auto">' +
      '<table class="min-w-full">' +
      '<thead><tr>' +
      '<th class="px-4 py-3">Tên Sheet</th>' +
      '<th class="px-4 py-3">Trạng thái</th>' +
      '<th class="px-4 py-3">Số lỗi</th>' +
      '<th class="px-4 py-3">Thông báo tổng quát</th>' +
      '<th class="px-4 py-3" style="width: 450px;">Chi tiết lỗi cụ thể</th>' +
      '</tr></thead>' +
      '<tbody>' +
      (function() {
        var rows = [];
        function escapeHtml(text) {
          if (text === null || text === undefined) return '';
          return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
        }

        for (var sheetName in results) {
          var result = results[sheetName];
          var detailContent = '';

          if (result.valid) {
            detailContent = '<div class="text-green-600 font-medium">✓ Tất cả dữ liệu hợp lệ</div>';
          } else {
            if (result.details && Array.isArray(result.details) && result.details.length > 0) {
              detailContent = '<div class="error-detail"><ul>';
              result.details.forEach(function(err) {
                if (typeof err === 'object' && err !== null) {
                  var row = err.row || '?';
                  var col = err.col ? String.fromCharCode(64 + parseInt(err.col)) : '?';
                  var position = 'Hàng ' + row + ', Cột ' + col;

                  detailContent += '<li class="error-item">' +
                    '<span class="error-pos">' + position + ':</span> ' + escapeHtml(err.message || 'Lỗi không xác định') + '<br>';
                  if (err.current !== undefined) detailContent += '→ Hiện tại: <span class="error-current">' + escapeHtml(err.current) + '</span><br>';
                  if (err.expected !== undefined) detailContent += '→ Mong đợi: <span class="error-expected">' + escapeHtml(err.expected) + '</span>';
                  detailContent += '</li>';
                } else {
                  detailContent += '<li class="error-item">' + escapeHtml(err) + '</li>';
                }
              });
              detailContent += '</ul></div>';
            } else if (result.message) {
              detailContent = '<div class="error-item">' + escapeHtml(result.message) + '</div>';
            } else {
              detailContent = '<div class="no-detail">Có lỗi nhưng không có thông tin chi tiết</div>';
            }
          }

          rows.push('<tr' + (result.valid ? '' : ' class="error-row"') + '>' +
                    '<td class="border px-4 py-3 font-semibold">' + escapeHtml(sheetName) + '</td>' +
                    '<td class="border px-4 py-3"><i class="status-icon fa ' + (result.valid ? 'fa-check-circle text-green-600' : 'fa-times-circle text-white') + '"></i> <strong>' + (result.valid ? 'OK' : 'LỖI') + '</strong></td>' +
                    '<td class="border px-4 py-3 text-center text-lg font-bold">' + result.errorCount + '</td>' +
                    '<td class="border px-4 py-3">' + (result.message ? escapeHtml(result.message) : '-') + '</td>' +
                    '<td class="border px-4 py-3">' + detailContent + '</td>' +
                    '</tr>');
        }
        return rows.join('');
      })() +
      '</tbody></table>' +
      '</div>' +
      (hasErrors ? '<audio id="errorSound"><source src="https://www.soundjay.com/buttons/beep-01a.mp3" type="audio/mpeg"></audio>' : '') +
      '<script>' +
      'function playSoundIfErrors(hasErrors) {' +
      '  if (hasErrors === "true") {' +
      '    var audio = document.getElementById("errorSound");' +
      '    if (audio) audio.play().catch(function(error) { console.log("Lỗi phát âm thanh: ", error); });' +
      '  }' +
      '}' +
      '</script>' +
      '<div class="text-center mt-8">' +
      '<button onclick="google.script.host.close()" class="bg-' + (hasErrors ? 'red' : 'blue') + '-600 text-white px-8 py-4 rounded-lg hover:bg-' + (hasErrors ? 'red' : 'blue') + '-700 transition font-semibold text-lg shadow-lg">Đóng</button>' +
      '</div>' +
      '</div>' +
      '</body>' +
      '</html>'
    ).setWidth(950).setHeight(650);

    ui.showModalDialog(htmlOutput, hasErrors ? 'CẢNH BÁO HỆ THỐNG:  Phát hiện lỗi...Thông tin đã gửi đến quản lý để xử lý kịp thời ' : 'Tất cả dữ liệu đều OK');

    if (hasErrors) {
      try {
        var emailSubject = hasErrors ? 'CẢNH BÁO: Kết quả kiểm tra dữ liệu - Có lỗi' : 'Kết quả kiểm tra dữ liệu - Tất cả OK';
        var emailBody = `
  <h2 style="color:#dc2626;">⚠️ CẢNH BÁO: Phát hiện lỗi trong kiểm tra mã khuôn</h2>
  <p><strong>Thời gian bắt đầu:</strong> ${startTimeFormatted}</p>
  <p><strong>Thời gian kết thúc:</strong> ${endTimeFormatted}</p>
  <p><strong>Thời gian thực thi:</strong> ${duration.toFixed(2)} giây</p>
  <hr>
  <h3>📋 Tóm tắt kết quả</h3>
  <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%;">
    <tr style="background-color: #10b981; color: white;">
      <th style="padding: 10px;">Tên Sheet</th>
      <th style="padding: 10px;">Trạng thái</th>
      <th style="padding: 10px;">Số lỗi</th>
    </tr>
    ${Object.keys(results).map(function(sheetName) {
      var result = results[sheetName];
      return `<tr style="background-color: ${result.valid ? '#f0fdf4' : '#fee2e2'};">
                <td style="padding: 10px;"><strong>${sheetName}</strong></td>
                <td style="padding: 10px; color: ${result.valid ? '#166534' : '#dc2626'}; font-weight: bold;">${result.valid ? 'OK' : 'LỖI'}</td>
                <td style="padding: 10px; text-align: center; font-weight: bold; color: ${result.valid ? '#166534' : '#dc2626'};">${result.errorCount}</td>
              </tr>`;
    }).join('')}
  </table>
  <hr>
  <h3>🔍 Chi tiết lỗi cụ thể</h3>
  ${Object.keys(results).filter(sheetName => !results[sheetName].valid && results[sheetName].details && results[sheetName].details.length > 0).map(sheetName => {
    var result = results[sheetName];
    var detailList = result.details.map(err => {
      if (typeof err === 'object' && err !== null) {
        var position = err.row && err.col ? `Hàng ${err.row}, Cột ${String.fromCharCode(64 + parseInt(err.col))}` : 'Vị trí không xác định';
        var msg = err.message || 'Lỗi không xác định';
        var current = err.current !== undefined ? `<strong>Hiện tại:</strong> <span style="background:#fee2e2; padding:2px 6px; border-radius:4px;">${err.current}</span>` : '';
        var expected = err.expected !== undefined ? `<strong>Mong đợi:</strong> <span style="background:#fef3c7; padding:2px 6px; border-radius:4px;">${err.expected}</span>` : '';
        return `<li style="margin:8px 0; padding:10px; background:#fff5f5; border-left:4px solid #dc2626; border-radius:4px;">
                  <strong>${position}:</strong> ${msg}<br>
                  ${current ? current + '<br>' : ''}
                  ${expected}
                </li>`;
      } else {
        return `<li style="margin:8px 0;">${err}</li>`;
      }
    }).join('');
    return `<h4 style="color:#dc2626;">📌 Sheet: ${sheetName} (${result.errorCount} lỗi)</h4>
            <ul style="padding-left:20px;">${detailList}</ul>`;
  }).join('<hr>')}
  <hr>
  <p><strong>Vui lòng mở file Google Sheets để xem chi tiết đầy đủ và sửa lỗi (các ô lỗi đã được tô đỏ).</strong></p>
  <p>Trân trọng,<br><em>Script kiểm tra tự động</em></p>
`;
        var emailAddress = 'si.dt@sanshovn.com.vn,lanh.hv@sanshovn.com.vn,lien.dt@sanshovn.com.vn';
        MailApp.sendEmail({
          to: emailAddress,
          subject: emailSubject,
          htmlBody: emailBody
        });
      } catch (e) {
        Logger.log('Lỗi gửi email: ' + e.message);
      }
    }

    if (erroredSheet) {
      spreadsheet.setActiveSheet(erroredSheet);
      erroredSheet.setActiveRange(erroredSheet.getRange(7, 1));
    }
  } catch (e) {
    Logger.log('Lỗi chính trong checkAllData: ' + e.message);
    ui.alert('Lỗi nghiêm trọng', 'Đã xảy ra lỗi: ' + e.message + '\nVui lòng kiểm tra log để biết chi tiết.', ui.ButtonSet.OK);
  }
}



function checkSheet7040GA(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*A$/) !== null;
    var isValidB = (colB.includes('7189-7040-30') && colB.match(/.*A$/)) ||
                   (colB.includes('7172-1333') && colB.match(/.*D$/)) ||
                   (colB.includes('7137-3079') && colB.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng A', current: colA, expected: 'Kết thúc bằng A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7040-30A, 7172-1333D hoặc 7137-3079A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}
 

function checkSheet7040GB(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*B$/) !== null;
    var isValidB = (colB.includes('7189-7040-30') && colB.match(/.*F$/)) ||
                    (colB.includes('7189-7040') && colB.match(/.*B$/)) ||
                   (colB.includes('7189-7040-40') && colB.match(/.*B$/)) ||
                   (colB.includes('7172-1333') && colB.match(/.*D$/)) ||
                   (colB.includes('7137-3079') && colB.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng B', current: colA, expected: 'Kết thúc bằng B'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7040-30F, 7189-7040-40B, 7172-1333D hoặc 7137-3079A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheet8794B(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 7) return { valid: true, errorCount: 0, message: 'Không có dữ liệu', details: [] };

  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues(); // A, B, C
  var errors = [];
  var details = [];
  var valid = true;

  var lastLot = "";       
  var lastBEnd = "";      

  const FIXED_CODE_C = '7158-7116';  

  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var rawA = (data[i][0] || '').toString().trim();
    var rawB = (data[i][1] || '').toString().trim();
    var rawC = (data[i][2] || '').toString().trim();

    // === KIỂM TRA MỚI: ĐUÔI PHẦN SAU $ CỦA CỘT A PHẢI LÀ 'B' ===
    var currentLot = '';
    var endA = '';  // chữ cuối phần lot

    if (rawA.includes('$')) {
      var partsA = rawA.split('$');
      currentLot = partsA[1].trim();
      if (currentLot.length > 0) {
        endA = currentLot.slice(-1).toUpperCase();
      }
    }

    // Báo lỗi nếu:
    // - Không có $ → NG
    // - Có $ nhưng phần sau rỗng → NG
    // - Đuôi không phải 'B' → NG
    if (!rawA.includes('$') || currentLot === '' || endA !== 'B') {
      errors.push({row: row, col: 1});
      details.push({
        row: row, 
        col: 1, 
        message: 'Phần sau $ của cột A phải tồn tại và kết thúc bằng B',
        current: rawA,
        expected: 'ví dụ: ...$...B'
      });
      valid = false;
    }

    // === CỘT B (giữ nguyên) ===
    if (rawB !== '') {
      var currentBEnd = '';
      if (rawB.includes('$')) {
        currentBEnd = rawB.split('$')[1].slice(-1).toUpperCase();
      } else {
        currentBEnd = rawB.slice(-1).toUpperCase();
      }

      if (currentBEnd !== 'A' && currentBEnd !== 'B' && currentBEnd !== 'C') {
        errors.push({row: row, col: 2});
        details.push({row: row, col: 2, message: 'Đuôi B không phải A/B/C', current: currentBEnd});
        valid = false;
      } else if (lastLot && currentLot === lastLot && lastBEnd && currentBEnd !== lastBEnd) {
        errors.push({row: row, col: 2});
        details.push({row: row, col: 2, message: 'Xung đột đuôi B trong cùng lot', current: currentBEnd, expected: lastBEnd});
        valid = false;
      }

      lastLot = currentLot;
      lastBEnd = currentBEnd;
    }

    // === CỘT C (giữ nguyên) ===
    if (rawC !== '') {
      var codeC = '';
      var endC = '';
      if (rawC.includes('$')) {
        var partsC = rawC.split('$');
        codeC = partsC[0].trim();
        endC = partsC[1].slice(-1).toUpperCase();
      } else {
        codeC = rawC;
        endC = rawC.slice(-1).toUpperCase();
      }

      if (!codeC.includes(FIXED_CODE_C) || (endC !== 'A' && endC !== 'B')) {
        errors.push({row: row, col: 3});
        details.push({row: row, col: 3, message: 'Mã C không hợp lệ hoặc đuôi không phải A/B', current: rawC});
        valid = false;
      }
    }

    // === QUY TẮC PHỤ THUỘC B → C (giữ nguyên) ===
    if (rawB !== '' && rawC !== '') {
      var bEnd = rawB.includes('$') ? rawB.split('$')[1].slice(-1).toUpperCase() : rawB.slice(-1).toUpperCase();
      var cEnd = rawC.includes('$') ? rawC.split('$')[1].slice(-1).toUpperCase() : rawC.slice(-1).toUpperCase();

      if ((bEnd === 'A' || bEnd === 'B') && cEnd === 'B') {
        errors.push({row: row, col: 3});
        details.push({row: row, col: 3, message: 'Đuôi C không được B khi đuôi B là A hoặc B', current: cEnd, expected: 'A'});
        valid = false;
      }
      if (bEnd === 'C' && cEnd === 'A') {
        errors.push({row: row, col: 3});
        details.push({row: row, col: 3, message: 'Đuôi C không được A khi đuôi B là C', current: cEnd, expected: 'B'});
        valid = false;
      }
    }
  }

  // Tô đỏ tất cả lỗi (bao gồm cột A nếu có)
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }

  return { 
    valid: valid, 
    errorCount: errors.length, 
    message: errors.length ? 'Phát hiện ' + errors.length + ' lỗi' : 'Không có lỗi',
    details: details 
  };
}
function checkSheet7040GE(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*E$/) !== null;
    var isValidB = (colB.includes('7189-7040-30') && colB.match(/.*E$/)) ||
                   (colB.includes('7172-1333') && colB.match(/.*C$/)) ||
                   (colB.includes('7137-3079') && colB.match(/.*C$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng E', current: colA, expected: 'Kết thúc bằng E'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7040-30E, 7172-1333C hoặc 7137-3079C'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVS001A(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*A$/) !== null;
    var isValidB = (colB.includes('7186-1896') && colB.match(/.*A$/)) ||
                   (colB.includes('7172-0221') && colB.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng A', current: colA, expected: 'Kết thúc bằng A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7186-1896A hoặc 7172-0221A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVS001B(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*B$/) !== null;
    var isValidB = (colB.includes('7186-1896') && colB.match(/.*B$/)) ||
                   (colB.includes('7172-0221') && colB.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng B', current: colA, expected: 'Kết thúc bằng B'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7186-1896B hoặc 7172-0221A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVS001C(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var isValidA = colA.match(/.*C$/) !== null;
    var isValidB = (colB.includes('7186-1896') && colB.match(/.*C$/)) ||
                   (colB.includes('7172-0221') && colB.match(/.*B$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng C', current: colA, expected: 'Kết thúc bằng C'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7186-1896C hoặc 7172-0221B'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVA019A(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 7) return { valid: true, errorCount: 0, message: 'Không có dữ liệu', details: [] };

  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues();
  var errors = [];
  var details = [];
  var valid = true;

  // Thêm: Theo dõi lot thực tế (sau $) → chữ cuối B và C đầu tiên gặp + danh sách row
  var lotRealMap = {}; // key: lotReal (ví dụ "251209A"), value: { bEnd: 'A', cEnd: 'B', rows: [...] }

  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    var colC = data[i][2] ? removeMinusOne(data[i][2]) : '';

    if (colA === "" && colB === "" && colC === "") continue;

    // === LOGIC GỐC CỦA VA019A - GIỮ NGUYÊN HOÀN TOÀN KHÔNG THAY ĐỔI ===
    var isValidA = colA !== '' && colA.match(/.*A$/) !== null;

    var isValidB = colB === '' ||
                   (colB.includes('7189-3342-40') && colB.match(/.*A$/)) ||
                   (colB.includes('7189-3342-90') && colB.match(/.*A$/)) ||
                   (colB.includes('7189-3342-30') && colB.match(/.*C$/));

    var isValidC = colC === '' ||
                   (colC.includes('7158-6425') && (colC.match(/.*A$/) || colC.match(/.*B$/)));

    // Quy tắc tổ hợp gốc
    var isValidRow = false;
    if (colA !== '' && colB !== '' && colC !== '') {
      var colCEnd = colC.slice(-1).toUpperCase();
      isValidRow =
        (colA.includes('7289-3342-40') && colA.match(/.*A$/) &&
         colB.includes('7189-3342-40') && colB.match(/.*A$/) &&
         colC.includes('7158-6425') && colCEnd === 'B') ||
        (colA.includes('7289-3342-90') && colA.match(/.*A$/) &&
         colB.includes('7189-3342-90') && colB.match(/.*A$/) &&
         colC.includes('7158-6425') && colCEnd === 'B') ||
        (colA.includes('7289-3342-30') && colA.match(/.*A$/) &&
         colB.includes('7189-3342-30') && colB.match(/.*C$/) &&
         colC.includes('7158-6425') && colCEnd === 'B') ||
        (colA.includes('7289-3342-30') && colA.match(/.*A$/) &&
         colB.includes('7189-3342-30') && colB.match(/.*C$/) &&
         colC.includes('7158-6425') && colCEnd === 'A');
    }

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã A không hợp lệ', current: colA, expected: 'Kết thúc bằng A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã B không hợp lệ', current: colB, expected: '7189-3342-40A, 7189-3342-90A hoặc 7189-3342-30C'});
      valid = false;
    }
    if (!isValidC) {
      errors.push({row: row, col: 3});
      details.push({row: row, col: 3, message: 'Mã C không hợp lệ', current: colC, expected: '7158-6425 với đuôi A hoặc B'});
      valid = false;
    }
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({row: row, col: null, message: 'Tổ hợp A-B-C không hợp lệ'});
      valid = false;
    }

        // === QUY TẮC MỚI: KIỂM SOÁT CHỮ CUỐI CỦA CẢ CỘT B VÀ CỘT C KHI CÙNG LOT SAU $ ===
    if (colA && colA.includes('$') && colA.match(/.*A$/) !== null) {
      var parts = colA.split('$');
      var lotReal = parts[parts.length - 1].trim(); // lot thực tế sau $ cuối cùng

      var currentBEnd = colB ? colB.slice(-1).toUpperCase() : '';
      var currentCEnd = colC ? colC.slice(-1).toUpperCase() : '';

      if (lotRealMap[lotReal] === undefined) {
        // Lần đầu gặp lot thực tế này → ghi nhớ chữ cuối B và C + lưu dòng đầu tiên
        lotRealMap[lotReal] = {
          bEnd: currentBEnd,
          cEnd: currentCEnd,
          rows: [row]  // vẫn lưu danh sách rows để tiện sau này nếu cần
        };
      } else {
        // Đã gặp trước đó
        lotRealMap[lotReal].rows.push(row);

        // Kiểm tra xung đột (chỉ khi ô không trống)
        var conflictB = (currentBEnd !== '' && lotRealMap[lotReal].bEnd !== '' && lotRealMap[lotReal].bEnd !== currentBEnd);
        var conflictC = (currentCEnd !== '' && lotRealMap[lotReal].cEnd !== '' && lotRealMap[lotReal].cEnd !== currentCEnd);

        if (conflictB || conflictC) {
          // CHỈ TÔ ĐỎ DÒNG HIỆN TẠI (dòng gây xung đột)
          errors.push({row: row, col: 1});
          errors.push({row: row, col: 2});
          errors.push({row: row, col: 3});

          var conflictMsg = 'Xung đột đuôi trong cùng lot ' + lotReal;
          if (conflictB) conflictMsg += ' (B: ' + currentBEnd + ' ≠ ' + lotRealMap[lotReal].bEnd + ')';
          if (conflictC) conflictMsg += ' (C: ' + currentCEnd + ' ≠ ' + lotRealMap[lotReal].cEnd + ')';

          details.push({row: row, col: null, message: conflictMsg});

          valid = false;

          // Tùy chọn: cũng có thể đánh dấu dòng đầu tiên của lot để người dùng biết "chuẩn là gì"
          // Nhưng theo yêu cầu "chỉ bôi đỏ dòng lỗi" → không tô dòng đầu
        }
      }
    
    }
  }

  // Tô đỏ lỗi
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }

  return { 
    valid: valid, 
    errorCount: errors.length, 
    message: errors.length ? 'Phát hiện ' + errors.length + ' lỗi' : 'Không có lỗi',
    details: details 
  };
}

function checkSheetVA019B(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 7) return { valid: true, errorCount: 0, message: 'Không có dữ liệu', details: [] };

  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  var lotRealMap = {};

  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    var colC = data[i][2] ? removeMinusOne(data[i][2]) : '';

    if (colA === "" && colB === "" && colC === "") continue;

    var colBEnd = colB ? String(colB).slice(-1).toUpperCase() : '';
    var colCEnd = colC ? String(colC).slice(-1).toUpperCase() : '';

    var prefixA = String(colA).includes('$') ? String(colA).split('$')[0].trim() : String(colA).trim();
    var prefixB = String(colB).includes('$') ? String(colB).split('$')[0].trim() : String(colB).trim();
    var codeA = prefixA.replace(/^7289-/, '');
    var codeB = prefixB.replace(/^7189-/, '');
    var isSamePrefix = (codeA === codeB);
    var is90 = (codeA === '3342-90');

    var isValidA = /B$/i.test(colA);
    var isValidB = String(colB).includes('7189-3342') && (colBEnd === 'B' || colBEnd === 'C' || colBEnd === 'F');
    var isValidC = String(colC).includes('7158-6425') && (colCEnd === 'C' || colCEnd === 'E');

    // Cặp đuôi hợp lệ: B-C, C-C, F-E  |  F-C / B-E / C-E = NG
    var pairOK =
      (colBEnd === 'B' && colCEnd === 'C') ||
      (colBEnd === 'C' && colCEnd === 'C') ||
      (colBEnd === 'F' && colCEnd === 'E');

    // -90 chỉ được F-E
    var combo90OK = !is90 || (colBEnd === 'F' && colCEnd === 'E');

    var isValidRow = isSamePrefix && isValidA && pairOK && combo90OK;

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã A không hợp lệ', current: colA, expected: 'Kết thúc bằng B'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã B không hợp lệ', current: colB, expected: '7189-3342 đuôi B / C / F'});
      valid = false;
    }
    if (!isValidC) {
      errors.push({row: row, col: 3});
      details.push({row: row, col: 3, message: 'Mã C không hợp lệ', current: colC, expected: '7158-6425 đuôi C hoặc E'});
      valid = false;
    }
    if (!isSamePrefix) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      details.push({
        row: row, col: null,
        message: 'Phần trước $ của A và B không khớp',
        current: prefixA + ' | ' + prefixB,
        expected: '3342 với 3342, hoặc 3342-90 với 3342-90'
      });
      valid = false;
    }
    if (colBEnd === 'F' && colCEnd === 'C') {
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({
        row: row, col: null,
        message: 'GF + GC không hợp lệ',
        current: 'B=' + colBEnd + ' / C=' + colCEnd,
        expected: 'GF phải đi với GE (F + E)'
      });
      valid = false;
    }
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({
        row: row, col: null,
        message: 'Tổ hợp A-B-C không hợp lệ',
        current: prefixA + ' | B=' + colBEnd + ' | C=' + colCEnd,
        expected: is90 ? '3342-90 chỉ được F + E' : '3342: B+C hoặc C+C hoặc F+E (không F+C)'
      });
      valid = false;
    }

    if (colA && String(colA).includes('$') && isValidA) {
      var lotReal = String(colA).split('$').pop().trim();
      if (lotRealMap[lotReal] === undefined) {
        lotRealMap[lotReal] = { bEnd: colBEnd, cEnd: colCEnd, rows: [row] };
      } else {
        lotRealMap[lotReal].rows.push(row);
        var conflictB = (colBEnd !== '' && lotRealMap[lotReal].bEnd !== '' && lotRealMap[lotReal].bEnd !== colBEnd);
        var conflictC = (colCEnd !== '' && lotRealMap[lotReal].cEnd !== '' && lotRealMap[lotReal].cEnd !== colCEnd);
        if (conflictB || conflictC) {
          errors.push({row: row, col: 1});
          errors.push({row: row, col: 2});
          errors.push({row: row, col: 3});
          var conflictMsg = 'Xung đột đuôi trong cùng lot ' + lotReal;
          if (conflictB) conflictMsg += ' (B: ' + colBEnd + ' ≠ ' + lotRealMap[lotReal].bEnd + ')';
          if (conflictC) conflictMsg += ' (C: ' + colCEnd + ' ≠ ' + lotRealMap[lotReal].cEnd + ')';
          details.push({row: row, col: null, message: conflictMsg});
          valid = false;
        }
      }
    }
  }

  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }

  return {
    valid: valid,
    errorCount: errors.length,
    message: errors.length ? 'Phát hiện ' + errors.length + ' lỗi' : 'Không có lỗi',
    details: details
  };
}
function checkSheetVA019C(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 7) return { valid: true, errorCount: 0, message: 'Không có dữ liệu', details: [] };

  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues();
  var errors = [];
  var details = [];
  var valid = true;

  // Thêm: Theo dõi lot thực tế (sau $) → chữ cuối B và C đầu tiên gặp + danh sách row
  var lotRealMap = {}; // key: lotReal (ví dụ "251209C"), value: { bEnd: 'D', cEnd: 'A', rows: [...] }

  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    var colC = data[i][2] ? removeMinusOne(data[i][2]) : '';

    if (colA === "" && colB === "" && colC === "") continue;

    // === LOGIC GỐC CỦA VA019C - GIỮ NGUYÊN HOÀN TOÀN KHÔNG THAY ĐỔI ===
    var isValidA = colA.match(/.*C$/) !== null;
    var isValidB = (colB.includes('7189-3342-30') && colB.match(/.*D$/)) ||
                   (colB.includes('7189-3342-30') && colB.match(/.*E$/));
    var isValidC = (colC.includes('7158-6425') && colC.match(/.*A$/)) ||
                   (colC.includes('7158-6425') && colC.match(/.*D$/));

    var colCEnd = colC ? colC.slice(-1).toUpperCase() : '';

    var isValidRow = (colA.includes('7289-3342-30') && colA.match(/.*C$/) &&
                      colB.includes('7189-3342-30') && colB.match(/.*D$/) &&
                      colC.includes('7158-6425') && colCEnd === 'D') ||
                     (colA.includes('7289-3342-30') && colA.match(/.*C$/) &&
                      colB.includes('7189-3342-30') && colB.match(/.*D$/) &&
                      colC.includes('7158-6425') && colCEnd === 'A') ||
                     (colA.includes('7289-3342-30') && colA.match(/.*C$/) &&
                      colB.includes('7189-3342-30') && colB.match(/.*E$/) &&
                      colC.includes('7158-6425') && colCEnd === 'D');

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã A không hợp lệ', current: colA, expected: 'Kết thúc bằng C'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã B không hợp lệ', current: colB, expected: '7189-3342-30D hoặc 7189-3342-30E'});
      valid = false;
    }
    if (!isValidC) {
      errors.push({row: row, col: 3});
      details.push({row: row, col: 3, message: 'Mã C không hợp lệ', current: colC, expected: '7158-6425 với đuôi A hoặc D'});
      valid = false;
    }
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({row: row, col: null, message: 'Tổ hợp A-B-C không hợp lệ'});
      valid = false;
    }

       // === QUY TẮC MỚI: KIỂM SOÁT CHỮ CUỐI CỦA CẢ CỘT B VÀ CỘT C KHI CÙNG LOT SAU $ ===
    if (colA && colA.includes('$') && colA.match(/.*C$/) !== null) {
      var parts = colA.split('$');
      var lotReal = parts[parts.length - 1].trim();

      var currentBEnd = colB ? colB.slice(-1).toUpperCase() : '';
      var currentCEnd = colC ? colC.slice(-1).toUpperCase() : '';

      if (lotRealMap[lotReal] === undefined) {
        lotRealMap[lotReal] = {
          bEnd: currentBEnd,
          cEnd: currentCEnd,
          rows: [row]
        };
      } else {
        lotRealMap[lotReal].rows.push(row);

        var conflictB = (currentBEnd !== '' && lotRealMap[lotReal].bEnd !== '' && lotRealMap[lotReal].bEnd !== currentBEnd);
        var conflictC = (currentCEnd !== '' && lotRealMap[lotReal].cEnd !== '' && lotRealMap[lotReal].cEnd !== currentCEnd);

        if (conflictB || conflictC) {
          errors.push({row: row, col: 1});
          errors.push({row: row, col: 2});
          errors.push({row: row, col: 3});

          var conflictMsg = 'Xung đột đuôi trong cùng lot ' + lotReal;
          if (conflictB) conflictMsg += ' (B: ' + currentBEnd + ' ≠ ' + lotRealMap[lotReal].bEnd + ')';
          if (conflictC) conflictMsg += ' (C: ' + currentCEnd + ' ≠ ' + lotRealMap[lotReal].cEnd + ')';

          details.push({row: row, col: null, message: conflictMsg});

          valid = false;
        }
      }
    
    }
  }

  // Tô đỏ lỗi
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }

  return { 
    valid: valid, 
    errorCount: errors.length, 
    message: errors.length ? 'Phát hiện ' + errors.length + ' lỗi' : 'Không có lỗi',
    details: details 
  };
}

function checkSheetVA003(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 3).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    var colC = data[i][2] ? removeMinusOne(data[i][2]) : '';
    if (colA === "" && colB === "" && colC === "") continue;
    var isValidA = (colA.includes('7283-4578') && colA.match(/.*A$/)) ||
                   (colA.includes('7298-8296') && colA.match(/.*A$/));
    var isValidB = (colB.includes('7183-4578') && colB.match(/.*A$/)) ||
                   (colB.includes('7198-8296') && colB.match(/.*A$/));
    var isValidC = colC.includes('7158-6425') && colC.match(/.*A$/);
    var isValidRow = (colA.includes('7283-4578') && colA.match(/.*A$/) &&
                      colB.includes('7183-4578') && colB.match(/.*A$/) &&
                      colC.includes('7158-6425') && colC.match(/.*A$/)) ||
                     (colA.includes('7298-8296') && colA.match(/.*A$/) &&
                      colB.includes('7198-8296') && colB.match(/.*A$/) &&
                      colC.includes('7158-6425') && colC.match(/.*A$/));
    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã A không hợp lệ', current: colA, expected: '7283-4578A hoặc 7298-8296A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã B không hợp lệ', current: colB, expected: '7183-4578A hoặc 7198-8296A'});
      valid = false;
    }
    if (!isValidC) {
      errors.push({row: row, col: 3});
      details.push({row: row, col: 3, message: 'Mã C không hợp lệ', current: colC, expected: '7158-6425A'});
      valid = false;
    }
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      errors.push({row: row, col: 3});
      details.push({row: row, col: null, message: 'Tổ hợp A-B-C không hợp lệ'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheetVS051(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;
    var colAEnd = colA.slice(-1).toUpperCase();
    var colBEnd = colB.slice(-1).toUpperCase();
    var isValidRow = (colA.includes('7286-7859') && colAEnd === 'A' && colB.includes('7186-7859') && colBEnd === 'A') ||
                     (colA.includes('7286-7859') && colAEnd === 'B' && colB.includes('7186-7859') && colBEnd === 'B');
    if (!isValidRow) {
      errors.push({row: row, col: 1});
      errors.push({row: row, col: 2});
      details.push({row: row, col: null, message: 'Tổ hợp A-B không hợp lệ', current: colA + ' & ' + colB, expected: '7286-7859A & 7186-7859A hoặc 7286-7859B & 7186-7859B'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function checkSheet7041GA(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;

    // Cột A phải kết thúc bằng GA 
    var isValidA = colA.match(/.*A$/i) !== null;

    // Cột B chỉ được 3 mã sau và đúng đuôi GA 
    var isValidB = (colB.includes('7189-7041-30') && colB.match(/.*A$/i)) ||
                   (colB.includes('7172-1334') && colB.match(/.*A$/i)) ||
                   (colB.includes('7137-3080') && colB.match(/.*A$/i));

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng A', current: colA, expected: 'Kết thúc bằng A'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7041-30A, 7172-1334A hoặc 7137-3080A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}
function checkSheet7041GB(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(7, 1, lastRow - 6, 2).getValues();
  var errors = [];
  var details = [];
  var valid = true;
  for (var i = 0; i < data.length; i++) {
    var row = i + 7;
    var colA = removeMinusOne(data[i][0]);
    var colB = removeMinusOne(data[i][1]);
    if (colA === "" && colB === "") continue;

    // Cột A phải kết thúc bằng B
    var isValidA = colA.match(/.*B$/i) !== null;

    // Cột B chỉ được 3 mã sau và đúng đuôi 
    var isValidB = (colB.includes('7189-7041-30') && colB.match(/.*B$/i)) ||
                   (colB.includes('7172-1334') && colB.match(/.*A$/i)) ||
                   (colB.includes('7137-3080') && colB.match(/.*A$/i));

    if (!isValidA) {
      errors.push({row: row, col: 1});
      details.push({row: row, col: 1, message: 'Mã không kết thúc bằng B', current: colA, expected: 'Kết thúc bằng B'});
      valid = false;
    }
    if (!isValidB) {
      errors.push({row: row, col: 2});
      details.push({row: row, col: 2, message: 'Mã không hợp lệ', current: colB, expected: '7189-7041-30B, 7172-1334A hoặc 7137-3080A'});
      valid = false;
    }
  }
  if (errors.length) {
    var errorRanges = errors.map(e => String.fromCharCode(64 + e.col) + e.row);
    sheet.getRangeList(errorRanges).setBackground('#ff0000');
  }
  return { valid: valid, errorCount: errors.length, message: errors.length ? 'Phát hiện lỗi ở các hàng: ' + [...new Set(errors.map(e => e.row))].join(', ') : 'Không có lỗi', details: details };
}

function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('checkAllData')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// HỆ THỐNG: KIỂM TRA MÃ KHUÔN & TEM NHÃN – BẢN HOÀN HẢO CUỐI CÙNG
// LB-VA019A: CỘT A CỐ ĐỊNH LB-VA019A-7289-3342-A | CỘT B CHỈ 3 MÃ | CỘT C PHẢI A
// ═══════════════════════════════════════════════════════════════════════
function checkLabelQuick() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const RULES = [
    { sheet: 'LB-VA019A', fixedA: 'LB-VA019A-7289-3342-A', codesB: ['7289-3342-30', '7289-3342-40', '7289-3342-90'], endC: 'A' },
    { sheet: 'LB-VA019B', fixedA: 'LB-VA019B-7289-3342-B', codesB: '7289-3342', endC: 'B' },
    { sheet: 'LB-VA019C', fixedA: 'LB-VA019C-7289-3342-C', codesB: '7289-3342-30', endC: 'C' },
    { sheet: 'LB-VA015A', fixedA: 'LB-VA015A-7289-7040-30-A', codesB: '7289-7040-30', endC: 'A' },
    { sheet: 'LB-VA015B', fixedA: 'LB-VA015B-7289-7040-30-B', codesB: ['7289-7040', '7289-7040-30', '7289-7040-40'], endC: 'B' },
    { sheet: 'LB-VA015E', fixedA: 'LB-VA015E-7289-7040-30-E', codesB: '7289-7040-30', endC: 'E' },
    { sheet: 'LB-VA016A', fixedA: 'LB-VA016A-7289-7041-30-A', codesB: '7289-7041-30', endC: 'A' },
    { sheet: 'LB-VA016B', fixedA: 'LB-VA016B-7289-7041-30-B', codesB: '7289-7041-30', endC: 'B' },
    { sheet: 'LB-VA001B', fixedA: 'LB-VA001B-7283-8794-70-B', codesB: '7283-8794-70', endC: 'B' },
    { sheet: 'LB-VA006', fixedA: 'LB-VA006-7283-0651-70-A', codesB: '7283-0651-70', endC: 'A' },
    { sheet: 'LB-VA018', fixedA: 'LB-VA018-7289-8063-30-A', codesB: '7289-8063-30', endC: 'B' },
    { sheet: 'LB-VS001A', fixedA: 'LB-VS001A-7286-1896-A', codesB: '7286-1896', endC: 'A' },
    { sheet: 'LB-VS001B', fixedA: 'LB-VS001B-7286-1896-B', codesB: '7286-1896', endC: 'B' },
    { sheet: 'LB-VS001C', fixedA: 'LB-VS001C-7286-1896-C', codesB: '7286-1896', endC: 'C' },
    { sheet: 'LB-VS002', fixedA: 'LB-VS002-7286-1902-A', codesB: '7286-1902', endC: 'A' },
    { sheet: 'LB-VS004', fixedA: 'LB-VS004-7286-1903-A', codesB: '7286-1903', endC: 'A' },
    { sheet: 'LB-VS012C', fixedA: 'LB-VS012C-7287-8821-C', codesB: '7287-8821', endC: 'C' },
    { sheet: 'LB-VS051B', fixedA: 'LB-VS051B-7286-7859-B', codesB: '7286-7859', endC: 'B' }
  ];

  let totalErrors = 0;
  const errorDetails = [];       // Chi tiết lỗi để gửi email
  const sheetStatus = [];        // Chỉ để hiển thị OK/LỖI ngắn gọn trong log cũ (nếu cần)

  RULES.forEach(r => {
    const sheet = ss.getSheetByName(r.sheet);
    if (!sheet) {
      errorDetails.push(`<li><strong>Không tìm thấy sheet:</strong> ${r.sheet}</li>`);
      sheetStatus.push(`Không tìm thấy sheet: ${r.sheet}`);
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 7) {
      errorDetails.push(`<li><strong>Không có dữ liệu:</strong> ${r.sheet}</li>`);
      sheetStatus.push(`Không có dữ liệu: ${r.sheet}`);
      return;
    }

    const startRow = 7;
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
    sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).setBackground(null);

    let sheetHasError = false;

    data.forEach((row, i) => {
      const rowNum = startRow + i;
      const a = (row[0] || '').toString().trim();
      const b = (row[1] || '').toString().trim();
      const c = (row[2] || '').toString().trim();

      // Kiểm tra cột A
      if (a !== r.fixedA) {
        sheet.getRange(rowNum, 1).setBackground('#ff5252');
        errorDetails.push(`<li><strong>${r.sheet} - Dòng ${rowNum}</strong>: Cột A sai → thực tế "<em>${a}</em>", mong đợi "<em>${r.fixedA}</em>"</li>`);
        sheetHasError = true;
        totalErrors++;
      }

      // Kiểm tra cột B
      let bValid = false;
      if (Array.isArray(r.codesB)) {
        bValid = r.codesB.includes(b);
      } else {
        bValid = (b === r.codesB);
      }

      if (b && !bValid) {
        sheet.getRange(rowNum, 2).setBackground('#ff5252');
        const expectedB = Array.isArray(r.codesB) ? r.codesB.join(', ') : r.codesB;
        errorDetails.push(`<li><strong>${r.sheet} - Dòng ${rowNum}</strong>: Cột B sai → thực tế "<em>${b}</em>", mong đợi một trong "<em>${expectedB}</em>"</li>`);
        sheetHasError = true;
        totalErrors++;
      }

      // Kiểm tra cột C
      if (c && !c.toUpperCase().endsWith(r.endC)) {
        sheet.getRange(rowNum, 3).setBackground('#ff5252');
        errorDetails.push(`<li><strong>${r.sheet} - Dòng ${rowNum}</strong>: Cột C sai → thực tế "<em>${c}</em>", phải kết thúc bằng "<em>${r.endC}</em>"</li>`);
        sheetHasError = true;
        totalErrors++;
      }
    });

    const status = sheetHasError ? `LỖI: ${r.sheet}` : `OK: ${r.sheet}`;
    sheetStatus.push(status);
  });

  const timeStr = Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy');

  // ================== HIỂN THỊ KẾT QUẢ ==================
  if (totalErrors === 0) {
    ui.showModalDialog(HtmlService.createHtmlOutput(`
      <div style="text-align:center;padding:80px;background:#e8f5e8;font-family:Arial;">
        <h1 style="color:#2e7d32">HOÀN HẢO!</h1>
        <div style="font-size:120px">✅</div>
        <h2>Tất cả dữ liệu đều chuẩn</h2>
        <button onclick="google.script.host.close()"
          style="padding:18px 60px;font-size:22px;background:#4caf50;color:white;border:none;border-radius:10px">
          OK
        </button>
      </div>
    `).setWidth(520).setHeight(420), 'KIỂM TRA');

    return;
  }

  // ================== GỬI EMAIL CHI TIẾT KHI CÓ LỖI ==================
  MailApp.sendEmail({
    to: 'si.dt@sanshovn.com.vn,lanh.hv@sanshovn.com.vn,lien.dt@sanshovn.com.vn',
    subject: `LỖI MÃ KHUÔN & TEM NHÃN – ${totalErrors} lỗi – ${timeStr}`,
    htmlBody: `
      <h2 style="color:#c62828;">Phát hiện ${totalErrors} lỗi trong file kiểm tra nhãn</h2>
      <p><strong>Thời gian kiểm tra:</strong> ${timeStr}</p>
      <hr>
      <h3>Chi tiết lỗi:</h3>
      <ol>${errorDetails.join('')}</ol>
      <hr>
      <p>Vui lòng mở file Google Sheets và sửa các ô được tô đỏ.</p>
      <p>Trân trọng,<br>Script tự động kiểm tra</p>
    `
  });

  ui.showModalDialog(HtmlService.createHtmlOutput(`
    <div style="text-align:center;padding:80px;background:#ffebee;font-family:Arial;">
      <h1 style="color:#c62828">Có ${totalErrors} Lỗi<br>đã gửi thông tin chi tiết qua email...</h1>
      <div style="font-size:120px">❌</div>
      <button onclick="google.script.run.jumpToError()"
        style="padding:18px 60px;font-size:22px;background:#e53935;color:white;border:none;border-radius:10px">
        Xác nhận ngay !!
      </button>
    </div>
  `).setWidth(520).setHeight(420), 'CẢNH BÁO');
}
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KIỂM TRA MÃ KHUÔN & TEM NHÃN')
    .addItem('Kiểm tra toàn bộ từ dòng 7', 'checkLabelQuick')
    .addToUi();
}

// ==================== RESET HOÀN HẢO – KHÔNG MẤT DÒNG 7 NỮA ====================
function resetAllSheetsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'XÁC NHẬN RESET',
    'Xóa toàn bộ dữ liệu từ dòng 7 trở xuống trong TẤT CẢ các sheet?\n\n' +
    'Dòng dữ liệu mới nhất sẽ được kéo lên đúng dòng 7.\nKHÔNG THỂ HOÀN TÁC!',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Đã hủy', 'Không thực hiện reset.', ui.ButtonSet.OK);
    return;
  }

  ui.showModalDialog(HtmlService.createHtmlOutput(`
    <div style="text-align:center;padding:80px;background:#fff3e0;font-family:'Roboto',sans-serif;">
      <h2 style="font-size:32px;color:#ef6c00;">Đang reset dữ liệu...</h2>
      <p style="font-size:20px;color:#e65100;">Đang xử lý toàn bộ sheet</p>
      <div style="font-size:90px;">🗑️➡️📋</div>
    </div>
  `).setWidth(520).setHeight(420), 'ĐANG XỬ LÝ');

  const sheets = ss.getSheets();
  let processedCount = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === 'CheckLog') return; // Giữ lịch sử

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    // === 1. Xóa sạch nội dung + màu nền từ dòng 8 trở xuống (GIỮ NGUYÊN DÒNG 7) ===
    const maxRows = sheet.getMaxRows();
    if (maxRows > 7) {
      sheet.getRange(8, 1, maxRows - 7, lastCol)
           .clearContent()
           .setBackground(null);
    }

    // === 2. Tìm khối dữ liệu cuối cùng (dòng mới nhất) ===
    const lastRowAfterClear = sheet.getLastRow();
    if (lastRowAfterClear < 8) {
      processedCount++;
      return; // Không có dữ liệu cần di chuyển
    }

    const values = sheet.getRange(8, 1, lastRowAfterClear - 7, lastCol).getValues();

    // Tìm dòng cuối cùng có dữ liệu
    let lastDataIndex = -1;
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i].some(cell => cell !== '' && cell !== null && cell !== undefined)) {
        lastDataIndex = i;
        break;
      }
    }
    if (lastDataIndex === -1) {
      processedCount++;
      return;
    }

    // Tìm đầu khối dữ liệu liên tục
    let firstDataIndex = lastDataIndex;
    for (let i = lastDataIndex - 1; i >= 0; i--) {
      const allEmpty = values[i].every(cell => cell === '' || cell === null || cell === undefined);
      if (allEmpty) break;
      firstDataIndex = i;
    }

    const blockHeight = lastDataIndex - firstDataIndex + 1;
    const sourceRow = 8 + firstDataIndex; // Dòng thực tế trong sheet

    // === 3. Di chuyển lên dòng 7 ===
    const sourceRange = sheet.getRange(sourceRow, 1, blockHeight, lastCol);
    const targetRange = sheet.getRange(7, 1, blockHeight, lastCol);
    sourceRange.moveTo(targetRange);

    // === 4. Xóa vùng cũ (an toàn tuyệt đối) ===
    if (sourceRow + blockHeight - 1 <= maxRows) {
      sheet.deleteRows(sourceRow, blockHeight);
    }

    processedCount++;
  });

  ui.showModalDialog(HtmlService.createHtmlOutput(`
    <div style="text-align:center;padding:80px;background:#e8f5e8;font-family:'Roboto',sans-serif;">
      <h1 style="font-size:48px;color:#2e7d32;">HOÀN TẤT!</h1>
      <p style="font-size:24px;color:#1b5e20;">
        Đã reset thành công!<br>
        <b>Dòng mới nhất đã được kéo lên dòng 7</b><br>
        Xử lý ${processedCount} sheet.
      </p>
      <div style="font-size:100px;">✅</div>
      <button onclick="google.script.host.close()" style="margin-top:30px;padding:16px 50px;font-size:22px;background:#4caf50;color:white;border:none;border-radius:12px;cursor:pointer;">
        Đóng
      </button>
    </div>
  `).setWidth(560).setHeight(520), 'RESET THÀNH CÔNG');
}
// ==================== DANH SÁCH EMAIL ĐƯỢC PHÉP DÙNG TOOL VÀ CHỈNH SỬA CODE ====================
const ALLOWED_EMAILS = [
  'sitien353@gmail.com.vn',
  // Thêm email khác nếu cần
];

// ==================== ẨN MENU VỚI NGƯỜI KHÔNG ĐƯỢC PHÉP ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const userEmail = Session.getActiveUser().getEmail();

  if (ALLOWED_EMAILS.includes(userEmail) || userEmail === '') {
    ui.createMenu('🔧 TOOL KIỂM TRA MÃ KHUÔN')
      .addItem('🔍 Kiểm tra cấu tạo mã khuôn linh kiện', 'checkAllData')
      .addItem('🏷️ Kiểm tra mã khuôn & tem nhãn (dòng 7)', 'checkLabelQuick')
      .addSeparator()
      .addItem('♻️ Reset dữ liệu (kéo dòng mới lên dòng 7)', 'resetAllSheetsData')
      .addToUi();
  }
}

// ==================== CHẶN MỞ APPS SCRIPT EDITOR ====================
function onOpenScriptEditor() {
  const userEmail = Session.getActiveUser().getEmail();
  
  if (!ALLOWED_EMAILS.includes(userEmail) && userEmail !== '') {
    throw new Error('🚫 Bạn không có quyền truy cập hoặc chỉnh sửa script này.\nLiên hệ admin để được hỗ trợ.');
  }
}

//update speaker 28/08/2026 + hiển thị Số xe trong lịch sử
function getSheetCheckMap() {
  return {
    'VA001-B': checkSheet8794B,
    'VA015-A': checkSheet7040GA,
    'VA015-B': checkSheet7040GB,
    'VA015-E': checkSheet7040GE,
    'VS001-A': checkSheetVS001A,
    'VS001-B': checkSheetVS001B,
    'VS001-C': checkSheetVS001C,
    'VA019-A': checkSheetVA019A,
    'VA019-B': checkSheetVA019B,
    'VA019-C': checkSheetVA019C,
    'VA003': checkSheetVA003,
    'VS051': checkSheetVS051,
    'VA016-A': checkSheet7041GA,
    'VA016-B': checkSheet7041GB
  };
}

function getLabelRules() {
  return [
    { sheet: 'LB-VA019A', fixedA: 'LB-VA019A-7289-3342-A', codesB: ['7289-3342-30', '7289-3342-40', '7289-3342-90'], endC: 'A' },
    { sheet: 'LB-VA019B', fixedA: 'LB-VA019B-7289-3342-B', codesB: '7289-3342', endC: 'B' },
    { sheet: 'LB-VA019C', fixedA: 'LB-VA019C-7289-3342-C', codesB: '7289-3342-30', endC: 'C' },
    { sheet: 'LB-VA015A', fixedA: 'LB-VA015A-7289-7040-30-A', codesB: '7289-7040-30', endC: 'A' },
    { sheet: 'LB-VA015B', fixedA: 'LB-VA015B-7289-7040-30-B', codesB: ['7289-7040', '7289-7040-30', '7289-7040-40'], endC: 'B' },
    { sheet: 'LB-VA015E', fixedA: 'LB-VA015E-7289-7040-30-E', codesB: '7289-7040-30', endC: 'E' },
    { sheet: 'LB-VA016A', fixedA: 'LB-VA016A-7289-7041-30-A', codesB: '7289-7041-30', endC: 'A' },
    { sheet: 'LB-VA016B', fixedA: 'LB-VA016B-7289-7041-30-B', codesB: '7289-7041-30', endC: 'B' },
    { sheet: 'LB-VA001B', fixedA: 'LB-VA001B-7283-8794-70-B', codesB: '7283-8794-70', endC: 'B' },
    { sheet: 'LB-VA006',  fixedA: 'LB-VA006-7283-0651-70-A', codesB: '7283-0651-70', endC: 'A' },
    { sheet: 'LB-VA018',  fixedA: 'LB-VA018-7289-8063-30-A', codesB: '7289-8063-30', endC: 'B' },
    { sheet: 'LB-VS001A', fixedA: 'LB-VS001A-7286-1896-A', codesB: '7286-1896', endC: 'A' },
    { sheet: 'LB-VS001B', fixedA: 'LB-VS001B-7286-1896-B', codesB: '7286-1896', endC: 'B' },
    { sheet: 'LB-VS001C', fixedA: 'LB-VS001C-7286-1896-C', codesB: '7286-1896', endC: 'C' },
    { sheet: 'LB-VS002',  fixedA: 'LB-VS002-7286-1902-A', codesB: '7286-1902', endC: 'A' },
    { sheet: 'LB-VS004',  fixedA: 'LB-VS004-7286-1903-A', codesB: '7286-1903', endC: 'A' },
    { sheet: 'LB-VS012C', fixedA: 'LB-VS012C-7287-8821-C', codesB: '7287-8821', endC: 'C' },
    { sheet: 'LB-VS051B', fixedA: 'LB-VS051B-7286-7859-B', codesB: '7286-7859', endC: 'B' }
  ];
}

function listSheetNames() { return Object.keys(getSheetCheckMap()); }
function listLabelSheetNames() { return getLabelRules().map(function(r){ return r.sheet; }); }

function checkLabelSheetByRule(sheet, rule) {
  var last = sheet.getLastRow();
  if (last < 7) return { valid: false };
  var allowB = Array.isArray(rule.codesB) ? rule.codesB : [rule.codesB];
  var rows = sheet.getRange(7, 1, last - 6, 3).getDisplayValues();
  for (var i = 0; i < rows.length; i++) {
    var a = String(rows[i][0] || '').trim();
    var b = String(rows[i][1] || '').trim();
    var c = String(rows[i][2] || '').trim();
    if (!a && !b && !c) continue;
    if (a !== rule.fixedA) return { valid: false };
    if (allowB.indexOf(b) < 0) return { valid: false };
    if (rule.endC && (!c || c.slice(-1) !== rule.endC)) return { valid: false };
  }
  return { valid: true };
}

function getLatestCodeFromSheet(sheet) {
  var last = sheet.getLastRow();
  if (last < 7) return '';
  var raw = String(sheet.getRange(last, 1).getDisplayValue() || '').trim();
  var parts = raw.split('-');
  return parts.length >= 2 ? String(parts[1]).trim() : raw;
}

// [MỚI] Đọc "Số xe" mới nhất của sheet (chỉ đọc, không ảnh hưởng logic kiểm tra).
// Tự tìm cột có tiêu đề chứa "số xe" trong 6 dòng đầu; không có thì trả về ''.
function getXeValue_(sheet) {
  try {
    var lastCol = sheet.getLastColumn(), last = sheet.getLastRow();
    if (lastCol < 1 || last < 7) return '';
    var norm = function (s) {
      return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/\s+/g, ' ').trim();
    };
    var head = sheet.getRange(1, 1, 6, lastCol).getDisplayValues();
    var col = -1;
    for (var r = 0; r < head.length && col < 0; r++) {
      for (var c = 0; c < head[r].length; c++) {
        if (norm(head[r][c]).indexOf('so xe') >= 0) { col = c + 1; break; }
      }
    }
    if (col < 0) return '';
    var vals = sheet.getRange(7, col, last - 6, 1).getDisplayValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var v = String(vals[i][0] || '').trim();
      if (v) return v;
    }
  } catch (e) {}
  return '';
}

// [MỚI] Tự tìm cột "Tên" (theo tiêu đề trong 6 dòng đầu) và lấy giá trị mới nhất.
// Ưu tiên: "tên" > "họ tên" > "tên nhân viên" > "nhân viên" > "người"; cuối cùng là tiêu đề có từ "tên".
function getTenValue_(sheet) {
  try {
    var lastCol = sheet.getLastColumn(), last = sheet.getLastRow();
    if (lastCol < 1 || last < 7) return '';
    var norm = function (s) {
      return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
    };
    var head = sheet.getRange(1, 1, 6, lastCol).getDisplayValues();
    var tiers = [
      function (h) { return h === 'ten'; },
      function (h) { return h === 'ho ten' || h === 'ho va ten'; },
      function (h) { return h === 'ten nhan vien' || h === 'ten nv'; },
      function (h) { return h === 'nhan vien' || h === 'nv'; },
      function (h) { return h === 'nguoi' || h === 'nguoi scan' || h === 'nguoi thuc hien'; },
      function (h) { return (' ' + h + ' ').indexOf(' ten ') >= 0; }
    ];
    var col = -1;
    for (var t = 0; t < tiers.length && col < 0; t++) {
      for (var r = 0; r < head.length && col < 0; r++) {
        for (var c = 0; c < head[r].length; c++) {
          if (tiers[t](norm(head[r][c]))) { col = c + 1; break; }
        }
      }
    }
    if (col < 0) return '';
    var vals = sheet.getRange(7, col, last - 6, 1).getDisplayValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var v = String(vals[i][0] || '').trim();
      if (v) return v;
    }
  } catch (e) {}
  return '';
}

function ensureHistorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('IoT_History');
  if (!sheet) {
    sheet = ss.insertSheet('IoT_History');
    sheet.appendRow(['Thời gian', 'Loại', 'Line', 'Trạng thái', 'Nội dung', 'Số xe', 'Tên']);
  } else {
    if (!sheet.getRange(1, 6).getValue()) sheet.getRange(1, 6).setValue('Số xe');
    if (!sheet.getRange(1, 7).getValue()) sheet.getRange(1, 7).setValue('Tên');
  }
  return sheet;
}

function writeHistory(kind, sheetName, status, text, xe, ten) {
  var sheet = ensureHistorySheet_();
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 7).setValues([[
    Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy'),
    kind || '',
    sheetName || '',
    status || '',
    text || '',
    xe || '',
    ten || ''
  ]]);
}

function writeResultForESP32(announceText, hasErrors, speak) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('IoT_Result');
  if (!sheet) {
    sheet = ss.insertSheet('IoT_Result');
    sheet.appendRow(['Thời gian', 'Trạng thái', 'Câu thông báo', 'Đã phát']);
  }
  if (speak === false) return;
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 4).setValues([[
    Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy'),
    hasErrors ? 'NG' : 'OK',
    announceText,
    'Chưa'
  ]]);
}

function createAnnounceText(hasErrors, results) {
  if (!hasErrors) return 'Tất cả mã khuôn hợp lệ. OK';
  var names = [];
  for (var name in results) {
    if (results[name] && results[name].valid === false) names.push(name);
  }
  if (!names.length) return 'Cảnh báo. Không đạt.';
  return 'Cảnh báo. Line ' + names.join(', ') + '. Không đạt.';
}

function setNextAuto_(key) {
  PropertiesService.getScriptProperties().setProperty(key, String(Date.now() + 60000));
}

function getAutoStatus() {
  var p = PropertiesService.getScriptProperties();
  var now = Date.now();
  function left(k) {
    var t = Number(p.getProperty(k) || 0);
    return t > now ? Math.ceil((t - now) / 1000) : 0;
  }
  return { moldSec: left('nextAuto_mold'), labelSec: left('nextAuto_label') };
}

function scheduleFullCheck() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAllSilent') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('checkAllSilent').timeBased().after(60 * 1000).create();
  setNextAuto_('nextAuto_mold');
}

function scheduleFullLabelCheck() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAllLabelSilent') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('checkAllLabelSilent').timeBased().after(60 * 1000).create();
  setNextAuto_('nextAuto_label');
}

// ====================== [MỚI] EMAIL CẢNH BÁO KHI NG (WEB APP) ======================
// Giống bản Sheet: khi có lỗi sẽ gửi email chi tiết cho quản lý. Không thay đổi logic kiểm tra.
var ALERT_EMAILS = 'si.dt@sanshovn.com.vn,lanh.hv@sanshovn.com.vn,lien.dt@sanshovn.com.vn';

function escHtml_(t) {
  if (t === null || t === undefined) return '';
  return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}

function sendAlertEmail_(subject, title, source, bodyHtml, colored) {
  try {
    var timeStr = Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss dd/MM/yyyy');
    MailApp.sendEmail({
      to: ALERT_EMAILS,
      subject: subject + ' – ' + timeStr,
      htmlBody: '<h2 style="color:#dc2626;">⚠️ ' + escHtml_(title) + '</h2>' +
        '<p><strong>Thời gian:</strong> ' + timeStr + '</p>' +
        '<p><strong>Nguồn:</strong> ' + escHtml_(source) + '</p><hr>' + bodyHtml + '<hr>' +
        '<p><strong>Vui lòng mở file Google Sheets để xem chi tiết đầy đủ và sửa lỗi' +
        (colored ? ' (các ô lỗi đã được tô đỏ)' : '') + '.</strong></p>' +
        '<p>Trân trọng,<br><em>Script kiểm tra tự động</em></p>'
    });
    return true;
  } catch (e) {
    Logger.log('Lỗi gửi email: ' + e.message);
    return false;
  }
}

// Chi tiết lỗi của 1 sheet mã khuôn (dùng result.details do các hàm checkSheetXXX trả về)
function moldErrorHtml_(name, result) {
  var d = (result && result.details) || [];
  var items = d.slice(0, 50).map(function (err) {
    if (typeof err === 'object' && err !== null) {
      var pos = err.row
        ? ('Hàng ' + err.row + (err.col ? ', Cột ' + String.fromCharCode(64 + parseInt(err.col, 10)) : ''))
        : 'Vị trí không xác định';
      return '<li style="margin:8px 0;padding:10px;background:#fff5f5;border-left:4px solid #dc2626;border-radius:4px;">' +
        '<strong>' + pos + ':</strong> ' + escHtml_(err.message || 'Lỗi không xác định') + '<br>' +
        (err.current !== undefined ? '<strong>Hiện tại:</strong> <span style="background:#fee2e2;padding:2px 6px;border-radius:4px;">' + escHtml_(err.current) + '</span><br>' : '') +
        (err.expected !== undefined ? '<strong>Mong đợi:</strong> <span style="background:#fef3c7;padding:2px 6px;border-radius:4px;">' + escHtml_(err.expected) + '</span>' : '') +
        '</li>';
    }
    return '<li style="margin:8px 0;">' + escHtml_(err) + '</li>';
  }).join('');
  if (d.length > 50) items += '<li>... và ' + (d.length - 50) + ' lỗi khác</li>';
  return '<h4 style="color:#dc2626;">📌 Sheet: ' + escHtml_(name) + ' (' + ((result && result.errorCount) || 0) + ' lỗi)</h4>' +
    (items ? '<ul style="padding-left:20px;">' + items + '</ul>'
           : '<p>' + escHtml_((result && result.message) || 'Có lỗi nhưng không có thông tin chi tiết') + '</p>');
}

// Bảng tóm tắt OK/LỖI cho kiểm tra toàn bộ
function moldSummaryHtml_(results) {
  var rows = Object.keys(results).map(function (n) {
    var r = results[n] || {};
    return '<tr style="background-color:' + (r.valid ? '#f0fdf4' : '#fee2e2') + ';">' +
      '<td style="padding:10px;"><strong>' + escHtml_(n) + '</strong></td>' +
      '<td style="padding:10px;font-weight:bold;color:' + (r.valid ? '#166534' : '#dc2626') + ';">' + (r.valid ? 'OK' : 'LỖI') + '</td>' +
      '<td style="padding:10px;text-align:center;font-weight:bold;">' + (r.errorCount || 0) + '</td></tr>';
  }).join('');
  return '<h3>📋 Tóm tắt kết quả</h3>' +
    '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;">' +
    '<tr style="background-color:#dc2626;color:white;"><th>Tên Sheet</th><th>Trạng thái</th><th>Số lỗi</th></tr>' + rows + '</table>';
}

// Chi tiết lỗi label: dùng ĐÚNG điều kiện của checkLabelSheetByRule (chỉ để mô tả, không đổi kết quả)
function getLabelErrorDetails_(sheet, rule) {
  var out = [];
  if (!sheet) return ['<li><strong>Không tìm thấy sheet:</strong> ' + escHtml_(rule.sheet) + '</li>'];
  var last = sheet.getLastRow();
  if (last < 7) return ['<li><strong>Không có dữ liệu từ hàng 7:</strong> ' + escHtml_(rule.sheet) + '</li>'];
  var allowB = Array.isArray(rule.codesB) ? rule.codesB : [rule.codesB];
  var rows = sheet.getRange(7, 1, last - 6, 3).getDisplayValues();
  for (var i = 0; i < rows.length; i++) {
    var a = String(rows[i][0] || '').trim();
    var b = String(rows[i][1] || '').trim();
    var c = String(rows[i][2] || '').trim();
    if (!a && !b && !c) continue;
    var rn = 7 + i;
    if (a !== rule.fixedA) out.push('<li><strong>Dòng ' + rn + '</strong>: Cột A sai → thực tế "<em>' + escHtml_(a) + '</em>", mong đợi "<em>' + escHtml_(rule.fixedA) + '</em>"</li>');
    if (allowB.indexOf(b) < 0) out.push('<li><strong>Dòng ' + rn + '</strong>: Cột B sai → thực tế "<em>' + escHtml_(b || '(trống)') + '</em>", mong đợi một trong "<em>' + escHtml_(allowB.join(', ')) + '</em>"</li>');
    if (rule.endC && (!c || c.slice(-1) !== rule.endC)) out.push('<li><strong>Dòng ' + rn + '</strong>: Cột C sai → thực tế "<em>' + escHtml_(c || '(trống)') + '</em>", phải kết thúc bằng "<em>' + escHtml_(rule.endC) + '</em>"</li>');
  }
  return out;
}

function labelErrorHtml_(name, items) {
  var shown = items.slice(0, 50).join('');
  if (items.length > 50) shown += '<li>... và ' + (items.length - 50) + ' lỗi khác</li>';
  return '<h4 style="color:#dc2626;">📌 Sheet: ' + escHtml_(name) + ' (' + items.length + ' lỗi)</h4><ol>' + shown + '</ol>';
}

function checkOneSheetHeadless(sheetName) {
  var map = getSheetCheckMap();
  if (!map[sheetName]) {
    return { status: 'NG', text: 'Không có line ' + sheetName, sheet: sheetName };
  }
  writeResultForESP32('Đang kiểm tra mã khuôn line ' + sheetName, false, true);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var hasErrors = false;
  var code = '';
  var xe = '';
  var ten = '';
  var result = null;
  if (!sheet || sheet.getLastRow() < 7) {
    hasErrors = true;
  } else {
    result = map[sheetName](sheet);
    hasErrors = !result.valid;
    code = getLatestCodeFromSheet(sheet);
    xe = getXeValue_(sheet);
    ten = getTenValue_(sheet);
  }
  var status = hasErrors ? 'NG' : 'OK';
  var text = hasErrors
    ? ('Cảnh báo. Line ' + sheetName + '. Không đạt.')
    : ('Line ' + sheetName + '. OK');
  Utilities.sleep(5000);
  writeResultForESP32(text, hasErrors, true);
  writeHistory('KHUON', sheetName, status, text, xe, ten);
  scheduleFullCheck();
  var emailSent;
  if (hasErrors) {
    var r1 = result || { valid: false, errorCount: 0, details: [],
      message: !sheet ? 'Sheet không tìm thấy' : 'Không có dữ liệu từ hàng 7' };
    emailSent = sendAlertEmail_('CẢNH BÁO: Lỗi mã khuôn line ' + sheetName,
      'Phát hiện lỗi trong kiểm tra mã khuôn – Line ' + sheetName,
      'Web App SCAN SYSTEM 2.0 – kiểm tra mã khuôn line ' + sheetName,
      moldErrorHtml_(sheetName, r1), true);
  }
  return { status: status, text: text, sheet: sheetName, code: code, emailSent: emailSent };
}

function checkOneLabelHeadless(sheetName) {
  var rules = getLabelRules();
  var rule = null;
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].sheet === sheetName) { rule = rules[i]; break; }
  }
  if (!rule) {
    return { status: 'NG', text: 'Không có label ' + sheetName, sheet: sheetName };
  }
  writeResultForESP32('Đang kiểm tra label line ' + sheetName, false, true);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var hasErrors = !sheet || sheet.getLastRow() < 7 || !checkLabelSheetByRule(sheet, rule).valid;
  var status = hasErrors ? 'NG' : 'OK';
  var text = hasErrors
    ? ('Cảnh báo. Label ' + sheetName + '. Không đạt.')
    : ('Label ' + sheetName + '. OK');
  Utilities.sleep(5000);
  writeResultForESP32(text, hasErrors, true);
  writeHistory('LABEL', sheetName, status, text, sheet ? getXeValue_(sheet) : '', sheet ? getTenValue_(sheet) : '');
  scheduleFullLabelCheck();
  var emailSent;
  if (hasErrors) {
    var det = getLabelErrorDetails_(sheet, rule);
    emailSent = sendAlertEmail_('LỖI TEM NHÃN: ' + sheetName,
      'Phát hiện lỗi trong kiểm tra label – ' + sheetName,
      'Web App SCAN SYSTEM 2.0 – kiểm tra label ' + sheetName,
      labelErrorHtml_(sheetName, det), false);
  }
  return { status: status, text: text, sheet: sheetName, emailSent: emailSent };
}

function checkAllSilent() {
  var map = getSheetCheckMap();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = {};
  var hasErrors = false;
  for (var name in map) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 7) {
      results[name] = { valid: false, errorCount: 0, details: [], message: 'Sheet không tìm thấy hoặc không có dữ liệu từ hàng 7' };
      hasErrors = true;
      continue;
    }
    var result = map[name](sheet);
    results[name] = result;
    if (!result.valid) hasErrors = true;
  }
  var text = createAnnounceText(hasErrors, results);
  writeResultForESP32(text, hasErrors, hasErrors);
  writeHistory('KHUON-ALL', '', hasErrors ? 'NG' : 'OK', text);
  PropertiesService.getScriptProperties().deleteProperty('nextAuto_mold');
  var emailSent;
  if (hasErrors) {
    var body = moldSummaryHtml_(results) + '<hr><h3>🔍 Chi tiết lỗi cụ thể</h3>' +
      Object.keys(results).filter(function (n) { return !results[n].valid; })
        .map(function (n) { return moldErrorHtml_(n, results[n]); }).join('<hr>');
    emailSent = sendAlertEmail_('CẢNH BÁO: Kết quả kiểm tra dữ liệu - Có lỗi',
      'CẢNH BÁO: Phát hiện lỗi trong kiểm tra mã khuôn',
      'Web App SCAN SYSTEM 2.0 – kiểm tra toàn bộ mã khuôn', body, true);
  }
  return { status: hasErrors ? 'NG' : 'OK', text: text, speak: hasErrors, emailSent: emailSent };
}

function checkAllLabelSilent() {
  var rules = getLabelRules();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bad = [];
  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var sheet = ss.getSheetByName(rule.sheet);
    if (!sheet || sheet.getLastRow() < 7 || !checkLabelSheetByRule(sheet, rule).valid) bad.push(rule.sheet);
  }
  var hasErrors = bad.length > 0;
  var text = hasErrors
    ? ('Cảnh báo. Label ' + bad.join(', ') + '. Không đạt.')
    : 'Tất cả label hợp lệ. OK';
  writeResultForESP32(text, hasErrors, hasErrors);
  writeHistory('LABEL-ALL', '', hasErrors ? 'NG' : 'OK', text);
  PropertiesService.getScriptProperties().deleteProperty('nextAuto_label');
  var emailSent;
  if (hasErrors) {
    var total = 0, body = '';
    for (var k = 0; k < rules.length; k++) {
      if (bad.indexOf(rules[k].sheet) < 0) continue;
      var items = getLabelErrorDetails_(ss.getSheetByName(rules[k].sheet), rules[k]);
      total += items.length;
      body += labelErrorHtml_(rules[k].sheet, items);
    }
    emailSent = sendAlertEmail_('LỖI MÃ KHUÔN & TEM NHÃN – ' + total + ' lỗi',
      'Phát hiện ' + total + ' lỗi trong kiểm tra label',
      'Web App SCAN SYSTEM 2.0 – kiểm tra toàn bộ label', body, false);
  }
  return { status: hasErrors ? 'NG' : 'OK', text: text, emailSent: emailSent };
}

function checkAllDataHeadless() { return checkAllSilent(); }

function getRecentHistory(limit) {
  var sheet = ensureHistorySheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var n = Math.min(limit || 20, last - 1);
  var rows = sheet.getRange(2, 1, n, 7).getDisplayValues();
  return rows.map(function(r) {
    return { time: r[0], kind: r[1], line: r[2], status: r[3], text: r[4], xe: r[5], ten: r[6] };
  });
}

function getHistoryStats() {
  var sheet = ensureHistorySheet_();
  var last = sheet.getLastRow();
  var ok = 0, ng = 0;
  if (last >= 2) {
    var vals = sheet.getRange(2, 4, last - 1, 1).getDisplayValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === 'OK') ok++;
      if (String(vals[i][0]) === 'NG') ng++;
    }
  }
  return { ok: ok, ng: ng, total: ok + ng };
}

function espPing() {
  PropertiesService.getScriptProperties().setProperty('espPing', String(Date.now()));
  return { ok: true };
}

function getEspStatus() {
  var t = Number(PropertiesService.getScriptProperties().getProperty('espPing') || 0);
  return {
    online: t > 0 && (Date.now() - t) < 90000,
    last: t ? Utilities.formatDate(new Date(t), 'GMT+7', 'HH:mm:ss') : 'Chưa kết nối'
  };
}

function getLatestResult() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('IoT_Result');
  if (!sheet || sheet.getLastRow() < 2) {
    return { status: 'none', text: 'Chưa có kết quả', time: '', beep: 0, light: 'GREEN' };
  }
  var row = sheet.getRange(2, 1, 1, 4).getValues()[0];
  var status = String(row[1] || '');
  var text = String(row[2] || '');
  var beep = 0;
  if (/đang kiểm tra/i.test(text)) beep = 1;
  if (status === 'NG' || /không đạt/i.test(text)) beep = 3;
  // Tín hiệu đèn tháp cho ESP32-S3: GREEN = OK, YELLOW = đang kiểm tra, RED = phát hiện lỗi.
  // Dùng lại đúng điều kiện đã có ở trên (không thêm luật mới).
  var light = 'GREEN';
  if (beep === 1) light = 'YELLOW';
  if (beep === 3) light = 'RED';
  return { status: status, text: text, time: String(row[0]), beep: beep, light: light };
}

function getDashboard() {
  return {
    esp: getEspStatus(),
    latest: getLatestResult(),
    auto: getAutoStatus(),
    stats: getHistoryStats(),
    history: getRecentHistory(20)
  };
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ui';
  if (action === 'latest') {
    espPing();
    return ContentService.createTextOutput(JSON.stringify(getLatestResult()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'ping') {
    espPing();
    return ContentService.createTextOutput('{"ok":true}')
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>SCAN SYSTEM 2.0</title>
<style>
:root{--bg:#070b14;--card:#10182a;--line:#1e2a44;--text:#eef4ff;--mute:#8aa0c4;--blue2:#2563eb;--ok:#22c55e;--ng:#f43f5e;--warn:#eab308}
*{box-sizing:border-box}
body{margin:0;font-family:Arial,sans-serif;background:radial-gradient(900px 400px at 10% -10%,#1d4ed633,transparent 50%),var(--bg);color:var(--text);min-height:100vh}
.wrap{max-width:860px;margin:0 auto;padding:14px}
.hero,.card{background:linear-gradient(180deg,#152038,#10182a);border:1px solid var(--line);border-radius:22px;padding:16px;margin-bottom:14px}
.hero{position:sticky;top:0;z-index:40;box-shadow:0 12px 24px -12px rgba(0,0,0,.6)}
.brand{margin:0 0 4px;font-size:12px;letter-spacing:.16em;color:#60a5fa;font-weight:800}
h1{margin:0 0 8px;font-size:26px}
.sub{margin:0;color:var(--mute);font-size:13px}
.row{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px}
.pill{background:#0b1220;border:1px solid var(--line);border-radius:999px;padding:8px 12px;font-size:13px;font-weight:800}
.esp{display:flex;align-items:center;gap:8px}
.dot{width:10px;height:10px;border-radius:50%;background:#64748b}
.on .dot{background:var(--ok);box-shadow:0 0 0 6px #22c55e33}
.off .dot{background:var(--ng)}
.tabs{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:14px}
.tab{border:1px solid var(--line);border-radius:14px;padding:14px 8px;font-size:15px;font-weight:900;color:#cbd5e1;background:#0b1220;min-height:52px}
.tab.on-semi{background:linear-gradient(180deg,#38bdf8,#0284c7);color:#fff;border-color:#38bdf8}
.tab.on-auto{background:linear-gradient(180deg,#a78bfa,#6d28d9);color:#fff;border-color:#a78bfa}
.h2{margin:0 0 14px;font-size:13px;letter-spacing:.08em;font-weight:800;text-transform:uppercase}
.khuon .h2{color:#93c5fd}
.label .h2{color:#5eead4}
.khuon{border-color:#1d4ed6}
.label{border-color:#0f766e}
.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px 14px}
.btn{border:0;border-radius:16px;padding:16px 8px;font-size:14px;font-weight:800;color:#fff;background:linear-gradient(180deg,#4f8cff,var(--blue2));min-height:58px}
.btn.lb{background:linear-gradient(180deg,#14b8a6,#0f766e)}
.btn.okbtn{background:linear-gradient(180deg,#4ade80,#15803d)!important}
.btn.ngbtn{background:linear-gradient(180deg,#fb7185,#9f1239)!important;animation:blink 1s infinite}
.btn.checking{background:linear-gradient(180deg,#fde047,#eab308)!important;color:#1f2937!important}
@keyframes blink{50%{opacity:.55}}
.btn:disabled{opacity:.4}
.btn.active{display:flex;align-items:center;justify-content:center;gap:8px}
.spin{width:16px;height:16px;border-radius:50%;border:2px solid #fff6;border-top-color:#fff;animation:sp .7s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
#box{display:flex;align-items:center;gap:16px;padding:18px;border-radius:18px;border:1px solid var(--line);min-height:100px;transition:.25s}
#box.ok{background:linear-gradient(135deg,#052e16,#0a3d20);border-color:#15803d;color:#dcfce7;box-shadow:inset 4px 0 0 var(--ok)}
#box.ng{background:linear-gradient(135deg,#3f0b16,#2a0a12);border-color:#9f1239;color:#ffe4e6;box-shadow:inset 4px 0 0 var(--ng);animation:blink 1.4s infinite}
#box.wait{background:linear-gradient(135deg,#2a2410,#1e293b);border-color:#a16207;color:#fef3c7;box-shadow:inset 4px 0 0 var(--warn);font-size:16px;font-weight:800}
.rs-ic{flex:none;width:60px;height:60px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:30px;font-weight:900}
#box.ok .rs-ic{background:#22c55e26;color:#4ade80;box-shadow:0 0 0 6px #22c55e14}
#box.ng .rs-ic{background:#f43f5e26;color:#fb7185;box-shadow:0 0 0 6px #f43f5e14}
#box.wait .rs-ic{background:#eab30826;box-shadow:0 0 0 6px #eab30814}
.rs-ic .spin{width:28px;height:28px;border-width:3px;border-color:#eab30855;border-top-color:#facc15}
.rs-main{min-width:0;flex:1}
.rs-cap{font-size:10.5px;letter-spacing:.14em;color:var(--mute);font-weight:800}
.rs-title{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:3px 0;font-size:26px;font-weight:900;color:#fff}
.rs-badge{padding:3px 12px;border-radius:999px;font-size:12px;letter-spacing:.08em;color:#fff}
#box.ok .rs-badge{background:#15803d}#box.ng .rs-badge{background:#be123c}#box.wait .rs-badge{background:#a16207}
.rs-msg{font-size:14px;font-weight:600;opacity:.85}
.rs-chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:9px}
#rsTime{margin-top:10px;font-size:12px}
.wait{background:#1e293b;color:#fdba74}
.ok{background:#052e16;color:#86efac}
.ng{background:#3f0b16;color:#fda4af}
.modal-overlay{position:fixed;inset:0;background:rgba(2,6,15,.82);display:flex;align-items:center;justify-content:center;z-index:999;padding:20px}
.modal-overlay.hidden{display:none}
.modal-card{width:100%;max-width:460px;background:linear-gradient(180deg,#152038,#0d1526);border:2px solid var(--line);border-radius:26px;padding:32px 24px;text-align:center}
.modal-card.ok{border-color:var(--ok)}
.modal-card.ng{border-color:var(--ng);animation:blink 1s infinite}
.modal-status{font-size:44px;font-weight:900;margin-bottom:10px}
.modal-card.ok .modal-status{color:var(--ok)}
.modal-card.ng .modal-status{color:var(--ng)}
.modal-text{font-size:19px;font-weight:700;margin-bottom:10px}
.modal-time{color:var(--mute);font-size:13px;margin-bottom:26px}
.modal-mail{font-size:14px;font-weight:700;color:#fbbf24;margin:-14px 0 22px}
.modal-close{width:100%;padding:18px;border-radius:16px;border:0;font-size:17px;font-weight:900;color:#fff;background:linear-gradient(180deg,#4f8cff,var(--blue2));min-height:56px}
.okt{color:#4ade80;font-weight:800}.ngt{color:#fb7185;font-weight:800}
.hist{max-height:340px;overflow:auto;padding-right:4px;scrollbar-width:thin;scrollbar-color:#334155 transparent}
.hist::-webkit-scrollbar{width:6px}.hist::-webkit-scrollbar-thumb{background:#334155;border-radius:6px}
.hh,.item{display:grid;grid-template-columns:92px 112px 1fr 64px;gap:10px;align-items:center}
.hh{padding:0 14px 8px;font-size:10px;letter-spacing:.12em;color:var(--mute);font-weight:800;position:sticky;top:0;background:#10182a;z-index:2}
.hh span:last-child{text-align:right}
.item{padding:10px 14px;margin-bottom:6px;background:#0b1220;border:1px solid var(--line);border-left:4px solid var(--ok);border-radius:12px;font-size:13px;color:#cbd5e1;transition:.15s}
.item:hover{background:#111b30;transform:translateX(2px)}
.item.bad{border-left-color:var(--ng);background:#170d18}
.tm b{display:block;font-size:13px;color:#f1f5f9;font-variant-numeric:tabular-nums}
.tm small{color:var(--mute);font-size:11px}
.kd{justify-self:start;padding:3px 9px;border-radius:8px;font-size:10.5px;font-weight:800;letter-spacing:.05em;white-space:nowrap}
.kd.k{background:#1e3a8a55;color:#93c5fd;border:1px solid #1d4ed6}
.kd.l{background:#0f766e55;color:#5eead4;border:1px solid #0f766e}
.kd.a{background:#4c1d9555;color:#c4b5fd;border:1px solid #6d28d9}
.ln{display:flex;flex-wrap:wrap;align-items:center;gap:6px;min-width:0}
.ln b{color:#f8fafc;font-size:14px}
.ln .all{color:var(--mute);font-size:12px}
.chip{display:inline-flex;align-items:center;gap:5px;padding:2px 10px;border-radius:999px;font-size:11px;font-weight:700;border:1px solid transparent;white-space:nowrap}
.chip i{font-style:normal;font-weight:600;opacity:.7}
.chip.xe{background:#1e3a8a44;color:#bfdbfe;border-color:#1d4ed6}
.chip.ten{background:#14532d44;color:#bbf7d0;border-color:#15803d}
.st{justify-self:end;padding:5px 12px;border-radius:999px;font-size:12px;font-weight:900;letter-spacing:.06em}
.st.p{background:#052e16;color:#4ade80;border:1px solid #15803d}
.st.f{background:#3f0b16;color:#fb7185;border:1px solid #9f1239}
.empty{padding:26px;text-align:center;color:var(--mute);font-size:13px}
@media(max-width:560px){.hh{display:none}.item{grid-template-columns:76px 1fr 54px;grid-template-areas:"tm ln st" "tm kd kd";row-gap:6px}.tm{grid-area:tm}.kd{grid-area:kd}.ln{grid-area:ln}.st{grid-area:st}}
.time{margin-top:6px;font-size:12px;color:var(--mute)}
@media(min-width:720px){.grid{grid-template-columns:repeat(3,1fr)}h1{font-size:32px}}
</style>
</head>
<body>
<div class="wrap">
  <div class="hero">
    <p class="brand">SCAN SYSTEM 2.0</p>
    <h1>Kiểm tra mã khuôn</h1>
    <p class="sub">Chọn đúng line vừa scan. Semi line = VS. Auto line = VA. Logic kiểm tra không đổi.</p>
    <div class="row">
      <div id="esp" class="pill esp off"><span class="dot"></span><span id="espText">ESP32 Offline</span></div>
      <div class="pill" id="speak">Đèn: chờ thao tác</div>
      <div class="pill" id="stat">OK 0 / NG 0</div>
    </div>
    <div class="time" id="espTime"></div>
    <div class="time" id="cd">Tự động kiểm tra toàn bộ: --</div>
    <div class="tabs">
      <button class="tab on-semi" id="tabSemi" onclick="setMode('semi')">SEMI LINE</button>
      <button class="tab" id="tabAuto" onclick="setMode('auto')">AUTO LINE</button>
    </div>
  </div>
  <div class="card khuon">
    <p class="h2" id="hKhuon">Kiểm tra mã khuôn — SEMI (VS)</p>
    <div class="grid" id="grid"></div>
  </div>
  <div class="card label">
    <p class="h2" id="hLabel">Kiểm tra label — SEMI (VS)</p>
    <div class="grid" id="gridLb"></div>
  </div>
  <div class="card">
    <p class="h2">Kết quả</p>
    <div id="box" class="wait">Chưa kiểm tra</div>
    <div class="time" id="rsTime"></div>
  </div>
  <div class="card">
    <p class="h2">Lịch sử 20 lần gần nhất</p>
    <div class="hist" id="hist"></div>
  </div>
</div>
<div class="modal-overlay hidden" id="resultModal">
  <div class="modal-card" id="modalCard">
    <div class="modal-status" id="modalStatus">OK</div>
    <div class="modal-text" id="modalText"></div>
    <div class="modal-time" id="modalTime"></div>
    <div class="modal-mail" id="modalMail"></div>
    <button class="modal-close" onclick="closeResultModal()">ĐÓNG</button>
  </div>
</div>
<script>
var busy=false, mode='semi', allMold=[], allLabel=[];
function isSemi(n){ return n.indexOf('VS')>=0; }
function isAuto(n){ return n.indexOf('VA')>=0; }
var KIND={'KHUON':['MÃ KHUÔN','k'],'LABEL':['LABEL','l'],'KHUON-ALL':['TẤT CẢ · KHUÔN','a'],'LABEL-ALL':['TẤT CẢ · LABEL','a']};
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function paintResult(j,H){
  var st=j.light==='YELLOW'?'wait':(j.status==='NG'?'ng':'ok');
  var txt=String(j.text||'');
  var m=/(Line|Label)[ ]+([A-Za-z0-9-]+)([.,])/.exec(txt);
  var title,cap;
  if(m&&m[3]==='.'){title=m[2];cap=m[1]==='Label'?'KIỂM TRA LABEL':'KIỂM TRA MÃ KHUÔN';}
  else if(txt.indexOf('Cảnh báo')>=0){title='Nhiều line không đạt';cap='KIỂM TRA TỰ ĐỘNG';}
  else{title='Toàn bộ hệ thống';cap='KIỂM TRA TỰ ĐỘNG';}
  var label={ok:'OK',ng:'NG',wait:'ĐANG KIỂM TRA'}[st];
  var icon={ok:'✓',ng:'✕',wait:'<span class="spin"></span>'}[st];
  var h0=(H||[])[0], chips='';
  if(h0&&h0.text===txt){
    if(h0.xe) chips+='<span class="chip xe"><i>Xe</i>'+esc(h0.xe)+'</span>';
    if(h0.ten) chips+='<span class="chip ten"><i>👤</i>'+esc(h0.ten)+'</span>';
  }
  var b=document.getElementById('box');
  b.className=st;
  b.innerHTML='<div class="rs-ic">'+icon+'</div><div class="rs-main"><div class="rs-cap">'+cap+'</div>'+
    '<div class="rs-title">'+esc(title)+'<span class="rs-badge">'+label+'</span></div>'+
    '<div class="rs-msg">'+esc(txt)+'</div>'+(chips?'<div class="rs-chips">'+chips+'</div>':'')+'</div>';
  document.getElementById('rsTime').innerText='🕒 '+(j.time||'');
}
function paintDash(d){
  if(!d) return;
  document.getElementById('esp').className='pill esp '+(d.esp.online?'on':'off');
  document.getElementById('espText').innerText=d.esp.online?'ESP32 Online':'ESP32 Offline';
  document.getElementById('espTime').innerText='Ping: '+(d.esp.last||'');
  var a=d.auto||{}, parts=[];
  if(a.moldSec>0) parts.push('Mã khuôn '+a.moldSec+' giây');
  if(a.labelSec>0) parts.push('Label '+a.labelSec+' giây');
  document.getElementById('cd').innerText=parts.length?('Tự động kiểm tra toàn bộ sau: '+parts.join(' · ')):'Tự động kiểm tra toàn bộ: chờ thao tác';
  document.getElementById('stat').innerText='OK '+(d.stats.ok||0)+' / NG '+(d.stats.ng||0);
  var j=d.latest;
  if(j&&j.status!=='none'){
    paintResult(j,d.history);
    var lightKey=j.light||(j.status==='NG'?'RED':(j.status==='OK'?'GREEN':'YELLOW'));
    var lightInfo={GREEN:{label:'🟢 Đèn xanh — OK',cls:'okt'},YELLOW:{label:'🟡 Đèn vàng — Đang kiểm tra',cls:''},RED:{label:'🔴 Đèn đỏ — Không đạt',cls:'ngt'}}[lightKey]||{label:'Đèn: --',cls:''};
    var speakEl=document.getElementById('speak');
    speakEl.innerText=lightInfo.label;
    speakEl.classList.remove('okt','ngt');
    if(lightInfo.cls) speakEl.classList.add(lightInfo.cls);
  }
  var H=d.history||[];
  document.getElementById('hist').innerHTML=H.length?
    '<div class="hh"><span>THỜI GIAN</span><span>LOẠI</span><span>LINE / CHI TIẾT</span><span>KẾT QUẢ</span></div>'+
    H.map(function(r){
      var t=String(r.time||'').split(' ');
      var k=KIND[r.kind]||[r.kind||'--','a'];
      var bad=r.status==='NG';
      var det=r.line?'<b>'+esc(r.line)+'</b>':'<span class="all">'+esc(r.text)+'</span>';
      var xe=r.xe?'<span class="chip xe"><i>Xe</i>'+esc(r.xe)+'</span>':'';
      var ten=r.ten?'<span class="chip ten"><i>👤</i>'+esc(r.ten)+'</span>':'';
      return '<div class="item'+(bad?' bad':'')+'"><div class="tm"><b>'+esc(t[0])+'</b><small>'+esc(t[1]||'')+'</small></div>'+
        '<span class="kd '+k[1]+'">'+k[0]+'</span><div class="ln">'+det+xe+ten+'</div>'+
        '<span class="st '+(bad?'f':'p')+'">'+esc(r.status)+'</span></div>';
    }).join('')
    :'<div class="empty">Chưa có lịch sử</div>';
}
function markBtn(btn,status){
  if(!btn) return;
  btn.classList.remove('okbtn','ngbtn','checking');
  if(btn._resetTimer){clearTimeout(btn._resetTimer);btn._resetTimer=null;}
  if(status==='OK') btn.classList.add('okbtn');
  if(status==='NG') btn.classList.add('ngbtn');
  if(status==='OK'||status==='NG'){
    btn._resetTimer=setTimeout(function(){btn.classList.remove('okbtn','ngbtn');btn._resetTimer=null;},5000);
  }
}
function setBusy(v,btn){
  busy=v;
  document.querySelectorAll('.btn').forEach(function(b){b.disabled=v;});
  if(btn){
    if(v){
      btn.classList.remove('okbtn','ngbtn');
      btn.classList.add('active','checking');
      btn.innerHTML='<span class="spin"></span><span>ĐANG KIỂM TRA</span>';
    } else {
      btn.classList.remove('active','checking');
      btn.innerText=btn.dataset.name;
    }
  }
}
function openResultModal(j){
  var card=document.getElementById('modalCard');
  card.classList.remove('ok','ng');
  card.classList.add(j.status==='NG'?'ng':'ok');
  document.getElementById('modalStatus').innerText=j.status||'';
  document.getElementById('modalText').innerText=j.text||'';
  document.getElementById('modalTime').innerText=j.time||'';
  document.getElementById('modalMail').innerText=j.emailSent===true?'📧 Đã gửi email cảnh báo đến quản lý':(j.emailSent===false?'⚠️ Chưa gửi được email — kiểm tra quyền gửi mail':'');
  document.getElementById('resultModal').classList.remove('hidden');
}
function closeResultModal(){ document.getElementById('resultModal').classList.add('hidden'); }
function run(fn,name,btn){
  if(busy) return;
  setBusy(true,btn);
  document.getElementById('box').className='wait';
  document.getElementById('box').innerText='SCAN SYSTEM — '+name;
  google.script.run
    .withSuccessHandler(function(j){setBusy(false,btn);markBtn(btn,j.status);openResultModal(j);tick();})
    .withFailureHandler(function(e){document.getElementById('box').className='ng';document.getElementById('box').innerText='Lỗi: '+e.message;setBusy(false,btn);})
    [fn](name);
}
function drawGrid(id,names,fn,cls){
  var g=document.getElementById(id); g.innerHTML='';
  (names||[]).forEach(function(n){
    var b=document.createElement('button');
    b.className=cls; b.innerText=n; b.dataset.name=n;
    b.onclick=function(){run(fn,n,b);};
    g.appendChild(b);
  });
}
function renderMode(){
  var mold=allMold.filter(mode==='semi'?isSemi:isAuto);
  var lab=allLabel.filter(mode==='semi'?isSemi:isAuto);
  document.getElementById('tabSemi').className='tab'+(mode==='semi'?' on-semi':'');
  document.getElementById('tabAuto').className='tab'+(mode==='auto'?' on-auto':'');
  document.getElementById('hKhuon').innerText=mode==='semi'?'Kiểm tra mã khuôn — SEMI (VS)':'Kiểm tra mã khuôn — AUTO (VA)';
  document.getElementById('hLabel').innerText=mode==='semi'?'Kiểm tra label — SEMI (VS)':'Kiểm tra label — AUTO (VA)';
  drawGrid('grid',mold,'checkOneSheetHeadless','btn');
  drawGrid('gridLb',lab,'checkOneLabelHeadless','btn lb');
}
function setMode(m){ mode=m; renderMode(); }
function tick(){ google.script.run.withSuccessHandler(paintDash).getDashboard(); }
google.script.run.withSuccessHandler(function(n){ allMold=n||[]; renderMode(); }).listSheetNames();
google.script.run.withSuccessHandler(function(n){ allLabel=n||[]; renderMode(); }).listLabelSheetNames();
tick();
setInterval(tick,1000);
</script>
</body>
</html>
  `).setTitle('SCAN SYSTEM 2.0').addMetaTag('viewport','width=device-width, initial-scale=1');
}
