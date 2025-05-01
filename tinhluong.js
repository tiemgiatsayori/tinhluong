let originalData = []; // Store the original parsed data globally

document
  .getElementById("fileInput")
  .addEventListener("change", handleFileUpload);

function handleFileUpload(event) {
  const file = event.target.files[0];
  const errorMessage = document.getElementById("errorMessage");
  errorMessage.style.display = "none";
  errorMessage.textContent = "";

  if (!file) {
    return;
  }

  // Validate file type
  if (
    ![
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
    ].includes(file.type)
  ) {
    errorMessage.textContent =
      "Vui lòng chọn một tệp Excel hợp lệ (.xlsx hoặc .xls).";
    errorMessage.style.display = "block";
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = "Cham_Cong";
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
      });

      if (jsonData.length === 0) {
        throw new Error("Sheet is empty!");
      }

      // Extract headers dynamically from the first row
      const headers = jsonData[0];

      // Trim all values in the dataset
      const trimmedData = jsonData.map((row) =>
        row.map((cell) => (cell || "").trim())
      );

      // Store trimmed data globally
      originalData = trimmedData;

      // Display table with trimmed data
      renderTable(headers, trimmedData.slice(1));

      // Extract unique "Nhân viên" list (case-sensitive for rendering)
      const nameIndex = headers.indexOf("Tên");
      if (nameIndex === -1) {
        throw new Error("Column 'Tên' not found in the uploaded file.");
      }

      const uniqueNames = [
        ...new Set(
          trimmedData
            .slice(1)
            .map((row) => row[nameIndex])
            .filter((name) => name)
        ),
      ];

      renderEmployeeReport(uniqueNames);
    } catch (error) {
      console.error("Error processing the Excel file:", error);
      errorMessage.textContent = `Có lỗi xảy ra: ${error.message}`;
      errorMessage.style.display = "block";
    } finally {
      reader.onload = null;
    }
  };

  reader.readAsArrayBuffer(file);
}

function renderTable(headers, tableData) {
  let tableHTML =
    "<thead><tr>" +
    headers.map((h) => `<th>${h}</th>`).join("") +
    "</tr></thead><tbody>";
  tableData.forEach((row) => {
    tableHTML +=
      "<tr>" +
      headers.map((_, i) => `<td>${row[i] || ""}</td>`).join("") +
      "</tr>";
  });
  tableHTML += "</tbody>";
  document.getElementById("excelTable").innerHTML = tableHTML;
}

function renderEmployeeReport(uniqueNames) {
  const employeeList = document.getElementById("employeeList");
  const toggleText = document.getElementById("toggleText");
  const reportSection = document.getElementById("report");

  // Reset list
  employeeList.innerHTML = "";

  // Show first 10 names, hide the rest
  uniqueNames.forEach((name, index) => {
    let nameItem = `<div class="${
      index >= 10 ? "extra-names" : ""
    }">${name}</div>`;
    employeeList.innerHTML += nameItem;
  });

  // Show/hide "Xem thêm" text
  if (uniqueNames.length > 10) {
    toggleText.style.display = "inline-block";
    toggleText.textContent = "Xem thêm";
  } else {
    toggleText.style.display = "none";
  }

  document.getElementById("employeeCount").textContent = uniqueNames.length;
  reportSection.style.display = "block";

  // Toggle Show More / Show Less functionality
  toggleText.onclick = function () {
    let extraNames = document.querySelectorAll(".extra-names");
    if (toggleText.textContent === "Xem thêm") {
      extraNames.forEach((el) => (el.style.display = "block"));
      toggleText.textContent = "Ẩn bớt";
    } else {
      extraNames.forEach((el) => (el.style.display = "none"));
      toggleText.textContent = "Xem thêm";
    }
  };
}

// Export Button Functionality
document.getElementById("exportButton").addEventListener("click", exportExcel);

function exportExcel() {
  if (originalData.length === 0) {
    alert("Không có dữ liệu để xuất!");
    return;
  }

  const headers = originalData[0];
  const tableData = originalData.slice(1);

  const nameIndex = headers.indexOf("Tên");
  if (nameIndex === -1) {
    alert("Không tìm thấy cột 'Tên' trong dữ liệu!");
    return;
  }

  // Group data by employee name (case-insensitive)
  const groupedData = {};
  const nameMapping = new Map(); // Map lowercase name to original name

  tableData.forEach((row) => {
    if (row?.length) {
      const originalName = row[nameIndex];
      const lowerCaseName = originalName.toLowerCase();
  
      if (!groupedData[lowerCaseName]) {
        groupedData[lowerCaseName] = [];
        nameMapping.set(lowerCaseName, originalName); // Store the original name
      }
  
      groupedData[lowerCaseName].push(row); // Add the row to the employee's array
    }
  });

  // Define the desired columns to include in the export
  const desiredColumns = ["Ngày", "Tên", "Vào/Tan Ca", "Chi nhánh"];
  const desiredColumnIndices = desiredColumns
    .map((col) => headers.indexOf(col))
    .filter((index) => index !== -1);

  if (desiredColumnIndices.length === 0) {
    alert("Không tìm thấy các cột cần thiết để xuất!");
    return;
  }

  // Filter headers to include only desired columns
  const filteredHeaders = desiredColumnIndices.map((index) => headers[index]);

  // Add the new column "Giờ lương" to the headers
  filteredHeaders.push("Giờ lương");

  // Create a new workbook
  const workbook = XLSX.utils.book_new();

  function customRound(value) {
    const integerPart = Math.floor(value); // Get the integer part
    const decimalPart = value - integerPart; // Get the decimal part

    if (decimalPart < 0.5) {
      return integerPart; // Drop the decimal part
    } else if (decimalPart >= 0.5 && decimalPart < 0.75) {
      return integerPart + 0.5; // Round to 0.5
    } else {
      return integerPart + 1; // Round up to the next integer
    }
  }

  // Process each employee's data
  for (const [lowerCaseName, rows] of Object.entries(groupedData)) {
    const originalName = nameMapping.get(lowerCaseName); // Get the original name for the sheet title

    // Filter data rows to include only desired columns
    const filteredData = rows.map((row) => {
      const rowData = desiredColumnIndices.map((index) => row[index]);
      return rowData;
    });

    // Sort the filtered data by the "Ngày" column in ascending order
    const dateIndex = filteredHeaders.indexOf("Ngày"); // Find the index of the "Ngày" column
    filteredData.sort((a, b) => {
      const dateA = new Date(a[dateIndex]); // Parse the date from the "Ngày" column
      const dateB = new Date(b[dateIndex]); // Parse the date from the "Ngày" column
      return dateA - dateB; // Compare the dates for sorting
    });

    // Create a new array to hold the final processed data
    const processedData = [];

    // Calculate "Giờ lương" values
    for (let i = 0; i < filteredData.length; i++) {
      const currentRow = filteredData[i];
      const shiftTypeIndex = filteredHeaders.indexOf("Vào/Tan Ca");

      // Check for consecutive "Tan Ca" entries
      if (
        currentRow[shiftTypeIndex] === "Tan Ca" &&
        i > 0 &&
        filteredData[i - 1][shiftTypeIndex] === "Tan Ca"
      ) {
        // Insert a new row with "Không chấm vào ca" in the "Ngày" column
        const newRow = Array(filteredHeaders.length + 1).fill(""); // Create an empty row
        newRow[dateIndex] = "Không chấm vào ca"; // Add text to the "Ngày" column
        processedData.push(newRow); // Insert the new row
      }

      if (currentRow[shiftTypeIndex] === "Vào Ca") {
        // Find the next "Tan Ca" entry
        let tanCaEntry = null;
        if (
          filteredData[i + 1] &&
          filteredData[i + 1][shiftTypeIndex] === "Tan Ca"
        ) {
          tanCaEntry = filteredData[i + 1];
        }

        if (tanCaEntry) {
          // Parse dates and calculate time difference in hours
          const dateString = currentRow[dateIndex]; // e.g., '13/02/2025 08:00'
          const [day, month, yearTime] = dateString.split("/");
          const [year, time] = yearTime.split(" ");
          const [hours, minutes] = time.split(":");
          const enterDate = new Date(year, month - 1, day, hours, minutes);

          const exitDateString = tanCaEntry[dateIndex];
          const [exitDay, exitMonth, exitYearTime] = exitDateString.split("/");
          const [exitYear, exitTime] = exitYearTime.split(" ");
          const [exitHours, exitMinutes] = exitTime.split(":");
          const exitDate = new Date(
            exitYear,
            exitMonth - 1,
            exitDay,
            exitHours,
            exitMinutes
          );

          const timeDifference = (exitDate - enterDate) / (1000 * 60 * 60); // Difference in hours

          // Check if the time difference is greater than 24
          if (timeDifference > 24) {
            currentRow.push(""); // Leave "Giờ lương" blank
          } else {
            currentRow.push(customRound(timeDifference)); // Use custom rounding
          }
        } else {
          // No matching "Tan Ca" found, skip calculation and add an empty row
          currentRow.push(""); // Leave "Giờ lương" blank
          processedData.push(currentRow); // Add the current row
          const newRow = Array(filteredHeaders.length + 1).fill(""); // Create an empty row
          newRow[dateIndex] = "Không chấm tan ca"; // Add text to the "Ngày" column
          processedData.push(newRow);
          continue; // Skip to the next iteration
        }
      } else {
        currentRow.push(""); // Not a "Vào Ca" row
      }

      // Add the current row to the processed data
      processedData.push(currentRow);
    }

    const gioLuongIndex = filteredHeaders.indexOf("Giờ lương");

    if (gioLuongIndex !== -1) {
      const sumStartRow = 2;
      const sumEndRow = processedData.length + 1;
      const colLetter = XLSX.utils.encode_col(gioLuongIndex);

      // Row: Tổng giờ lương
      const sumRow = Array(filteredHeaders.length).fill("");
      sumRow[gioLuongIndex - 1] = "Tổng giờ lương";
      sumRow[gioLuongIndex] = {
        f: `SUM(${colLetter}${sumStartRow}:${colLetter}${sumEndRow})`,
      };
      processedData.push(sumRow);

      // Row: Lương mỗi giờ
      const luongMoiGioRow = Array(filteredHeaders.length).fill("");
      luongMoiGioRow[gioLuongIndex - 1] = "Lương mỗi giờ";
      luongMoiGioRow[gioLuongIndex] = 15000;
      processedData.push(luongMoiGioRow);

      // Row: Tổng lương
      const tongLuongRow = Array(filteredHeaders.length).fill("");
      tongLuongRow[gioLuongIndex - 1] = "Tổng lương";

      const tongGioExcelRow = processedData.length - 1 + 1; // "Tổng giờ lương" Excel row
      const rateExcelRow = processedData.length + 1; // "Lương mỗi giờ" Excel row

      tongLuongRow[gioLuongIndex] = {
        f: `${colLetter}${tongGioExcelRow}*${colLetter}${rateExcelRow}`,
      };
      processedData.push(tongLuongRow);

      // Row: Tạm ứng
      const tamUngRow = Array(filteredHeaders.length).fill("");
      tamUngRow[gioLuongIndex - 1] = "Tạm ứng";
      tamUngRow[gioLuongIndex] = 0;
      processedData.push(tamUngRow);

      // Row: Thực lãnh = Tổng lương - Tạm ứng
      const thucLanhRow = Array(filteredHeaders.length).fill("");
      thucLanhRow[gioLuongIndex - 1] = "Thực lãnh";

      const tongLuongExcelRow = processedData.length; // just pushed "Tạm ứng"
      const tamUngExcelRow = processedData.length + 1;

      thucLanhRow[gioLuongIndex] = {
        f: `${colLetter}${tongLuongExcelRow} - ${colLetter}${tamUngExcelRow}`,
      };
      processedData.push(thucLanhRow);
    }

    // Combine filtered headers and processed data
    const sheetData = [filteredHeaders, ...processedData];
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

    // Set column widths
    worksheet["!cols"] = [
      { wch: 15 }, // Ngày
      { wch: 20 }, // Tên
      { wch: 25 }, // Vào/Tan Ca
      { wch: 20 }, // Chi nhánh
      { wch: 15 }, // Giờ lương
    ];

    // Apply styles to header cells
    filteredHeaders.forEach((header, colIndex) => {
      // Optional: Format "Giờ lương" column numbers
      if (header === "Giờ lương") {
        for (let rowIndex = 1; rowIndex <= processedData.length; rowIndex++) {
          const cell =
            worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];
          if (cell && !isNaN(cell.v)) {
            cell.t = "n";
            cell.z = "0.0";
          }
        }
      }
    });

    XLSX.utils.book_append_sheet(workbook, worksheet, originalName); // Use original name as sheet name
  }

  // Get the current date
  const today = new Date();
  const day = String(today.getDate()).padStart(2, "0"); // Ensure two digits for day
  const month = String(today.getMonth() + 1).padStart(2, "0"); // Ensure two digits for month (months are zero-based)
  const year = today.getFullYear(); // Full year

  // Create the file name with the date appended
  const fileName = `Bảng_lương_nhân_viên_Ori_${day}_${month}_${year}.xlsx`;

  // Export the workbook as an Excel file
  XLSX.writeFile(workbook, fileName);
}
