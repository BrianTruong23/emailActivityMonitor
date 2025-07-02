const excelUrl = 'result/log2.xlsx';

fetch(excelUrl)
    .then(response => {
    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }
    return response.arrayBuffer();
    })
    .then(arrayBuffer => {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    let html = XLSX.utils.sheet_to_html(worksheet);

    // Remove default inline styles from SheetJS HTML
    html = html.replace(/<table/g, '<table class="excel-table"');
    document.getElementById('output').innerHTML = html;
    })
    .catch(error => {
    console.error('Error loading Excel file:', error);
    document.getElementById('output').textContent = '⚠️ Failed to load Excel file.';
    });


document.getElementById("runButton").addEventListener("click", () => {
  document.getElementById("status").textContent = "Running script...";
  fetch("/run-script", { method: "POST" })
    .then(response => response.json())
    .then(data => {
      if (data.status === "success") {
        document.getElementById("status").textContent = "✅ Script ran successfully.";
      } else {
        document.getElementById("status").textContent = "❌ Error: " + data.message;
      }
    })
    .catch(error => {
      document.getElementById("status").textContent = "❌ Request failed.";
      console.error(error);
    });
});
