const excelUrl = '/get-excel';

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
  const statusDiv = document.getElementById("status");
  statusDiv.textContent = "Running script...";

  fetch("/run-script", { method: "POST" })
    .then(response => response.json())
    .then(data => {
        if (data.status === "success") {
        statusDiv.textContent = "✅ Script ran successfully.";
        // Reload Excel after script runs
        loadExcelLog();
        // Hide the message after 3 seconds
        setTimeout(() => {
            statusDiv.textContent = "";
        }, 3000);
        }

    })
    .catch(error => {
      statusDiv.textContent = "❌ Request failed.";
      console.error(error);
    });
});



function loadExcelLog() {
  const outputDiv = document.getElementById("output");
  outputDiv.textContent = "Loading Excel file...";

  fetch("/get-status")
  .then(response => response.json())
  .then(rows => {
    renderExcelTable(rows);
  })
  .catch(error => {
    console.error(error);
    outputDiv.textContent = "⚠️ Failed to load status data.";
  });

}

function updateStatus(messageId, newStatus) {
  fetch('/update_status', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message_id: messageId,
      status: newStatus
    })
  })
  .then(response => response.json())
  .then(data => {
    const statusDiv = document.getElementById("statusMessage");
    if (data.success) {
      statusDiv.textContent = "✅ Status updated successfully.";
      statusDiv.style.color = "green";
      loadExcelLog();
      setTimeout(() => {
        statusDiv.textContent = "";
      }, 3000);
    } else {
      statusDiv.textContent = "❌ Failed to update status.";
      statusDiv.style.color = "red";
    }
  })
  .catch(error => {
    const statusDiv = document.getElementById("statusMessage");
    statusDiv.textContent = "❌ Error updating status.";
    statusDiv.style.color = "red";
    console.error(error);
  });
}



function renderExcelTable(rows) {
  let html = `
    <table class="excel-table">
      <thead>
        <tr>
          <th>Message_ID</th>
          <th>Email Sender</th>
          <th>Date</th>
          <th>Time Received</th>
          <th>Wait Time</th>
          <th>Status</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody>
  `;

  rows.forEach(row => {
    html += `
      <tr>
        <td>${row.Message_ID}</td>
        <td>${row["Email Sender"]}</td>
        <td>${row.Date}</td>
        <td>${row["Time Received"]}</td>
        <td>${row["Wait Time"]}</td>
        <td>
          <select id="status-${row.Message_ID}">
            <option ${row.Status === "Not started" ? "selected" : ""}>Not started</option>
            <option ${row.Status === "In progress" ? "selected" : ""}>In progress</option>
            <option ${row.Status === "Done" ? "selected" : ""}>Done</option>
          </select>
        </td>
        <td>
          <button onclick="updateStatus('${row.Message_ID}', document.getElementById('status-${row.Message_ID}').value)">💾 Save</button>
        </td>
      </tr>
    `;
  });

  html += "</tbody></table>";

  document.getElementById("output").innerHTML = html;
}



function loadInbox() {
  const inboxDiv = document.getElementById("emailInbox");
  inboxDiv.textContent = "Loading emails...";

  fetch("/get-emails")
    .then(response => {
      if (!response.ok) {
        throw new Error(`HTTP error: ${response.status}`);
      }
      return response.json();
    })
    .then(emails => {
      if (emails.length === 0) {
        inboxDiv.textContent = "✅ No unread emails.";
        return;
      }

      // Sort by date descending
      emails.sort((a, b) => {
        const dateA = new Date(a.date);
        const dateB = new Date(b.date);
        return dateB - dateA;
      });

      const ul = document.createElement("ul");
      ul.style.listStyle = "none";
      ul.style.padding = "0";
      ul.style.margin = "0";

      emails.forEach(email => {
        const li = document.createElement("li");
        li.style.marginBottom = "12px";
        li.style.padding = "10px";
        li.style.background = "#fff";
        li.style.border = "1px solid #ddd";
        li.style.borderRadius = "4px";

        li.innerHTML = `
          <strong>From:</strong> ${email.from}<br/>
          <strong>Subject:</strong> ${email.subject}<br/>
          <small>${email.date}</small>
        `;
        ul.appendChild(li);
      });

      inboxDiv.innerHTML = "";
      inboxDiv.appendChild(ul);
    })
    .catch(error => {
      console.error(error);
      inboxDiv.textContent = "❌ Failed to load emails.";
    });
}
