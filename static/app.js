async function lookup() {
  const survey_id = document.getElementById("input").value;
  const loadingIndicator = document.getElementById("loading");
  const resultsTable = document.getElementById("results").getElementsByTagName("tbody")[0];

  loadingIndicator.style.display = "inline"; // Show loading indicator
  resultsTable.innerHTML = ""; // Clear previous results

  try {
    const payload = { survey_id };

    const res = await fetch("/api/lookup", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const surveys = await res.json();

    if (Array.isArray(surveys)) {
      surveys.forEach(s => {
        const row = resultsTable.insertRow();

        const nameCell = row.insertCell(0);
        const pathCell = row.insertCell(1);
        const statusCell = row.insertCell(2);
        const actionCell = row.insertCell(3);

        nameCell.textContent = s.title || "N/A";
        pathCell.textContent = s.path || "N/A";
        statusCell.textContent = s.state || "N/A";

        const exportButton = document.createElement("button");
        exportButton.textContent = "Export word document survey draft";
        exportButton.onclick = () => exportSurvey(s.path.split("/")[2]);
        actionCell.appendChild(exportButton);
      });
    } else {
      const row = resultsTable.insertRow();
      const cell = row.insertCell(0);
      cell.colSpan = 4;
      cell.textContent = surveys.error || "Unexpected response format";
      cell.style.color = "red";
    }
  } catch (error) {
    console.error("Error during lookup:", error);
    alert("An error occurred while searching. Please try again.");
  } finally {
    loadingIndicator.style.display = "none"; // Hide loading indicator
  }
}

async function exportSurvey(surveyId) {
  const res = await fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ survey_id: surveyId })
  });

  const blob = await res.blob();
  const url = window.URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = `survey_${surveyId}.docx`;
  document.body.appendChild(a);
  a.click();
  a.remove();

  window.URL.revokeObjectURL(url);
}