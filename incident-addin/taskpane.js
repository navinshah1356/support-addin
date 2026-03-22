let currentIncident = null;

// ✅ INIT
function initApp() {
  console.log("🚀 App Initialized");

  const bind = (id, fn) => {
    const el = document.getElementById(id);
    if (el) el.onclick = fn;
  };

  bind("toggleTheme", () => {
    document.body.classList.toggle("dark");
  });

  // ✅ FIXED: Only ONE binding per button
  bind("btnCreate", handleCreateClick);
  bind("btnFollowUp", showLogDetailsView);
  bind("btnClose", showCloseView);

  bind("btnFinalCreate", finalCreateIncident);
  bind("btnCloseFinal", closeIncidentFinal);
  bind("btnFetchIncident", fetchIncidentForClose);
}

// 👉 Office init
if (typeof Office !== "undefined") {
  Office.onReady(() => initApp());
} else {
  window.onload = initApp;
}

// ================= UI CONTROL =================

function hideAllSections() {
  document.getElementById("aiSection").classList.add("hidden");
  document.getElementById("previewSection").classList.add("hidden");
  document.getElementById("closeSection").classList.add("hidden");
}

// ✅ CREATE VIEW + PREVIEW
async function handleCreateClick() {
  hideAllSections();

  document.getElementById("aiSection").classList.remove("hidden");
  document.getElementById("previewSection").classList.remove("hidden");

  await createIncidentPreview(); // 🔥 important
}

// ✅ LOG DETAILS (renamed from follow-up)
function showLogDetailsView() {
  hideAllSections();

  document.getElementById("aiSection").classList.remove("hidden");

  document.getElementById("summaryBox").value =
    "📝 Log details / notes can be captured here...";
}

// ✅ CLOSE VIEW
function showCloseView() {
  hideAllSections();
  document.getElementById("closeSection").classList.remove("hidden");
}

// ================= UTIL =================

function showToast(message) {
  const toast = document.getElementById("toast");
  if (!toast) return;

  toast.innerText = message;
  toast.classList.add("show");

  setTimeout(() => toast.classList.remove("show"), 3000);
}

function showLoader(show) {
  const loader = document.getElementById("loader");
  if (!loader) return;
  loader.classList.toggle("hidden", !show);
}

function formatDate(date) {
  try {
    return new Date(date).toISOString();
  } catch {
    return new Date().toISOString();
  }
}

// ================= EMAIL =================

async function getEmailDetails() {
  try {
    const item = Office.context.mailbox.item;

    return new Promise((resolve) => {
      item.body.getAsync("text", (result) => {
        resolve({
          subject: item.subject || "No Subject",
          from: item.from?.emailAddress || "Unknown",
          date: item.dateTimeCreated,
          body: result.value || "No description",
        });
      });
    });
  } catch {
    return {
      subject: "Test",
      from: "test@example.com",
      date: new Date(),
      body: "Test body",
    };
  }
}

// ================= CREATE =================

function generateIncidentNumber() {
  return "INC" + Math.floor(100000 + Math.random() * 900000);
}

async function buildIncident() {
  const email = await getEmailDetails();
  const incidentNumber = generateIncidentNumber();

  return {
    action: "create",
    incidentNumber,
    subject: "Incident " + incidentNumber + ": " + email.subject,
    reportedTime: formatDate(email.date),
    openedBy: email.from,
    description: email.body
  };
}

async function createIncidentPreview() {
  showLoader(true);

  const incident = await buildIncident();
  currentIncident = incident;

  document.getElementById("incNumber").innerText = incident.incidentNumber;
  document.getElementById("incSubject").innerText = incident.subject;
  document.getElementById("incTime").innerText = incident.reportedTime;
  document.getElementById("incUser").innerText = incident.openedBy;
  document.getElementById("incDescription").value = incident.description;

  showToast("Preview ready ✅");
  showLoader(false);
}

async function finalCreateIncident() {
  if (!currentIncident) {
    showToast("Click Create first ⚠️");
    return;
  }

  showLoader(true);

  currentIncident.description =
    document.getElementById("incDescription").value;

  try {
    const response = await fetch("https://prod-12.eastasia.logic.azure.com:443/workflows/03e1e813fba5481a945dd1ec560aa754/triggers/When_an_HTTP_request_is_received/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=hQl8F9EEKZEuLDNJJyxAsUIf5UbGu1AKOKNKVK3aANU", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(currentIncident),
    });

    if (response.ok) showToast("Incident created 🎉");
    else showToast("Create failed ❌");

  } catch {
    showToast("Error creating incident ❌");
  }

  showLoader(false);
}

// ================= FETCH =================

async function fetchIncidentForClose() {
  const incId = document.getElementById("closeIncidentId").value;

  if (!incId) {
    showToast("Enter Incident Number ⚠️");
    return;
  }

  showLoader(true);

  const payload = {
    action: "fetch",
    incidentNumber: incId
  };

  try {
    const res = await fetch("https://prod-12.eastasia.logic.azure.com:443/workflows/03e1e813fba5481a945dd1ec560aa754/triggers/When_an_HTTP_request_is_received/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=hQl8F9EEKZEuLDNJJyxAsUIf5UbGu1AKOKNKVK3aANU", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const result = await res.json();

    if (!result.found) {
      showToast("❌ Incident not found");
      return;
    }

    const data = result.data;

    document.getElementById("incNumber").innerText = data.IncidentNumber;
    document.getElementById("incSubject").innerText = data.Subject;
    document.getElementById("incTime").innerText = data.ReportedTime;
    document.getElementById("incUser").innerText = data.OpenedBy;
    document.getElementById("incDescription").value = data.Description;

    showToast("Incident loaded ✅");

  } catch {
    showToast("Error fetching incident ❌");
  }

  showLoader(false);
}

// ================= CLOSE =================

async function closeIncidentFinal() {
  const incId = document.getElementById("closeIncidentId").value;
  const rootCause = document.getElementById("rootCause").value;
  const resolution = document.getElementById("resolution").value;

  if (!incId || !rootCause || !resolution) {
    showToast("Fill all fields ⚠️");
    return;
  }

  showLoader(true);

  const payload = {
    action: "close",
    incidentNumber: incId,
    rootCause,
    resolution,
    closedBy: "navinshah1356@outlook.com"
  };

  try {
    const res = await fetch("https://prod-12.eastasia.logic.azure.com:443/workflows/03e1e813fba5481a945dd1ec560aa754/triggers/When_an_HTTP_request_is_received/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=hQl8F9EEKZEuLDNJJyxAsUIf5UbGu1AKOKNKVK3aANU", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    if (res.ok) showToast("Incident closed ✅");
    else showToast("Close failed ❌");

  } catch {
    showToast("Error closing incident ❌");
  }

  showLoader(false);
}
