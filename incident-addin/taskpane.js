let currentIncident = null;

// ✅ INIT (WORKS in both Outlook + Browser)
function initApp() {
  console.log("🚀 App Initialized");

  const bind = (id, fn) => {
    const el = document.getElementById(id);
    if (el) el.onclick = fn;
  };

  bind("toggleTheme", () => {
    document.body.classList.toggle("dark");
  });

  bind("btnCreate", createIncidentPreview);
  bind("btnFollowUp", followUp);
  bind("btnClose", closeIncident);
  bind("btnFinalCreate", finalCreateIncident);
}

// 👉 Safe Office init
if (typeof Office !== "undefined") {
  Office.onReady(() => {
    console.log("✅ Office ready");
    initApp();
  });
} else {
  console.warn("⚠️ Running in browser mode (no Office)");
  window.onload = initApp;
}

// 🔔 TOAST
function showToast(message) {
  const toast = document.getElementById("toast");
  if (!toast) return;

  toast.innerText = message;
  toast.classList.add("show");

  setTimeout(() => {
    toast.classList.remove("show");
  }, 3000);
}

// 🔄 LOADER
function showLoader(show) {
  const loader = document.getElementById("loader");
  if (!loader) return;
  loader.classList.toggle("hidden", !show);
}

// 📅 FORMAT DATE
function formatDate(date) {
  try {
    return new Date(date).toLocaleString();
  } catch {
    return date;
  }
}

// 📩 GET EMAIL DETAILS (SAFE)
async function getEmailDetails() {
  try {
    // Browser mode fallback
    if (typeof Office === "undefined") {
      console.warn("⚠️ No Office context");
      return {
        subject: "Test Subject (Browser Mode)",
        from: "test@example.com",
        date: new Date(),
        body: "This is test data (no Outlook context)",
      };
    }

    const mailbox = Office.context.mailbox;

    if (!mailbox || !mailbox.item) {
      return {
        subject: "No Email Open",
        from: "unknown",
        date: new Date(),
        body: "No email selected",
      };
    }

    const item = mailbox.item;

    return new Promise((resolve) => {
      item.body.getAsync("text", (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          resolve({
            subject: item.subject || "No Subject",
            from: item.from?.emailAddress || "Unknown",
            date: item.dateTimeCreated,
            body: "Unable to read email body",
          });
        } else {
          resolve({
            subject: item.subject || "No Subject",
            from: item.from?.emailAddress || "Unknown",
            date: item.dateTimeCreated,
            body: result.value || "No description",
          });
        }
      });
    });
  } catch (err) {
    console.error("❌ Email fetch error:", err);
    return {
      subject: "Error fetching email",
      from: "unknown",
      date: new Date(),
      body: "Something went wrong",
    };
  }
}

// 🔢 GENERATE INCIDENT NUMBER
function generateIncidentNumber() {
  return "INC" + Math.floor(100000 + Math.random() * 900000);
}

// 📦 BUILD INCIDENT
async function buildIncident() {
  const email = await getEmailDetails();
  const incidentNumber = generateIncidentNumber();

  return {
    incidentNumber: incidentNumber,
    subject: "Incident " + incidentNumber + ": " + email.subject,
    reportedTime: formatDate(email.date),
    openedBy: email.from,
    description: email.body
  };
}

// 👀 PREVIEW INCIDENT
async function createIncidentPreview() {
  try {
    showLoader(true);

    const incident = await buildIncident();
    currentIncident = incident;

    document.getElementById("incNumber").innerText = incident.incidentNumber || "-";
    document.getElementById("incSubject").innerText = incident.subject || "-";
    document.getElementById("incTime").innerText = incident.reportedTime || "-";
    document.getElementById("incUser").innerText = incident.openedBy || "-";
    document.getElementById("incDescription").value = incident.description || "";

    showToast("Preview ready ✅");

  } catch (err) {
    console.error("❌ Preview error:", err);
    showToast("Error preparing preview ❌");
  } finally {
    showLoader(false);
  }
}

// 🚀 FINAL CREATE (n8n API)
async function finalCreateIncident() {
  try {
    if (!currentIncident) {
      showToast("Click 'Create Incident' first ⚠️");
      return;
    }

    showLoader(true);

    currentIncident.description =
      document.getElementById("incDescription").value;

    const response = await fetch(
      "https://datahubin.app.n8n.cloud/webhook/create-incident",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(currentIncident),
      }
    );

    const data = await response.json();

    showToast("Incident " + data.incidentNumber + " created 🎉");

    // Reset UI
    currentIncident = null;
    document.getElementById("incNumber").innerText = "-";
    document.getElementById("incSubject").innerText = "-";
    document.getElementById("incTime").innerText = "-";
    document.getElementById("incUser").innerText = "-";
    document.getElementById("incDescription").value = "";

  } catch (err) {
    console.error("❌ Create error:", err);
    showToast("Error creating incident ❌");
  } finally {
    showLoader(false);
  }
}

// 🤖 FOLLOW-UP
async function followUp() {
  showLoader(true);

  setTimeout(() => {
    document.getElementById("summaryBox").value =
      "AI-generated follow-up summary will appear here...";
    showToast("Follow-up ready ✉️");
    showLoader(false);
  }, 1200);
}

// ✅ CLOSE INCIDENT
async function closeIncident() {
  showLoader(true);

  setTimeout(() => {
    document.getElementById("summaryBox").value =
      "AI-generated closure summary will appear here...";
    showToast("Closure summary ready ✅");
    showLoader(false);
  }, 1200);
}