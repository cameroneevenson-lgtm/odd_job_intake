/*
 * Job Intake task pane.
 *
 * Reads the open message's DXF/PDF attachments, asks for a job number, and
 * POSTs to the loopback listener. It holds no intake logic of its own - the
 * server runs the same job_intake_service.create_intake the desktop app uses.
 */
(function () {
  "use strict";

  var CONFIG = window.JOB_INTAKE_CONFIG || {};
  var ALLOWED = /\.(dxf|pdf)$/i;
  var DXF = /\.dxf$/i;

  var attachments = [];   // [{id, name, size}]
  var labelRequired = false;
  var checkTimer = null;

  function $(id) { return document.getElementById(id); }

  function show(el, visible) { el.hidden = !visible; }

  function setError(message) {
    var box = $("error");
    box.textContent = message || "";
    show(box, Boolean(message));
  }

  function api(path, options) {
    options = options || {};
    options.headers = Object.assign(
      { "Authorization": "Bearer " + CONFIG.token },
      options.headers || {}
    );
    return fetch(CONFIG.apiBase + path, options).then(function (response) {
      return response.json()
        .catch(function () { return {}; })
        .then(function (body) {
          return { ok: response.ok, status: response.status, body: body };
        });
    });
  }

  // --- reading the message's attachments -------------------------------------

  function loadAttachments() {
    var item = Office.context.mailbox.item;
    var all = (item && item.attachments) || [];
    attachments = all.filter(function (a) {
      // Inline images and embedded items aren't job files; skip them by name.
      return ALLOWED.test(a.name || "");
    }).map(function (a) {
      return { id: a.id, name: a.name, size: a.size };
    });

    var list = $("attachment-list");
    list.innerHTML = "";
    attachments.forEach(function (a) {
      var li = document.createElement("li");
      li.textContent = a.name;
      list.appendChild(li);
    });

    var dxfCount = attachments.filter(function (a) { return DXF.test(a.name); }).length;
    var summary = $("attachment-summary");

    if (!attachments.length) {
      summary.textContent = "This email has no .dxf or .pdf attachments.";
      return false;
    }
    if (!dxfCount) {
      summary.textContent =
        "No .dxf attachment found - there would be nothing to nest.";
      return false;
    }
    summary.textContent =
      dxfCount + " DXF" + (dxfCount === 1 ? "" : "s") +
      " and " + (attachments.length - dxfCount) + " reference file(s) will be sent.";
    return true;
  }

  /* Office hands back attachment bytes one call at a time, so these are
   * fetched in sequence rather than in parallel - a job's attachment count is
   * small, and serial calls keep the failure attributable to one file. */
  function fetchAttachmentContent(index, collected) {
    if (index >= attachments.length) {
      return Promise.resolve(collected);
    }
    var attachment = attachments[index];
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.getAttachmentContentAsync(
        attachment.id,
        function (result) {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(
              "Could not read " + attachment.name + ": " +
              ((result.error && result.error.message) || "unknown error")
            ));
            return;
          }
          var content = result.value;
          if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
            reject(new Error(
              attachment.name + " came back as " + content.format +
              " rather than base64; it can't be filed."
            ));
            return;
          }
          collected.push({ name: attachment.name, contentBytes: content.content });
          resolve(collected);
        }
      );
    }).then(function (acc) {
      return fetchAttachmentContent(index + 1, acc);
    });
  }

  // --- job number: does it need a Label? -------------------------------------

  function checkJobNumber() {
    var jobNumber = $("job-number").value.trim().toUpperCase();
    var hint = $("job-hint");
    if (!jobNumber) {
      hint.textContent = "";
      show($("label-row"), false);
      labelRequired = false;
      return;
    }
    api("/api/job-intake/check?job_number=" + encodeURIComponent(jobNumber))
      .then(function (response) {
        if (!response.ok) {
          hint.textContent = response.body.error || "That job number wasn't accepted.";
          show($("label-row"), false);
          labelRequired = false;
          return;
        }
        labelRequired = Boolean(response.body.label_required);
        show($("label-row"), labelRequired);
        hint.textContent = labelRequired
          ? ""
          : jobNumber + " is a fresh job folder.";
      })
      .catch(function () {
        hint.textContent = "Couldn't reach the shop app to check that number.";
      });
  }

  // --- submit -----------------------------------------------------------------

  function submit(event) {
    event.preventDefault();
    setError("");

    var jobNumber = $("job-number").value.trim().toUpperCase();
    var label = $("job-label").value.trim();
    if (!jobNumber) {
      setError("Enter a job number.");
      return;
    }
    if (labelRequired && !label) {
      setError("This job already exists, so a Label is required.");
      return;
    }

    var button = $("submit");
    button.disabled = true;
    button.textContent = "Sending…";

    var item = Office.context.mailbox.item;
    fetchAttachmentContent(0, [])
      .then(function (payloadAttachments) {
        return api("/api/job-intake", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            job_number: jobNumber,
            label: label,
            attachments: payloadAttachments,
            email_subject: (item && item.subject) || "",
            email_sender: (item && item.from && item.from.emailAddress) || ""
          })
        });
      })
      .then(function (response) {
        if (!response.ok) {
          setError(response.body.error || ("The intake failed (HTTP " + response.status + ")."));
          button.disabled = false;
          button.textContent = "Send to Job Intake";
          return;
        }
        showResult(response.body);
      })
      .catch(function (err) {
        setError(err.message || "The intake could not be sent.");
        button.disabled = false;
        button.textContent = "Send to Job Intake";
      });
  }

  function showResult(entry) {
    show($("intake-form"), false);
    var box = $("result");
    var unmatched = entry.po_unmatched || [];

    var html =
      "<h2>Filed</h2>" +
      "<p><strong>" + escapeHtml(entry.job_number) +
      (entry.label ? " / " + escapeHtml(entry.label) : "") + "</strong></p>" +
      "<dl>" +
      row("PO number", entry.po_number || "not found") +
      row("Due date", entry.due_date || "not found") +
      row("Parts", String(entry.parts)) +
      row("Folder", entry.job_folder) +
      "</dl>";

    if (unmatched.length) {
      html += "<p class=\"warn\">PO lines that matched no attached DXF:</p><ul>";
      unmatched.forEach(function (line) {
        html += "<li>" + escapeHtml(line) + "</li>";
      });
      html += "</ul>";
    }

    /* Deliberately explicit: no RADAN work has happened. The user still has to
     * set material and thickness in the desktop app, which is the whole
     * reason this doesn't try to finish the job. */
    html +=
      "<p class=\"muted\">Now open the Job Intake tab in the shop app to set " +
      "material and thickness, then create the RPD and import parts.</p>";

    box.innerHTML = html;
    show(box, true);
  }

  function row(label, value) {
    return "<dt>" + escapeHtml(label) + "</dt><dd>" + escapeHtml(String(value)) + "</dd>";
  }

  function escapeHtml(text) {
    var div = document.createElement("div");
    div.textContent = text == null ? "" : String(text);
    return div.innerHTML;
  }

  // --- startup ----------------------------------------------------------------

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      $("boot").textContent = "This pane only runs in Outlook.";
      return;
    }

    // Confirm the desktop app is actually up before showing a form that can
    // only fail without it. /api/health needs no token, so a failure here
    // means "not running" rather than "bad token".
    fetch(CONFIG.apiBase + "/api/health")
      .then(function (response) {
        if (!response.ok) { throw new Error("unhealthy"); }
        show($("boot"), false);
        show($("intake-form"), true);
        var hasWork = loadAttachments();
        $("submit").disabled = !hasWork;
        var jobInput = $("job-number");
        jobInput.addEventListener("input", function () {
          window.clearTimeout(checkTimer);
          checkTimer = window.setTimeout(checkJobNumber, 400);
        });
        $("intake-form").addEventListener("submit", submit);
      })
      .catch(function () {
        $("boot").textContent =
          "The shop app isn't running on this PC. Open the Ops Suite (or Odd Job " +
          "Intake), then reopen this pane.";
      });
  });
})();
