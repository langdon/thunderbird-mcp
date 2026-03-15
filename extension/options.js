/* global browser */
"use strict";

const statusDot = document.getElementById("statusDot");
const statusText = document.getElementById("statusText");
const serverPort = document.getElementById("serverPort");
const connFile = document.getElementById("connFile");
const buildInfo = document.getElementById("buildInfo");
const accountList = document.getElementById("accountList");
const saveBtn = document.getElementById("saveBtn");
const saveStatus = document.getElementById("saveStatus");

const toolList = document.getElementById("toolList");
const saveToolsBtn = document.getElementById("saveToolsBtn");
const saveToolsStatus = document.getElementById("saveToolsStatus");

let currentAccounts = [];
let currentTools = [];

const TOOL_GROUPS = [
  { label: "Messages", tools: ["searchMessages", "getMessage", "getRecentMessages", "displayMessage", "sendMail", "replyToMessage", "forwardMessage", "updateMessage", "deleteMessages"] },
  { label: "Folders", tools: ["createFolder", "renameFolder", "moveFolder", "deleteFolder", "emptyTrash", "emptyJunk"] },
  { label: "Contacts", tools: ["searchContacts", "createContact", "updateContact", "deleteContact"] },
  { label: "Calendar", tools: ["listCalendars", "listEvents", "createEvent", "updateEvent", "deleteEvent", "createTask"] },
  { label: "Filters", tools: ["listFilters", "createFilter", "updateFilter", "deleteFilter", "reorderFilters", "applyFilters"] },
  { label: "System", tools: ["listAccounts", "listFolders", "getAccountAccess"] },
];

async function loadServerInfo() {
  try {
    const info = await browser.mcpServer.getServerInfo();
    if (info.running) {
      statusDot.className = "status-dot running";
      statusText.textContent = "Running";
      serverPort.textContent = info.port || "--";
      connFile.textContent = info.connectionFile || "--";
    } else {
      statusDot.className = "status-dot stopped";
      statusText.textContent = "Not running";
      serverPort.textContent = "--";
      connFile.textContent = "--";
    }
    if (info.buildVersion) {
      // Parse git describe: "v0.2.0-7-g1461f1a+dirty" → tag, commits, hash, dirty
      const m = info.buildVersion.match(/^(v[\d.]+)(?:-(\d+)-g([0-9a-f]+))?(\+dirty)?$/);
      let display;
      if (m) {
        const [, tag, commits, hash, dirty] = m;
        display = tag;
        if (commits && commits !== "0") display += ` +${commits}`;
        display += ` (${hash || tag})`;
        if (dirty) {
          display += " +dirty";
          if (info.buildDate) {
            display += " " + info.buildDate.replace("T", " ").replace(/\.\d+Z$/, " UTC");
          }
        }
      } else {
        display = info.buildVersion;
      }
      buildInfo.textContent = display;
    } else {
      buildInfo.textContent = "--";
    }
  } catch (e) {
    statusDot.className = "status-dot stopped";
    statusText.textContent = "Error: " + e.message;
  }
}

async function loadAccountAccess() {
  try {
    const data = await browser.mcpServer.getAccountAccessConfig();
    currentAccounts = data.accounts || [];

    if (currentAccounts.length === 0) {
      accountList.innerHTML = "<li>No accounts found.</li>";
      return;
    }

    accountList.innerHTML = "";
    for (const acct of currentAccounts) {
      const li = document.createElement("li");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.id = "acct-" + acct.id;
      checkbox.value = acct.id;
      checkbox.checked = acct.allowed;
      checkbox.addEventListener("change", onAccountChange);

      const label = document.createElement("label");
      label.htmlFor = checkbox.id;
      label.textContent = acct.name;

      const typeSpan = document.createElement("span");
      typeSpan.className = "account-type";
      typeSpan.textContent = acct.type;
      label.appendChild(typeSpan);

      li.appendChild(checkbox);
      li.appendChild(label);
      accountList.appendChild(li);
    }

    saveBtn.disabled = false;
    saveStatus.textContent = "";
  } catch (e) {
    accountList.innerHTML = "<li>Error loading accounts: " + e.message + "</li>";
  }
}

function onAccountChange() {
  saveStatus.textContent = "";
}

saveBtn.addEventListener("click", async () => {
  saveBtn.disabled = true;
  saveStatus.textContent = "";
  saveStatus.className = "save-status";

  const checkboxes = accountList.querySelectorAll('input[type="checkbox"]');
  const checked = [];
  let allChecked = true;
  for (const cb of checkboxes) {
    if (cb.checked) {
      checked.push(cb.value);
    } else {
      allChecked = false;
    }
  }

  // If all are checked, send empty array (= allow all)
  const allowedIds = allChecked ? [] : checked;

  try {
    const result = await browser.mcpServer.setAccountAccess(allowedIds);
    if (result.error) {
      saveStatus.textContent = result.error;
      saveStatus.className = "save-status error";
    } else {
      saveStatus.textContent = "Saved.";
      // Reload to reflect updated state
      await loadAccountAccess();
    }
  } catch (e) {
    saveStatus.textContent = "Error: " + e.message;
    saveStatus.className = "save-status error";
  }
  saveBtn.disabled = false;
});

async function loadToolAccess() {
  try {
    const data = await browser.mcpServer.getToolAccessConfig();
    currentTools = data.tools || [];

    if (currentTools.length === 0) {
      toolList.innerHTML = "<li>No tools found.</li>";
      return;
    }

    // Index tools by name for quick lookup
    const toolMap = {};
    for (const tool of currentTools) {
      toolMap[tool.name] = tool;
    }

    toolList.innerHTML = "";

    // Collect all grouped tool names to detect ungrouped ones
    const groupedTools = new Set(TOOL_GROUPS.flatMap(g => g.tools));

    // Build the groups to render, appending "Other" if any tools are ungrouped
    const ungrouped = currentTools.filter(t => !groupedTools.has(t.name)).map(t => t.name);
    const groups = ungrouped.length > 0
      ? [...TOOL_GROUPS, { label: "Other", tools: ungrouped }]
      : TOOL_GROUPS;

    for (const group of groups) {
      // Group header
      const header = document.createElement("li");
      header.className = "tool-group-header";
      header.textContent = group.label;
      toolList.appendChild(header);

      // Tools in this group
      for (const toolName of group.tools) {
        const tool = toolMap[toolName];
        if (!tool) continue;

        const li = document.createElement("li");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.id = "tool-" + tool.name;
        checkbox.value = tool.name;
        checkbox.checked = tool.enabled;
        if (tool.undisableable) {
          checkbox.disabled = true;
        }
        checkbox.addEventListener("change", () => {
          saveToolsStatus.textContent = "";
        });

        const label = document.createElement("label");
        label.htmlFor = checkbox.id;
        label.textContent = tool.name;
        if (tool.undisableable) {
          const lockSpan = document.createElement("span");
          lockSpan.className = "account-type";
          lockSpan.textContent = "required";
          label.appendChild(lockSpan);
        }

        li.appendChild(checkbox);
        li.appendChild(label);
        toolList.appendChild(li);
      }
    }

    saveToolsBtn.disabled = false;
    saveToolsStatus.textContent = "";
  } catch (e) {
    toolList.innerHTML = "<li>Error loading tools: " + e.message + "</li>";
  }
}

saveToolsBtn.addEventListener("click", async () => {
  saveToolsBtn.disabled = true;
  saveToolsStatus.textContent = "";
  saveToolsStatus.className = "save-status";

  const checkboxes = toolList.querySelectorAll('input[type="checkbox"]');
  const disabled = [];
  for (const cb of checkboxes) {
    if (!cb.checked && !cb.disabled) {
      disabled.push(cb.value);
    }
  }

  try {
    const result = await browser.mcpServer.setToolAccess(disabled);
    if (result.error) {
      saveToolsStatus.textContent = result.error;
      saveToolsStatus.className = "save-status error";
    } else {
      saveToolsStatus.textContent = "Saved.";
      await loadToolAccess();
    }
  } catch (e) {
    saveToolsStatus.textContent = "Error: " + e.message;
    saveToolsStatus.className = "save-status error";
  }
  saveToolsBtn.disabled = false;
});

loadServerInfo();
loadAccountAccess();
loadToolAccess();
