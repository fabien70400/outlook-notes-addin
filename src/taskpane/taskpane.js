/* global Office, document, localStorage */

const NOTE_PROPERTY = "tandemFollowUpNote";
const SAVED_AT_PROPERTY = "tandemFollowUpSavedAt";
const RECAP_STORAGE_KEY = "tandemFollowUpRecapIndex";

let noteTextEl, saveButtonEl, clearButtonEl, statusEl, subjectEl, fromEl;
let tabNoteEl, tabRecapEl, panelNoteEl, panelRecapEl, recapListEl, recapEmptyEl, refreshRecapButtonEl;
let customProps = null;
let currentItemId = null;

Office.onReady(() => {
  noteTextEl = document.getElementById("note-text");
  saveButtonEl = document.getElementById("save-button");
  clearButtonEl = document.getElementById("clear-button");
  statusEl = document.getElementById("status");
  subjectEl = document.getElementById("mail-subject");
  fromEl = document.getElementById("mail-from");

  tabNoteEl = document.getElementById("tab-note");
  tabRecapEl = document.getElementById("tab-recap");
  panelNoteEl = document.getElementById("panel-note");
  panelRecapEl = document.getElementById("panel-recap");
  recapListEl = document.getElementById("recap-list");
  recapEmptyEl = document.getElementById("recap-empty");
  refreshRecapButtonEl = document.getElementById("refresh-recap-button");

  saveButtonEl.addEventListener("click", onSave);
  clearButtonEl.addEventListener("click", onClear);
  tabNoteEl.addEventListener("click", () => showTab("note"));
  tabRecapEl.addEventListener("click", () => showTab("recap"));
  refreshRecapButtonEl.addEventListener("click", renderRecapList);

  loadItem();

  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadItem);
});

function loadItem() {
  document.getElementById("loading").classList.remove("hidden");
  document.getElementById("app").classList.add("hidden");
  setStatus("");
  showTab("note");

  const item = Office.context.mailbox.item;
  currentItemId = item.itemId || null;
  subjectEl.textContent = item.subject || "(sans objet)";
  fromEl.textContent = item.from ? `De : ${item.from.displayName} <${item.from.emailAddress}>` : "";

  item.loadCustomPropertiesAsync((result) => {
    document.getElementById("loading").classList.add("hidden");
    document.getElementById("app").classList.remove("hidden");

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Impossible de charger la note existante.", true);
      return;
    }

    customProps = result.value;
    const existingNote = customProps.get(NOTE_PROPERTY) || "";
    const savedAt = customProps.get(SAVED_AT_PROPERTY);
    noteTextEl.value = existingNote;
    setSavedStatus(savedAt);
  });
}

function showTab(tabName) {
  const isNote = tabName === "note";
  tabNoteEl.classList.toggle("active", isNote);
  tabRecapEl.classList.toggle("active", !isNote);
  panelNoteEl.classList.toggle("hidden", !isNote);
  panelRecapEl.classList.toggle("hidden", isNote);

  if (!isNote) {
    renderRecapList();
  }
}

function onSave() {
  if (!customProps) {
    return;
  }
  const value = noteTextEl.value.trim();
  const now = new Date().toISOString();

  customProps.set(NOTE_PROPERTY, value);
  customProps.set(SAVED_AT_PROPERTY, now);

  saveButtonEl.disabled = true;
  customProps.saveAsync((result) => {
    saveButtonEl.disabled = false;
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Erreur lors de l'enregistrement.", true);
      return;
    }
    setSavedStatus(now);
    upsertRecapEntry(value, now);
  });
}

function onClear() {
  if (!customProps) {
    return;
  }
  noteTextEl.value = "";
  customProps.set(NOTE_PROPERTY, "");
  customProps.set(SAVED_AT_PROPERTY, "");

  customProps.saveAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Erreur lors de la suppression.", true);
      return;
    }
    setStatus("Note effacee.");
    removeRecapIndexEntry(currentItemId);
  });
}

function setSavedStatus(isoDate) {
  if (!isoDate) {
    setStatus("");
    return;
  }
  const d = new Date(isoDate);
  const formatted = d.toLocaleString("fr-FR");
  setStatus(`Enregistre le ${formatted}`);
}

function setStatus(message, isError) {
  statusEl.textContent = message;
  statusEl.classList.toggle("error", Boolean(isError));
}

function readRecapIndex() {
  try {
    const raw = localStorage.getItem(RECAP_STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function writeRecapIndex(entries) {
  localStorage.setItem(RECAP_STORAGE_KEY, JSON.stringify(entries));
}

function upsertRecapEntry(note, savedAt) {
  if (!currentItemId) {
    return;
  }
  const entries = readRecapIndex().filter((e) => e.itemId !== currentItemId);

  if (note) {
    entries.push({
      itemId: currentItemId,
      subject: subjectEl.textContent,
      from: fromEl.textContent,
      note,
      savedAt,
    });
  }

  writeRecapIndex(entries);
}

function removeRecapIndexEntry(itemId) {
  if (!itemId) {
    return;
  }
  const entries = readRecapIndex().filter((e) => e.itemId !== itemId);
  writeRecapIndex(entries);
}

function deleteFromRecap(itemId) {
  removeRecapIndexEntry(itemId);

  if (itemId === currentItemId && customProps) {
    noteTextEl.value = "";
    customProps.set(NOTE_PROPERTY, "");
    customProps.set(SAVED_AT_PROPERTY, "");
    customProps.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setStatus("Note effacee.");
      }
    });
  }

  renderRecapList();
}

function renderRecapList() {
  const entries = readRecapIndex().sort((a, b) => (a.savedAt < b.savedAt ? 1 : -1));

  recapListEl.innerHTML = "";
  recapEmptyEl.classList.toggle("hidden", entries.length > 0);

  entries.forEach((entry) => {
    const li = document.createElement("li");
    li.className = "recap-item";

    const header = document.createElement("div");
    header.className = "recap-item-header";

    const subject = document.createElement("div");
    subject.className = "recap-item-subject";
    subject.textContent = entry.subject || "(sans objet)";

    const deleteButton = document.createElement("button");
    deleteButton.className = "recap-item-delete";
    deleteButton.setAttribute("aria-label", "Supprimer la note");
    deleteButton.setAttribute("title", "Supprimer la note");
    deleteButton.textContent = "×";
    deleteButton.addEventListener("click", (event) => {
      event.stopPropagation();
      deleteFromRecap(entry.itemId);
    });

    header.appendChild(subject);
    header.appendChild(deleteButton);

    const from = document.createElement("div");
    from.className = "recap-item-from";
    from.textContent = entry.from || "";

    const note = document.createElement("div");
    note.className = "recap-item-note";
    note.textContent = entry.note;

    const date = document.createElement("div");
    date.className = "recap-item-date";
    date.textContent = entry.savedAt ? new Date(entry.savedAt).toLocaleString("fr-FR") : "";

    li.appendChild(header);
    li.appendChild(from);
    li.appendChild(note);
    li.appendChild(date);

    li.addEventListener("click", () => {
      Office.context.mailbox.displayMessageForm(entry.itemId);
    });

    recapListEl.appendChild(li);
  });
}
