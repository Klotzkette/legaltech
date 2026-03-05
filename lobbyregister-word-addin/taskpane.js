/* ============================================================
   Lobbyregister Word Add-in - Taskpane Logic
   ============================================================ */

(function () {
  "use strict";

  // --------------- Configuration ---------------
  const API_BASE = "https://www.lobbyregister.bundestag.de";
  const API_KEY = "5bHB2zrUuHR6YdPoZygQhWfg2CBrjUOi";
  const PAGE_SIZE = 20;

  // --------------- State ---------------
  let state = {
    results: [],
    totalCount: 0,
    currentPage: 0,
    totalPages: 0,
    query: "",
    sort: "NAME_ASC",
    filters: {},
    selectedEntry: null,
    isLoading: false,
  };

  // --------------- DOM References ---------------
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  const dom = {};
  function cacheDom() {
    dom.searchInput = $("#search-input");
    dom.btnSearch = $("#btn-search");
    dom.btnClear = $("#btn-clear-search");
    dom.btnToggleFilters = $("#btn-toggle-filters");
    dom.filtersPanel = $("#filters-panel");
    dom.filterActivity = $("#filter-activity");
    dom.filterInterest = $("#filter-interest");
    dom.filterLegalform = $("#filter-legalform");
    dom.filterMembersFrom = $("#filter-members-from");
    dom.filterMembersTo = $("#filter-members-to");
    dom.filterSort = $("#filter-sort");
    dom.btnApplyFilters = $("#btn-apply-filters");
    dom.btnResetFilters = $("#btn-reset-filters");
    dom.resultsInfo = $("#results-info");
    dom.resultsCount = $("#results-count");
    dom.resultsList = $("#results-list");
    dom.pagination = $("#pagination");
    dom.btnPrev = $("#btn-prev");
    dom.btnNext = $("#btn-next");
    dom.pageInfo = $("#page-info");
    dom.emptyState = $("#empty-state");
    dom.loading = $("#loading");
    dom.errorState = $("#error-state");
    dom.errorMessage = $("#error-message");
    dom.btnRetry = $("#btn-retry");
    dom.btnInsertAllTable = $("#btn-insert-all-table");
    dom.searchView = $("#search-view");
    dom.detailView = $("#detail-view");
    dom.btnBack = $("#btn-back");
    dom.detailContent = $("#detail-content");
    dom.btnInsertDetail = $("#btn-insert-detail");
    dom.toast = $("#toast");
    dom.toastMessage = $("#toast-message");
  }

  // --------------- Office.js Init ---------------
  let officeReady = false;

  Office.onReady(function (info) {
    officeReady = info.host === Office.HostType.Word;
    cacheDom();
    bindEvents();
  });

  // --------------- Event Binding ---------------
  function bindEvents() {
    dom.btnSearch.addEventListener("click", () => doSearch(0));
    dom.searchInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") doSearch(0);
    });
    dom.searchInput.addEventListener("input", () => {
      dom.btnClear.classList.toggle("hidden", !dom.searchInput.value);
    });
    dom.btnClear.addEventListener("click", () => {
      dom.searchInput.value = "";
      dom.btnClear.classList.add("hidden");
      dom.searchInput.focus();
    });
    dom.btnToggleFilters.addEventListener("click", toggleFilters);
    dom.btnApplyFilters.addEventListener("click", () => doSearch(0));
    dom.btnResetFilters.addEventListener("click", resetFilters);
    dom.btnPrev.addEventListener("click", () => doSearch(state.currentPage - 1));
    dom.btnNext.addEventListener("click", () => doSearch(state.currentPage + 1));
    dom.btnRetry.addEventListener("click", () => doSearch(state.currentPage));
    dom.btnBack.addEventListener("click", showSearchView);
    dom.btnInsertDetail.addEventListener("click", insertDetailIntoDocument);
    dom.btnInsertAllTable.addEventListener("click", insertAllResultsAsTable);
  }

  // --------------- Filter UI ---------------
  function toggleFilters() {
    const hidden = dom.filtersPanel.classList.toggle("hidden");
    dom.btnToggleFilters.querySelector("span").textContent = hidden
      ? "Filter anzeigen"
      : "Filter ausblenden";
  }

  function resetFilters() {
    dom.filterActivity.value = "";
    dom.filterInterest.value = "";
    dom.filterLegalform.value = "";
    dom.filterMembersFrom.value = "";
    dom.filterMembersTo.value = "";
    dom.filterSort.value = "NAME_ASC";
  }

  function collectFilters() {
    return {
      activity: dom.filterActivity.value,
      interest: dom.filterInterest.value,
      legalform: dom.filterLegalform.value,
      membersFrom: dom.filterMembersFrom.value ? parseInt(dom.filterMembersFrom.value, 10) : null,
      membersTo: dom.filterMembersTo.value ? parseInt(dom.filterMembersTo.value, 10) : null,
      sort: dom.filterSort.value,
    };
  }

  // --------------- API ---------------
  async function fetchFromApi(query, sort, page) {
    const url = new URL(API_BASE + "/sucheDetailJson");
    if (query) url.searchParams.set("q", query);
    if (sort) url.searchParams.set("sort", sort);
    url.searchParams.set("apikey", API_KEY);

    const response = await fetch(url.toString(), {
      method: "GET",
      headers: { Accept: "application/json" },
    });

    if (!response.ok) {
      throw new Error("API-Fehler: " + response.status + " " + response.statusText);
    }

    return response.json();
  }

  // --------------- Client-side filtering & pagination ---------------
  function applyFilters(results, filters) {
    return results.filter((item) => {
      const entry = item.registerEntryDetail;
      if (!entry) return false;

      // Activity filter
      if (filters.activity) {
        const activities = entry.activities || [];
        const match = activities.some((a) => a.code === filters.activity);
        if (!match) return false;
      }

      // Field of interest filter
      if (filters.interest) {
        const fois = entry.fieldsOfInterest || [];
        const match = fois.some((f) => f.code === filters.interest);
        if (!match) return false;
      }

      // Legal form type filter
      if (filters.legalform) {
        const identity = entry.lobbyistIdentity;
        if (!identity || !identity.legalForm || identity.legalForm.type !== filters.legalform) {
          return false;
        }
      }

      // Members range filter
      if (filters.membersFrom != null || filters.membersTo != null) {
        const identity = entry.lobbyistIdentity;
        const members = identity ? identity.members : null;
        if (members == null) return false;
        if (filters.membersFrom != null && members < filters.membersFrom) return false;
        if (filters.membersTo != null && members > filters.membersTo) return false;
      }

      return true;
    });
  }

  function paginate(results, page) {
    const start = page * PAGE_SIZE;
    return results.slice(start, start + PAGE_SIZE);
  }

  // --------------- Search ---------------
  async function doSearch(page) {
    if (state.isLoading) return;

    const query = dom.searchInput.value.trim();
    const filters = collectFilters();

    showLoading(true);
    hideError();

    try {
      const data = await fetchFromApi(query, filters.sort, page);
      const allResults = data.results || [];

      // Apply client-side filters
      const filtered = applyFilters(allResults, filters);

      state.query = query;
      state.filters = filters;
      state.results = filtered;
      state.totalCount = filtered.length;
      state.totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
      state.currentPage = Math.min(page, state.totalPages - 1);

      renderResults();
    } catch (err) {
      showError(err.message || "Ein unbekannter Fehler ist aufgetreten.");
    } finally {
      showLoading(false);
    }
  }

  // --------------- Rendering ---------------
  function renderResults() {
    const pageResults = paginate(state.results, state.currentPage);

    if (state.results.length === 0) {
      dom.resultsList.innerHTML = "";
      dom.resultsInfo.classList.add("hidden");
      dom.pagination.classList.add("hidden");
      dom.emptyState.classList.remove("hidden");
      dom.emptyState.querySelector("p").textContent =
        "Keine Ergebnisse gefunden. Passen Sie Ihre Suche oder Filter an.";
      return;
    }

    dom.emptyState.classList.add("hidden");
    dom.resultsInfo.classList.remove("hidden");
    dom.resultsCount.textContent =
      state.totalCount + " Ergebnis" + (state.totalCount !== 1 ? "se" : "") + " gefunden";

    dom.resultsList.innerHTML = pageResults.map(renderResultCard).join("");

    // Bind card click events
    dom.resultsList.querySelectorAll(".result-card").forEach((card) => {
      card.addEventListener("click", () => {
        const idx = parseInt(card.dataset.index, 10);
        showDetail(state.results[idx]);
      });
    });

    // Bind per-card insert buttons
    dom.resultsList.querySelectorAll(".btn-insert-single").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        const idx = parseInt(btn.dataset.index, 10);
        insertSingleEntry(state.results[idx]);
      });
    });

    // Pagination
    if (state.totalPages > 1) {
      dom.pagination.classList.remove("hidden");
      dom.btnPrev.disabled = state.currentPage === 0;
      dom.btnNext.disabled = state.currentPage >= state.totalPages - 1;
      dom.pageInfo.textContent =
        "Seite " + (state.currentPage + 1) + " von " + state.totalPages;
    } else {
      dom.pagination.classList.add("hidden");
    }
  }

  function renderResultCard(item, globalIndex) {
    const entry = item.registerEntryDetail || {};
    const identity = entry.lobbyistIdentity || {};
    const account = entry.account || {};
    const name = identity.name || "Unbekannt";
    const regNr = item.registerNumber || account.registerNumber || "-";
    const legalForm = identity.legalForm
      ? identity.legalForm.code_de || identity.legalForm.legalFormText || ""
      : "";
    const city = identity.address ? identity.address.city || "" : "";
    const members = identity.members != null ? formatNumber(identity.members) : "-";
    const expenses = formatExpenses(entry.financialExpensesEuro);
    const activities = (entry.activities || [])
      .map((a) => a.de || a.text || a.code)
      .slice(0, 2)
      .join(", ");

    return (
      '<div class="result-card" data-index="' + globalIndex + '" tabindex="0">' +
      '  <div class="card-header">' +
      '    <span class="card-reg">' + escHtml(regNr) + "</span>" +
      '    <button class="btn-icon btn-insert-single" data-index="' + globalIndex + '" title="In Dokument einfuegen">' +
      '      <i class="ms-Icon ms-Icon--AddToShoppingList"></i>' +
      "    </button>" +
      "  </div>" +
      '  <h3 class="card-name">' + escHtml(name) + "</h3>" +
      '  <div class="card-meta">' +
      (legalForm ? '<span class="tag">' + escHtml(legalForm) + "</span>" : "") +
      (city ? '<span class="tag">' + escHtml(city) + "</span>" : "") +
      "  </div>" +
      '  <div class="card-stats">' +
      '    <span><i class="ms-Icon ms-Icon--People"></i> ' + members + "</span>" +
      '    <span><i class="ms-Icon ms-Icon--Money"></i> ' + expenses + "</span>" +
      "  </div>" +
      (activities ? '<div class="card-activity">' + escHtml(activities) + "</div>" : "") +
      "</div>"
    );
  }

  // --------------- Detail View ---------------
  function showDetail(item) {
    state.selectedEntry = item;
    dom.searchView.classList.remove("active");
    dom.detailView.classList.add("active");
    renderDetail(item);
    dom.detailView.scrollTop = 0;
  }

  function showSearchView() {
    dom.detailView.classList.remove("active");
    dom.searchView.classList.add("active");
  }

  function renderDetail(item) {
    const entry = item.registerEntryDetail || {};
    const identity = entry.lobbyistIdentity || {};
    const account = entry.account || {};
    const address = identity.address || {};
    const legalForm = identity.legalForm || {};

    let html = "";

    // Header
    html +=
      '<div class="detail-header">' +
      '  <span class="detail-reg">' + escHtml(item.registerNumber || account.registerNumber || "-") + "</span>" +
      '  <h2 class="detail-name">' + escHtml(identity.name || "Unbekannt") + "</h2>" +
      "</div>";

    // Basic info
    html += '<div class="detail-section">';
    html += "<h4>Allgemeine Angaben</h4>";
    html += "<table class='detail-table'>";
    html += detailRow("Rechtsform", legalForm.code_de || legalForm.legalFormText || "-");
    html += detailRow(
      "Sitz",
      formatAddress(address)
    );
    html += detailRow("Telefon", identity.phoneNumber || "-");
    if (identity.websites && identity.websites.length > 0) {
      html += detailRow("Website", identity.websites.join(", "));
    }
    if (identity.organizationEmails && identity.organizationEmails.length > 0) {
      html += detailRow("E-Mail", identity.organizationEmails.join(", "));
    }
    html += detailRow("Mitgliederzahl", identity.members != null ? formatNumber(identity.members) : "-");
    if (identity.membersCountDate) {
      html += detailRow("Stand Mitgliederzahl", formatDate(identity.membersCountDate));
    }
    html += detailRow("Erstveroeffentlichung", account.firstPublicationDate ? formatDate(account.firstPublicationDate) : "-");
    html += "</table></div>";

    // Legal representatives
    const reps = identity.legalRepresentatives || [];
    if (reps.length > 0) {
      html += '<div class="detail-section">';
      html += "<h4>Geschaeftsfuehrung / Vertretung</h4>";
      html += '<ul class="detail-list">';
      reps.forEach((r) => {
        const parts = [r.academicDegreeBefore, r.commonFirstName, r.lastName, r.academicDegreeAfter].filter(Boolean);
        const fn = r.function ? " (" + r.function + ")" : "";
        html += "<li>" + escHtml(parts.join(" ") + fn) + "</li>";
      });
      html += "</ul></div>";
    }

    // Named employees
    const employees = identity.namedEmployees || [];
    if (employees.length > 0) {
      html += '<div class="detail-section">';
      html += "<h4>Beschaeftigte in der Interessenvertretung</h4>";
      html += "<p>" + (entry.employeeCount ? entry.employeeCount.from + " - " + entry.employeeCount.to : "-") + " Personen</p>";
      html += '<ul class="detail-list">';
      employees.forEach((e) => {
        const parts = [e.academicDegreeBefore, e.commonFirstName, e.lastName, e.academicDegreeAfter].filter(Boolean);
        html += "<li>" + escHtml(parts.join(" ")) + "</li>";
      });
      html += "</ul></div>";
    } else if (entry.employeeCount) {
      html += '<div class="detail-section">';
      html += "<h4>Beschaeftigte</h4>";
      html += "<p>" + entry.employeeCount.from + " - " + entry.employeeCount.to + " Personen</p>";
      html += "</div>";
    }

    // Financial info
    html += '<div class="detail-section">';
    html += "<h4>Finanzangaben</h4>";
    html += "<table class='detail-table'>";
    if (entry.financialExpensesEuro && !entry.refuseFinancialExpensesInformation) {
      const fin = entry.financialExpensesEuro;
      html += detailRow(
        "Finanzaufwand",
        formatCurrency(fin.from) + " - " + formatCurrency(fin.to)
      );
      if (fin.fiscalYearStart && fin.fiscalYearEnd) {
        html += detailRow("Geschaeftsjahr", fin.fiscalYearStart + " bis " + fin.fiscalYearEnd);
      }
    } else {
      html += detailRow("Finanzaufwand", entry.refuseFinancialExpensesInformation ? "Keine Angabe" : "-");
    }
    html += "</table></div>";

    // Donations
    const donators = entry.donators || [];
    const donations = donators.filter((d) => d.categoryType === "DONATIONS");
    const allowances = donators.filter((d) => d.categoryType === "PUBLIC_ALLOWANCES");

    if (donations.length > 0) {
      html += '<div class="detail-section">';
      html += "<h4>Spenden / Zuwendungen</h4>";
      html += '<ul class="detail-list">';
      donations.forEach((d) => {
        const amount = d.donationEuro
          ? formatCurrency(d.donationEuro.from) + " - " + formatCurrency(d.donationEuro.to)
          : "";
        html += "<li><strong>" + escHtml(d.name || "-") + "</strong>";
        if (d.location) html += ", " + escHtml(d.location);
        if (amount) html += " &mdash; " + amount;
        html += "</li>";
      });
      html += "</ul></div>";
    }

    if (allowances.length > 0) {
      html += '<div class="detail-section">';
      html += "<h4>Oeffentliche Zuwendungen</h4>";
      html += '<ul class="detail-list">';
      allowances.forEach((d) => {
        const amount = d.donationEuro
          ? formatCurrency(d.donationEuro.from) + " - " + formatCurrency(d.donationEuro.to)
          : "";
        html += "<li><strong>" + escHtml(d.name || "-") + "</strong>";
        if (amount) html += " &mdash; " + amount;
        html += "</li>";
      });
      html += "</ul></div>";
    }

    // Activities
    const activities = entry.activities || [];
    if (activities.length > 0) {
      html += '<div class="detail-section">';
      html += "<h4>Taetigkeitsbereiche</h4>";
      html += '<div class="tag-list">';
      activities.forEach((a) => {
        html += '<span class="tag">' + escHtml(a.de || a.text || a.code) + "</span>";
      });
      html += "</div></div>";
    }

    // Fields of interest
    const fois = entry.fieldsOfInterest || [];
    if (fois.length > 0) {
      html += '<div class="detail-section">';
      html += "<h4>Interessenbereiche</h4>";
      html += '<div class="tag-list">';
      fois.forEach((f) => {
        html += '<span class="tag">' + escHtml(f.de || f.fieldOfInterestText || f.code) + "</span>";
      });
      html += "</div></div>";
    }

    // Legislative projects
    const projects = entry.legislativeProjects || [];
    if (projects.length > 0) {
      html += '<div class="detail-section">';
      html += "<h4>Gesetzgeberische Vorhaben</h4>";
      html += '<ul class="detail-list">';
      projects.forEach((p) => {
        html += "<li>" + escHtml(p.name || "-");
        if (p.printingNumber) html += " (Drs. " + escHtml(p.printingNumber) + ")";
        html += "</li>";
      });
      html += "</ul></div>";
    }

    // Activity description
    if (account.activityDescription) {
      html += '<div class="detail-section">';
      html += "<h4>Beschreibung der Taetigkeit</h4>";
      html += '<p class="activity-desc">' + escHtml(account.activityDescription) + "</p>";
      html += "</div>";
    }

    // Client organizations
    const clients = entry.clientOrganizations || [];
    if (clients.length > 0) {
      html += '<div class="detail-section">';
      html += "<h4>Auftraggeber</h4>";
      html += '<ul class="detail-list">';
      clients.forEach((c) => {
        html += "<li><strong>" + escHtml(c.name || "-") + "</strong>";
        if (c.address && c.address.city) html += ", " + escHtml(c.address.city);
        html += "</li>";
      });
      html += "</ul></div>";
    }

    dom.detailContent.innerHTML = html;
  }

  // --------------- Office.js Document Insertion ---------------
  async function insertDetailIntoDocument() {
    if (!officeReady || !state.selectedEntry) return;

    const item = state.selectedEntry;
    const entry = item.registerEntryDetail || {};
    const identity = entry.lobbyistIdentity || {};
    const account = entry.account || {};
    const address = identity.address || {};
    const legalForm = identity.legalForm || {};

    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        // Title
        const titlePara = body.insertParagraph(
          "Lobbyregister-Eintrag: " + (identity.name || "Unbekannt"),
          Word.InsertLocation.end
        );
        titlePara.styleBuiltIn = Word.BuiltInStyle.heading1;

        // Register number
        const regPara = body.insertParagraph(
          "Register-Nr.: " + (item.registerNumber || account.registerNumber || "-"),
          Word.InsertLocation.end
        );
        regPara.styleBuiltIn = Word.BuiltInStyle.heading2;

        // Basic info table
        const basicData = [
          ["Feld", "Wert"],
          ["Name", identity.name || "-"],
          ["Rechtsform", legalForm.code_de || legalForm.legalFormText || "-"],
          ["Sitz", formatAddress(address)],
          ["Telefon", identity.phoneNumber || "-"],
          ["Website", (identity.websites || []).join(", ") || "-"],
          ["E-Mail", (identity.organizationEmails || []).join(", ") || "-"],
          ["Mitgliederzahl", identity.members != null ? formatNumber(identity.members) : "-"],
          ["Erstveroeffentlichung", account.firstPublicationDate ? formatDate(account.firstPublicationDate) : "-"],
        ];

        const sectionLabel1 = body.insertParagraph("Allgemeine Angaben", Word.InsertLocation.end);
        sectionLabel1.styleBuiltIn = Word.BuiltInStyle.heading3;

        const table1 = body.insertTable(basicData.length, 2, Word.InsertLocation.end, basicData);
        table1.styleBuiltIn = Word.BuiltInStyle.gridTable4_Accent1;
        table1.headerRowCount = 1;

        // Legal representatives
        const reps = identity.legalRepresentatives || [];
        if (reps.length > 0) {
          const repLabel = body.insertParagraph("Geschaeftsfuehrung / Vertretung", Word.InsertLocation.end);
          repLabel.styleBuiltIn = Word.BuiltInStyle.heading3;
          reps.forEach((r) => {
            const parts = [r.academicDegreeBefore, r.commonFirstName, r.lastName, r.academicDegreeAfter].filter(Boolean);
            const fn = r.function ? " (" + r.function + ")" : "";
            body.insertParagraph("- " + parts.join(" ") + fn, Word.InsertLocation.end);
          });
        }

        // Employees
        if (entry.employeeCount) {
          const empLabel = body.insertParagraph("Beschaeftigte", Word.InsertLocation.end);
          empLabel.styleBuiltIn = Word.BuiltInStyle.heading3;
          body.insertParagraph(
            entry.employeeCount.from + " - " + entry.employeeCount.to + " Personen",
            Word.InsertLocation.end
          );
        }

        // Financial info
        const finLabel = body.insertParagraph("Finanzangaben", Word.InsertLocation.end);
        finLabel.styleBuiltIn = Word.BuiltInStyle.heading3;

        if (entry.financialExpensesEuro && !entry.refuseFinancialExpensesInformation) {
          const fin = entry.financialExpensesEuro;
          body.insertParagraph(
            "Finanzaufwand: " + formatCurrency(fin.from) + " - " + formatCurrency(fin.to),
            Word.InsertLocation.end
          );
          if (fin.fiscalYearStart && fin.fiscalYearEnd) {
            body.insertParagraph(
              "Geschaeftsjahr: " + fin.fiscalYearStart + " bis " + fin.fiscalYearEnd,
              Word.InsertLocation.end
            );
          }
        } else {
          body.insertParagraph("Finanzaufwand: Keine Angabe", Word.InsertLocation.end);
        }

        // Donations
        const donators = entry.donators || [];
        const donations = donators.filter((d) => d.categoryType === "DONATIONS");
        if (donations.length > 0) {
          const donLabel = body.insertParagraph("Spenden / Zuwendungen", Word.InsertLocation.end);
          donLabel.styleBuiltIn = Word.BuiltInStyle.heading3;
          donations.forEach((d) => {
            const amount = d.donationEuro
              ? " (" + formatCurrency(d.donationEuro.from) + " - " + formatCurrency(d.donationEuro.to) + ")"
              : "";
            body.insertParagraph("- " + (d.name || "-") + (d.location ? ", " + d.location : "") + amount, Word.InsertLocation.end);
          });
        }

        // Fields of interest
        const fois = entry.fieldsOfInterest || [];
        if (fois.length > 0) {
          const foiLabel = body.insertParagraph("Interessenbereiche", Word.InsertLocation.end);
          foiLabel.styleBuiltIn = Word.BuiltInStyle.heading3;
          body.insertParagraph(
            fois.map((f) => f.de || f.fieldOfInterestText || f.code).join(", "),
            Word.InsertLocation.end
          );
        }

        // Activities
        const activities = entry.activities || [];
        if (activities.length > 0) {
          const actLabel = body.insertParagraph("Taetigkeitsbereiche", Word.InsertLocation.end);
          actLabel.styleBuiltIn = Word.BuiltInStyle.heading3;
          body.insertParagraph(
            activities.map((a) => a.de || a.text || a.code).join(", "),
            Word.InsertLocation.end
          );
        }

        // Legislative projects
        const projects = entry.legislativeProjects || [];
        if (projects.length > 0) {
          const projLabel = body.insertParagraph("Gesetzgeberische Vorhaben", Word.InsertLocation.end);
          projLabel.styleBuiltIn = Word.BuiltInStyle.heading3;
          projects.forEach((p) => {
            const dr = p.printingNumber ? " (Drs. " + p.printingNumber + ")" : "";
            body.insertParagraph("- " + (p.name || "-") + dr, Word.InsertLocation.end);
          });
        }

        // Activity description
        if (account.activityDescription) {
          const descLabel = body.insertParagraph("Beschreibung der Taetigkeit", Word.InsertLocation.end);
          descLabel.styleBuiltIn = Word.BuiltInStyle.heading3;
          body.insertParagraph(account.activityDescription, Word.InsertLocation.end);
        }

        // Client organizations
        const clients = entry.clientOrganizations || [];
        if (clients.length > 0) {
          const clientLabel = body.insertParagraph("Auftraggeber", Word.InsertLocation.end);
          clientLabel.styleBuiltIn = Word.BuiltInStyle.heading3;
          clients.forEach((c) => {
            const loc = c.address && c.address.city ? ", " + c.address.city : "";
            body.insertParagraph("- " + (c.name || "-") + loc, Word.InsertLocation.end);
          });
        }

        // Separator
        body.insertParagraph("", Word.InsertLocation.end);
        const sep = body.insertParagraph(
          "Quelle: Lobbyregister des Deutschen Bundestages (https://www.lobbyregister.bundestag.de)",
          Word.InsertLocation.end
        );
        sep.font.italic = true;
        sep.font.size = 9;
        sep.font.color = "#666666";

        await context.sync();
      });

      showToast("Eintrag wurde ins Dokument eingefuegt.");
    } catch (err) {
      showToast("Fehler beim Einfuegen: " + err.message);
    }
  }

  async function insertSingleEntry(item) {
    state.selectedEntry = item;
    await insertDetailIntoDocument();
  }

  async function insertAllResultsAsTable() {
    if (!officeReady || state.results.length === 0) return;

    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        const titlePara = body.insertParagraph(
          "Lobbyregister - Suchergebnisse" +
            (state.query ? ' fuer "' + state.query + '"' : ""),
          Word.InsertLocation.end
        );
        titlePara.styleBuiltIn = Word.BuiltInStyle.heading1;

        // Build table data
        const headers = [
          "Register-Nr.",
          "Name",
          "Rechtsform",
          "Sitz",
          "Mitglieder",
          "Finanzaufwand",
          "Taetigkeitsbereich",
          "Interessenbereiche",
        ];

        const rows = [headers];

        state.results.forEach((item) => {
          const entry = item.registerEntryDetail || {};
          const identity = entry.lobbyistIdentity || {};
          const account = entry.account || {};
          const legalForm = identity.legalForm || {};

          rows.push([
            item.registerNumber || account.registerNumber || "-",
            identity.name || "-",
            legalForm.code_de || legalForm.legalFormText || "-",
            identity.address ? identity.address.city || "-" : "-",
            identity.members != null ? formatNumber(identity.members) : "-",
            formatExpenses(entry.financialExpensesEuro),
            (entry.activities || []).map((a) => a.de || a.code).join(", ") || "-",
            (entry.fieldsOfInterest || []).map((f) => f.de || f.code).join(", ") || "-",
          ]);
        });

        const table = body.insertTable(rows.length, headers.length, Word.InsertLocation.end, rows);
        table.styleBuiltIn = Word.BuiltInStyle.gridTable4_Accent1;
        table.headerRowCount = 1;

        // Source note
        body.insertParagraph("", Word.InsertLocation.end);
        const src = body.insertParagraph(
          "Quelle: Lobbyregister des Deutschen Bundestages | Abfrage: " + new Date().toLocaleDateString("de-DE"),
          Word.InsertLocation.end
        );
        src.font.italic = true;
        src.font.size = 9;
        src.font.color = "#666666";

        await context.sync();
      });

      showToast(state.results.length + " Eintraege als Tabelle eingefuegt.");
    } catch (err) {
      showToast("Fehler beim Einfuegen: " + err.message);
    }
  }

  // --------------- UI Helpers ---------------
  function showLoading(show) {
    state.isLoading = show;
    dom.loading.classList.toggle("hidden", !show);
    if (show) {
      dom.emptyState.classList.add("hidden");
      dom.resultsList.innerHTML = "";
      dom.resultsInfo.classList.add("hidden");
      dom.pagination.classList.add("hidden");
    }
  }

  function showError(msg) {
    dom.errorState.classList.remove("hidden");
    dom.errorMessage.textContent = msg;
    dom.emptyState.classList.add("hidden");
  }

  function hideError() {
    dom.errorState.classList.add("hidden");
  }

  function showToast(msg) {
    dom.toastMessage.textContent = msg;
    dom.toast.classList.remove("hidden");
    dom.toast.classList.add("show");
    setTimeout(() => {
      dom.toast.classList.remove("show");
      dom.toast.classList.add("hidden");
    }, 3000);
  }

  // --------------- Formatting Helpers ---------------
  function formatNumber(n) {
    if (n == null) return "-";
    return n.toLocaleString("de-DE");
  }

  function formatCurrency(n) {
    if (n == null) return "-";
    return n.toLocaleString("de-DE", { style: "currency", currency: "EUR", minimumFractionDigits: 0 });
  }

  function formatExpenses(fin) {
    if (!fin) return "-";
    return formatCurrency(fin.from) + " - " + formatCurrency(fin.to);
  }

  function formatAddress(addr) {
    if (!addr) return "-";
    const parts = [];
    if (addr.street) {
      parts.push(addr.street + (addr.streetNumber ? " " + addr.streetNumber : ""));
    }
    if (addr.zipCode || addr.city) {
      parts.push((addr.zipCode ? addr.zipCode + " " : "") + (addr.city || ""));
    }
    if (addr.country && addr.country.code && addr.country.code !== "DE") {
      parts.push(addr.country.code);
    }
    return parts.join(", ") || "-";
  }

  function formatDate(dateStr) {
    if (!dateStr) return "-";
    try {
      const d = new Date(dateStr);
      return d.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" });
    } catch (_) {
      return dateStr;
    }
  }

  function escHtml(str) {
    if (!str) return "";
    var div = document.createElement("div");
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
  }

  function detailRow(label, value) {
    return "<tr><td class='detail-label'>" + escHtml(label) + "</td><td>" + escHtml(value) + "</td></tr>";
  }
})();
