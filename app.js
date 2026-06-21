(() => {
  "use strict";

  const { foods, meals, sources } = window.APP_DATA;
  const storageKey = "lebensmitteleinkauf:selected:v1";

  const iconPaths = {
    basket: '<path d="M7 10 10 4M17 10l-3-6M4 10h16l-1.4 9H5.4L4 10Z"/><path d="M9 13v3M15 13v3"/>',
    download: '<path d="M12 3v12M7.5 10.5 12 15l4.5-4.5"/><path d="M4 18v2h16v-2"/>',
    search: '<circle cx="11" cy="11" r="6.5"/><path d="m16 16 4 4"/>',
    leaf: '<path d="M19 4C11 4 6 8 6 15c0 3 2 5 5 5 7 0 9-8 8-16Z"/><path d="M5 21c2-5 6-9 11-12"/>',
    meal: '<path d="M7 3v8M4.5 3v5c0 2 1 3 2.5 3s2.5-1 2.5-3V3M7 11v10"/><path d="M16 3c2 2 3 5 3 8v2h-5V9c0-3 1-5 2-6Zm0 10v8"/>',
    chart: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
  };

  const categoryIcons = {
    fish: '<path d="M4 12c4-5 9-6 14-2l3-3v10l-3-3c-5 4-10 3-14-2Z"/><circle cx="15" cy="11" r=".8"/>',
    animal: '<path d="M8 7c1-3 3-4 5-4 4 0 6 4 6 8s-3 8-7 8-7-3-7-7c0-2 1-4 3-5Z"/><path d="M9 7 6 4M16 6l2-3"/>',
    beans: '<path d="M10 4c4 1 5 5 3 9s-6 6-9 3-1-6 1-9c1-2 3-3 5-3ZM17 8c3 1 4 4 3 7s-4 5-7 4"/>',
    dairy: '<path d="M8 4h8v4l2 3v9H6v-9l2-3V4Z"/><path d="M8 8h8M8 14h10"/>',
    vegetable: '<path d="M12 8c-5 0-8 3-8 7s3 6 8 6 8-2 8-6-3-7-8-7Z"/><path d="M12 8c-2-3-1-5 0-6M12 7c2-3 4-3 6-2M11 7C9 4 7 4 5 5"/>',
    fruit: '<circle cx="12" cy="13" r="7"/><path d="M12 6c0-3 2-4 4-4M13 5c3-1 5 0 6 2"/>',
    nuts: '<path d="M12 3c4 3 7 6 7 11a7 7 0 0 1-14 0c0-5 3-8 7-11Z"/><path d="M12 5v16M8 9c1 1 2 2 4 2M16 13c-1 1-2 2-4 2"/>',
    grain: '<path d="M12 22V5M12 8c-3 0-5-2-5-4 3 0 5 2 5 4ZM12 12c3 0 5-2 5-4-3 0-5 2-5 4ZM12 16c-3 0-5-2-5-4 3 0 5 2 5 4ZM12 20c3 0 5-2 5-4-3 0-5 2-5 4Z"/>',
    herbs: '<path d="M5 20C7 12 11 7 19 4c0 8-4 13-12 14"/><path d="M7 18c4-4 7-7 11-11"/>',
    drink: '<path d="M6 7h11v8a5 5 0 0 1-5 5h-1a5 5 0 0 1-5-5V7Z"/><path d="M17 9h2a3 3 0 0 1 0 6h-2M8 3c0 1 1 1 1 2M13 3c0 1 1 1 1 2"/>',
  };

  const state = {
    view: "foods",
    search: "",
    mealSearch: "",
    category: "",
    score: "",
    priority: "",
    limit: window.innerWidth < 680 ? 18 : 28,
    selected: loadSelection(),
    confirmAction: null,
  };

  const dom = {
    foodGrid: document.querySelector("#foodGrid"),
    resultCount: document.querySelector("#resultCount"),
    searchInput: document.querySelector("#searchInput"),
    mealSearchInput: document.querySelector("#mealSearchInput"),
    categoryFilter: document.querySelector("#categoryFilter"),
    scoreFilter: document.querySelector("#scoreFilter"),
    priorityFilter: document.querySelector("#priorityFilter"),
    resetFilters: document.querySelector("#resetFilters"),
    activeFilters: document.querySelector("#activeFilters"),
    loadMore: document.querySelector("#loadMore"),
    shoppingPanel: document.querySelector("#shoppingPanel"),
    shoppingItems: document.querySelector("#shoppingItems"),
    shoppingActions: document.querySelector("#shoppingActions"),
    shoppingCount: document.querySelector("#shoppingCount"),
    categoryCount: document.querySelector("#categoryCount"),
    headerCount: document.querySelector("#headerCount"),
    mobileCount: document.querySelector("#mobileCount"),
    basketButton: document.querySelector("#basketButton"),
    mobileBasket: document.querySelector("#mobileBasket"),
    closeShopping: document.querySelector("#closeShopping"),
    scrim: document.querySelector("#scrim"),
    detailDialog: document.querySelector("#detailDialog"),
    detailContent: document.querySelector("#detailContent"),
    confirmDialog: document.querySelector("#confirmDialog"),
    confirmTitle: document.querySelector("#confirmTitle"),
    confirmText: document.querySelector("#confirmText"),
    cancelConfirm: document.querySelector("#cancelConfirm"),
    acceptConfirm: document.querySelector("#acceptConfirm"),
    toast: document.querySelector("#toast"),
  };

  function icon(name, viewBox = "0 0 24 24") {
    return `<svg viewBox="${viewBox}" aria-hidden="true">${iconPaths[name] || iconPaths.leaf}</svg>`;
  }

  function categoryIcon(category) {
    const value = category.toLowerCase();
    let key = "herbs";
    if (value.includes("fisch")) key = "fish";
    else if (value.includes("fleisch")) key = "animal";
    else if (value.includes("hülsen")) key = "beans";
    else if (value.includes("milch")) key = "dairy";
    else if (value.includes("gemüse")) key = "vegetable";
    else if (value === "obst") key = "fruit";
    else if (value.includes("nüsse")) key = "nuts";
    else if (value.includes("vollkorn")) key = "grain";
    else if (value.includes("getränke")) key = "drink";
    return `<svg viewBox="0 0 24 24" aria-hidden="true">${categoryIcons[key]}</svg>`;
  }

  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>'"]/g, (character) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      "'": "&#039;",
      '"': "&quot;",
    })[character]);
  }

  function loadSelection() {
    try {
      return new Set(JSON.parse(localStorage.getItem(storageKey) || "[]").map(Number));
    } catch {
      return new Set();
    }
  }

  function persistSelection() {
    localStorage.setItem(storageKey, JSON.stringify([...state.selected]));
  }

  function populateFilters() {
    const categories = [...new Set(foods.map((food) => food.category))];
    dom.categoryFilter.insertAdjacentHTML(
      "beforeend",
      categories.map((category) => `<option value="${escapeHtml(category)}">${escapeHtml(category)}</option>`).join(""),
    );
  }

  function filteredFoods() {
    const term = state.search.trim().toLocaleLowerCase("de");
    return foods.filter((food) => {
      const searchable = `${food.name} ${food.category} ${food.subcategory} ${food.compounds} ${food.benefit}`.toLocaleLowerCase("de");
      return (!term || searchable.includes(term))
        && (!state.category || food.category === state.category)
        && (!state.score || food.score === Number(state.score))
        && (!state.priority || food.priority.toLocaleLowerCase("de") === state.priority);
    });
  }

  function renderFoodCard(food) {
    const selected = state.selected.has(food.id);
    return `
      <article class="food-card${selected ? " is-selected" : ""}" data-id="${food.id}">
        <button class="select-food" type="button" aria-label="${escapeHtml(food.name)} ${selected ? "von der Einkaufsliste entfernen" : "zur Einkaufsliste hinzufügen"}" aria-pressed="${selected}"></button>
        <div class="food-card-body">
          <div class="food-card-top">
            <div>
              <span class="category-label"><span class="category-symbol">${categoryIcon(food.category)}</span>${escapeHtml(food.subcategory)}</span>
              <h3>${escapeHtml(food.name)}</h3>
            </div>
            <span class="priority-dot"><i></i>${escapeHtml(food.priority)}</span>
          </div>
          <div class="food-meta">
            <span class="meta-pill score">Score ${food.score} · ${escapeHtml(food.satiety)}</span>
            <span class="meta-pill">KH ${escapeHtml(food.carbs)}</span>
          </div>
          <p class="food-benefit">${escapeHtml(food.benefit)}</p>
          <div class="food-card-footer">
            <span class="compound">${escapeHtml(food.compounds)}</span>
            <button class="details-button" type="button">Details →</button>
          </div>
        </div>
      </article>`;
  }

  function renderFoods() {
    const result = filteredFoods();
    const visible = result.slice(0, state.limit);
    dom.foodGrid.innerHTML = visible.length
      ? visible.map(renderFoodCard).join("")
      : '<div class="empty-results"><strong>Keine Treffer.</strong><br />Versuche eine andere Suche oder setze die Filter zurück.</div>';
    dom.resultCount.textContent = `${result.length} ${result.length === 1 ? "Eintrag" : "Einträge"}`;
    dom.loadMore.hidden = visible.length >= result.length;
    renderActiveFilters();
  }

  function renderActiveFilters() {
    const chips = [];
    if (state.search) chips.push(`Suche: ${state.search}`);
    if (state.category) chips.push(state.category);
    if (state.score) chips.push(`Score ${state.score}`);
    if (state.priority) chips.push(`Priorität: ${state.priority}`);
    dom.activeFilters.innerHTML = chips.map((chip) => `<span class="filter-chip">${escapeHtml(chip)}</span>`).join("");
  }

  function selectedFoods() {
    return foods.filter((food) => state.selected.has(food.id));
  }

  function renderShoppingList() {
    const selected = selectedFoods();
    const groups = new Map();
    selected.forEach((food) => {
      if (!groups.has(food.category)) groups.set(food.category, []);
      groups.get(food.category).push(food);
    });
    dom.headerCount.textContent = selected.length;
    dom.mobileCount.textContent = selected.length;
    dom.shoppingCount.textContent = selected.length;
    dom.categoryCount.textContent = `${groups.size} ${groups.size === 1 ? "Kategorie" : "Kategorien"}`;
    dom.shoppingActions.classList.toggle("is-disabled", !selected.length);
    if (!selected.length) {
      dom.shoppingItems.innerHTML = `
        <div class="empty-list">
          <span class="empty-list-icon"><span class="button-icon">${icon("basket")}</span></span>
          <strong>Noch nichts ausgewählt</strong>
          <p>Tippe bei einem Lebensmittel auf das Kästchen. Deine Auswahl bleibt auf diesem Gerät gespeichert.</p>
        </div>`;
      return;
    }
    dom.shoppingItems.innerHTML = [...groups.entries()].map(([category, items]) => `
      <section class="shopping-group">
        <h3>${escapeHtml(category)}</h3>
        ${items.map((food) => `
          <div class="shopping-item">
            <span>${escapeHtml(food.name)}</span>
            <button class="remove-item" type="button" data-remove-id="${food.id}" aria-label="${escapeHtml(food.name)} entfernen">×</button>
          </div>`).join("")}
      </section>`).join("");
  }

  function toggleFood(id) {
    if (state.selected.has(id)) state.selected.delete(id);
    else state.selected.add(id);
    persistSelection();
    renderFoods();
    renderShoppingList();
  }

  function openDetails(id) {
    const food = foods.find((item) => item.id === id);
    if (!food) return;
    const links = String(food.sources).split(";").map((url) => url.trim()).filter(Boolean);
    dom.detailContent.innerHTML = `
      <div class="detail-content">
        <div class="detail-icon">${categoryIcon(food.category)}</div>
        <p class="eyebrow">${escapeHtml(food.category)}</p>
        <h2>${escapeHtml(food.name)}</h2>
        <p class="detail-subtitle">${escapeHtml(food.subcategory)}</p>
        <div class="detail-badges">
          <span class="meta-pill score">Sättigung ${food.score}/5 · ${escapeHtml(food.satiety)}</span>
          <span class="meta-pill">Kohlenhydrate ${escapeHtml(food.carbs)}</span>
          <span class="meta-pill">Priorität ${escapeHtml(food.priority)}</span>
        </div>
        ${detailSection("Warum sinnvoll", food.benefit)}
        ${detailSection("Aktive Stoffe", food.compounds)}
        ${detailSection("Praktische Verwendung", food.use)}
        ${detailSection("Hinweise / Einschränkungen", food.caution)}
        ${links.length ? `<section class="detail-section"><h3>Quellen</h3><div class="detail-sources">${links.map((url, index) => `<a href="${escapeHtml(url)}" target="_blank" rel="noreferrer">Quelle ${index + 1} ↗</a>`).join("")}</div></section>` : ""}
      </div>`;
    dom.detailDialog.showModal();
  }

  function detailSection(title, value) {
    if (!value) return "";
    return `<section class="detail-section"><h3>${escapeHtml(title)}</h3><p>${escapeHtml(value)}</p></section>`;
  }

  function formatMealDate(value) {
    const [year, month, day] = String(value || "").split("-");
    return year && month && day ? `${day}.${month}.${year}` : String(value || "");
  }

  function renderMeals() {
    const term = state.mealSearch.trim().toLocaleLowerCase("de");
    const visibleMeals = meals
      .map((meal, index) => ({ meal, index }))
      .filter(({ meal }) => {
        const searchable = `${meal.situation} ${meal.satiety} ${meal.combination} ${meal.reason} ${meal.variants} ${formatMealDate(meal.date)}`.toLocaleLowerCase("de");
        return !term || searchable.includes(term);
      });

    document.querySelector("#mealGrid").innerHTML = visibleMeals.length ? visibleMeals.map(({ meal, index }) => {
      const mealDate = formatMealDate(meal.date);
      return `
      <article class="meal-card">
        <div class="meal-meta">
          ${mealDate ? `<time class="meal-date" datetime="${escapeHtml(meal.date)}">${escapeHtml(mealDate)}</time>` : ""}
          <span class="meal-number">${String(index + 1).padStart(2, "0")}</span>
        </div>
        <p class="eyebrow">${escapeHtml(meal.satiety)} sättigend</p>
        <h2>${escapeHtml(meal.situation)}</h2>
        <p class="meal-combination">${escapeHtml(meal.combination)}</p>
        <div class="meal-details">
          <div><span>Warum sinnvoll</span><p>${escapeHtml(meal.reason)}</p></div>
          <div><span>Varianten</span><p>${escapeHtml(meal.variants)}</p></div>
        </div>
      </article>`;
    }).join("") : '<div class="empty-results">Keine passenden Empfehlungen gefunden.</div>';
  }

  function renderInsights() {
    const categoryCounts = [...foods.reduce((map, food) => map.set(food.category, (map.get(food.category) || 0) + 1), new Map())];
    const scoreCounts = [5, 4, 3, 2, 1].map((score) => ({ score, count: foods.filter((food) => food.score === score).length }));
    const average = foods.reduce((sum, food) => sum + food.score, 0) / foods.length;
    const topRated = scoreCounts[0].count + scoreCounts[1].count;
    const metrics = [
      [foods.length, "Lebensmittel gesamt", "vollständige Excel-Liste"],
      [topRated, "Score 4 oder 5", `${Math.round(topRated / foods.length * 100)} % der Auswahl`],
      [categoryCounts.length, "Oberkategorien", "klar gegliedert"],
      [average.toLocaleString("de-DE", { maximumFractionDigits: 1 }), "Ø Sättigungs-Score", "von maximal 5"],
    ];
    document.querySelector("#metricGrid").innerHTML = metrics.map(([value, label, note]) => `
      <article class="metric-card"><span>${label}</span><strong>${value}</strong><small>${note}</small></article>`).join("");
    const maxCount = Math.max(...categoryCounts.map(([, count]) => count));
    document.querySelector("#categoryChart").innerHTML = categoryCounts.map(([category, count]) => `
      <div class="bar-row" title="${escapeHtml(category)}: ${count}">
        <span class="bar-label">${escapeHtml(category)}</span>
        <span class="bar-track"><i class="bar-fill" style="width:${count / maxCount * 100}%"></i></span>
        <strong class="bar-value">${count}</strong>
      </div>`).join("");
    const colors = ["#145c38", "#5e9d78", "#205f86", "#86bad1", "#d2dcd6"];
    let start = 0;
    const segments = scoreCounts.map(({ count }, index) => {
      const end = start + count / foods.length * 100;
      const segment = `${colors[index]} ${start}% ${end}%`;
      start = end;
      return segment;
    });
    document.querySelector("#donutChart").style.background = `conic-gradient(${segments.join(",")})`;
    document.querySelector("#donutLegend").innerHTML = scoreCounts.map(({ score, count }, index) => `
      <div class="legend-row"><i style="background:${colors[index]}"></i><span>Score ${score}</span><strong>${count}</strong></div>`).join("");
    document.querySelector("#sourceList").innerHTML = sources.map((source) => `
      <a class="source-link" href="${escapeHtml(source.url)}" target="_blank" rel="noreferrer"><span>${escapeHtml(source.name.replace(/ – .*/, ""))}</span><span>↗</span></a>`).join("");
  }

  function setView(view, updateHash = true) {
    if (!document.querySelector(`[data-view-panel="${view}"]`)) view = "foods";
    state.view = view;
    document.querySelectorAll("[data-view-panel]").forEach((panel) => panel.classList.toggle("is-active", panel.dataset.viewPanel === view));
    document.querySelectorAll("[data-view]").forEach((button) => button.classList.toggle("is-active", button.dataset.view === view));
    if (updateHash) history.replaceState(null, "", `#${view === "foods" ? "lebensmittel" : view === "meals" ? "tagesbaukasten" : "auswertung"}`);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function openShopping() {
    if (window.innerWidth > 900) {
      dom.shoppingPanel.scrollIntoView({ behavior: "smooth", block: "start" });
      return;
    }
    dom.shoppingPanel.classList.add("is-open");
    dom.scrim.classList.add("is-visible");
    document.body.style.overflow = "hidden";
  }

  function closeShopping() {
    dom.shoppingPanel.classList.remove("is-open");
    dom.scrim.classList.remove("is-visible");
    document.body.style.overflow = "";
  }

  function listText() {
    return selectedFoods().map((food) => food.name).join("\r\n");
  }

  function downloadList() {
    if (!state.selected.size) return;
    const blob = new Blob([listText()], { type: "text/plain;charset=utf-8" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Einkaufsliste_antientzuendliche_Lebensmittel.txt";
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    showToast("Einkaufsliste wurde als TXT gespeichert.");
    openConfirm({
      title: "Markierungen jetzt löschen?",
      text: "Die TXT-Datei ist gespeichert. Möchtest du die aktuelle Auswahl leeren?",
      cancel: "Behalten",
      accept: "Liste leeren",
      action: clearSelection,
    });
  }

  async function copyList() {
    if (!state.selected.size) return;
    try {
      await navigator.clipboard.writeText(listText());
    } catch {
      const textArea = document.createElement("textarea");
      textArea.value = listText();
      document.body.append(textArea);
      textArea.select();
      document.execCommand("copy");
      textArea.remove();
    }
    showToast("Einkaufsliste kopiert.");
  }

  function openConfirm({ title, text, cancel = "Abbrechen", accept = "Liste leeren", action }) {
    state.confirmAction = action;
    dom.confirmTitle.textContent = title;
    dom.confirmText.textContent = text;
    dom.cancelConfirm.textContent = cancel;
    dom.acceptConfirm.textContent = accept;
    dom.confirmDialog.showModal();
  }

  function clearSelection() {
    state.selected.clear();
    persistSelection();
    renderFoods();
    renderShoppingList();
    closeShopping();
    showToast("Alle Markierungen wurden gelöscht.");
  }

  let toastTimer;
  function showToast(message) {
    clearTimeout(toastTimer);
    dom.toast.textContent = message;
    dom.toast.classList.add("is-visible");
    toastTimer = setTimeout(() => dom.toast.classList.remove("is-visible"), 2600);
  }

  function bindEvents() {
    document.addEventListener("click", (event) => {
      const viewButton = event.target.closest("[data-view]");
      if (viewButton) setView(viewButton.dataset.view);
    });
    dom.foodGrid.addEventListener("click", (event) => {
      const card = event.target.closest(".food-card");
      if (!card) return;
      const id = Number(card.dataset.id);
      if (event.target.closest(".select-food")) toggleFood(id);
      if (event.target.closest(".details-button")) openDetails(id);
    });
    dom.shoppingItems.addEventListener("click", (event) => {
      const remove = event.target.closest("[data-remove-id]");
      if (remove) toggleFood(Number(remove.dataset.removeId));
    });
    dom.searchInput.addEventListener("input", () => {
      state.search = dom.searchInput.value;
      state.limit = window.innerWidth < 680 ? 18 : 28;
      renderFoods();
    });
    dom.mealSearchInput.addEventListener("input", () => {
      state.mealSearch = dom.mealSearchInput.value;
      renderMeals();
    });
    [[dom.categoryFilter, "category"], [dom.scoreFilter, "score"], [dom.priorityFilter, "priority"]].forEach(([element, key]) => {
      element.addEventListener("change", () => {
        state[key] = element.value;
        state.limit = window.innerWidth < 680 ? 18 : 28;
        renderFoods();
      });
    });
    dom.resetFilters.addEventListener("click", () => {
      state.search = state.category = state.score = state.priority = "";
      dom.searchInput.value = dom.categoryFilter.value = dom.scoreFilter.value = dom.priorityFilter.value = "";
      renderFoods();
    });
    dom.loadMore.addEventListener("click", () => {
      state.limit += window.innerWidth < 680 ? 18 : 28;
      renderFoods();
    });
    document.querySelector("#showTopRated").addEventListener("click", () => {
      state.score = "5";
      dom.scoreFilter.value = "5";
      renderFoods();
      document.querySelector("#catalogTitle").scrollIntoView({ behavior: "smooth", block: "start" });
    });
    [dom.basketButton, dom.mobileBasket].forEach((button) => button.addEventListener("click", openShopping));
    [dom.closeShopping, dom.scrim].forEach((element) => element.addEventListener("click", closeShopping));
    document.querySelector("#downloadList").addEventListener("click", downloadList);
    document.querySelector("#copyList").addEventListener("click", copyList);
    document.querySelector("#clearList").addEventListener("click", () => openConfirm({
      title: "Einkaufsliste leeren?",
      text: `Alle ${state.selected.size} Markierungen werden entfernt.`,
      action: clearSelection,
    }));
    document.querySelector(".dialog-close").addEventListener("click", () => dom.detailDialog.close());
    dom.detailDialog.addEventListener("click", (event) => {
      if (event.target === dom.detailDialog) dom.detailDialog.close();
    });
    dom.cancelConfirm.addEventListener("click", () => dom.confirmDialog.close());
    dom.acceptConfirm.addEventListener("click", () => {
      dom.confirmDialog.close();
      if (state.confirmAction) state.confirmAction();
      state.confirmAction = null;
    });
    document.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        if (state.view === "meals") dom.mealSearchInput.focus();
        else {
          setView("foods");
          dom.searchInput.focus();
        }
      }
      if (event.key === "Escape") closeShopping();
    });
  }

  function initialize() {
    document.querySelectorAll("[data-icon]").forEach((slot) => { slot.innerHTML = icon(slot.dataset.icon); });
    populateFilters();
    renderFoods();
    renderShoppingList();
    renderMeals();
    renderInsights();
    bindEvents();
    const viewByHash = { "#lebensmittel": "foods", "#tagesbaukasten": "meals", "#auswertung": "insights" };
    setView(viewByHash[window.location.hash] || "foods", false);
  }

  initialize();
})();
