(() => {
  "use strict";

  const { foods, meals, sources, foodNames = [] } = window.APP_DATA;
  const appVersion = "meal-guide-cleanup-20260920-1";
  const appVersionFile = "app-version.json";
  const appRefreshParam = "appRefresh";
  const appRefreshSessionKey = "lebensmitteleinkauf:app-refresh-version:v1";
  const currentViewStorageKey = "lebensmitteleinkauf:current-view:v1";
  const currentViewScrollStorageKey = "lebensmitteleinkauf:current-view-scroll:v1";
  const storageKey = "lebensmitteleinkauf:selected:v1";
  const storageMetaKey = "lebensmitteleinkauf:selected:meta:v1";
  const bookmarkStorageKey = "lebensmitteleinkauf:bookmarks:v1";
  const sharedMealStorageKey = "lebensmitteleinkauf:shared-meals:v1";
  const sharedFoodStorageKey = "lebensmitteleinkauf:shared-foods:v1";
  const recipeStorageKey = "lebensmitteleinkauf:recipes:v1";
  const sharedMealUrlParam = "mahlzeit";
  const sharedFoodUrlParam = "lebensmittel";
  const pendingLoginKey = "lebensmitteleinkauf:onedrive-login-pending:v1";
  const manualLogoutKey = "lebensmitteleinkauf:onedrive-manual-logout:v1";
  const authReloadKey = "lebensmitteleinkauf:onedrive-auth-reload:v1";
  const authConfirmAfterReloadKey = "lebensmitteleinkauf:onedrive-confirm-after-reload:v1";
  const authReloadParam = "onedriveAuthRefresh";
  const authLogoutParam = "onedriveLogout";
  const appDataFileName = "lebensmitteleinkauf-data.json";
  const graphBaseUrl = "https://graph.microsoft.com/v1.0";
  const graphFilePath = `/me/drive/special/approot:/${appDataFileName}`;
  const graphContentPath = `${graphFilePath}:/content`;
  const msalConfig = {
    auth: {
      clientId: "3c4004e7-9323-440c-8977-96699d8d8e6f",
      authority: "https://login.microsoftonline.com/common",
      redirectUri: window.location.origin + window.location.pathname,
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: "localStorage",
      temporaryCacheLocation: "localStorage",
      storeAuthStateInCookie: true,
    },
  };
  const graphScopes = ["User.Read", "Files.ReadWrite.AppFolder"];
  const loginRequest = { scopes: graphScopes };
  const validFoodIds = new Set(foods.map((food) => food.id));
  const validMealIds = new Set(meals.map((meal) => meal.id));
  const localSnapshot = loadSelectionData();
  const foodById = new Map(foods.map((food) => [food.id, food]));
  const foodByName = new Map(foods.map((food) => [normalizeFoodName(food.name), food]));
  const mealIndexById = new Map(meals.map((meal, index) => [meal.id, index]));
  const mealGuideImages = {
    1: { src: "assets/meal-guide/step-1.png?v=meal-guide-cleanup-20260920-1", alt: "Bildanleitung zu Schritt 1: Eine Mahlzeit auswählen" },
    2: { src: "assets/meal-guide/step-2.png?v=meal-guide-cleanup-20260920-1", alt: "Bildanleitung zu Schritt 2: Text für die Rezeptsuche kopieren" },
    3: { src: "assets/meal-guide/step-3.png?v=meal-guide-cleanup-20260920-1", alt: "Bildanleitung zu Schritt 3: Den kopierten Text in eine KI einfügen" },
    4: { src: "assets/meal-guide/step-4.png?v=meal-guide-cleanup-20260920-1", alt: "Bildanleitung zu Schritt 4: Zutaten auf die Einkaufsliste setzen" },
  };

  const iconPaths = {
    basket: '<path d="M7 10 10 4M17 10l-3-6M4 10h16l-1.4 9H5.4L4 10Z"/><path d="M9 13v3M15 13v3"/>',
    download: '<path d="M12 3v12M7.5 10.5 12 15l4.5-4.5"/><path d="M4 18v2h16v-2"/>',
    search: '<circle cx="11" cy="11" r="6.5"/><path d="m16 16 4 4"/>',
    leaf: '<path d="M19 4C11 4 6 8 6 15c0 3 2 5 5 5 7 0 9-8 8-16Z"/><path d="M5 21c2-5 6-9 11-12"/>',
    meal: '<path d="M7 3v8M4.5 3v5c0 2 1 3 2.5 3s2.5-1 2.5-3V3M7 11v10"/><path d="M16 3c2 2 3 5 3 8v2h-5V9c0-3 1-5 2-6Zm0 10v8"/>',
    chart: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
    bookmark: '<path d="M6.5 3.5h11v17L12 17l-5.5 3.5v-17Z"/>',
    recipe: '<path d="M5 4.5A2.5 2.5 0 0 1 7.5 2H19v17H7.5A2.5 2.5 0 0 0 5 21.5v-17Z"/><path d="M5 19a2 2 0 0 1 2-2h12M9 7h6M9 11h7"/>',
    share: '<circle cx="18" cy="5" r="2.5"/><circle cx="6" cy="12" r="2.5"/><circle cx="18" cy="19" r="2.5"/><path d="m8.2 10.8 7.6-4.5M8.2 13.2l7.6 4.5"/>',
    help: '<circle cx="12" cy="12" r="9"/><path d="M9.8 9a2.4 2.4 0 0 1 4.6 1c0 2-2.4 2.1-2.4 4M12 17.5v.1"/>',
    arrowUp: '<path d="m6 10 6-6 6 6M12 4v16"/>',
    imageOpen: '<rect x="3" y="5" width="13" height="14" rx="2"/><path d="m5.5 16 3.2-3.2 2.5 2.5 1.8-1.8 3 3"/><path d="M14 3h7v7M21 3l-8 8"/>',
    cloud: '<path d="M17.5 18H8a5 5 0 1 1 1.2-9.85A6.5 6.5 0 0 1 21 12a3 3 0 0 1-3.5 6Z"/><path d="M12 13v7M9 16l3-3 3 3"/>',
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
    mealType: "",
    mealCategory: "",
    recipeSearch: "",
    recipeCategory: "",
    recipeItemKey: "",
    category: "",
    score: "",
    priority: "",
    limit: window.innerWidth < 680 ? 18 : 28,
    selected: new Set(localSnapshot.selected),
    bookmarkedFoods: new Set(localSnapshot.bookmarkedFoodIds),
    bookmarkedMeals: new Set(localSnapshot.bookmarkedMealIds),
    sharedMeals: new Set(localSnapshot.sharedMealIds),
    sharedFoods: new Set(localSnapshot.sharedFoodIds),
    recipes: localSnapshot.recipes,
    selectedSharedMeals: new Set(),
    selectedSharedFoods: new Set(),
    pendingSharedMeals: new Set(),
    pendingSharedFoods: new Set(),
    sharedTargetMealId: null,
    sharedTargetFoodId: null,
    localUpdatedAt: localSnapshot.updatedAt,
    sync: {
      msal: null,
      account: null,
      initialized: false,
      busy: false,
      allowInteractiveTokenRedirect: false,
      needsInteractiveToken: false,
      redirectAccessToken: "",
      redirectAccessTokenExpiresAt: 0,
      status: "local",
      title: "Nicht angemeldet",
      message: "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert.",
      lastRemoteUpdatedAt: "",
      lastRemoteEtag: "",
      hasRemoteData: false,
      conflictData: null,
    },
    confirmAction: null,
    confirmCancelAction: null,
  };

  const dom = {
    foodGrid: document.querySelector("#foodGrid"),
    mealGrid: document.querySelector("#mealGrid"),
    bookmarkedFoodGrid: document.querySelector("#bookmarkedFoodGrid"),
    bookmarkedMealGrid: document.querySelector("#bookmarkedMealGrid"),
    sharedMealGrid: document.querySelector("#sharedMealGrid"),
    sharedFoodGrid: document.querySelector("#sharedFoodGrid"),
    sharedMealCount: document.querySelector("#sharedMealCount"),
    sharedFoodCount: document.querySelector("#sharedFoodCount"),
    sharedMealSelectionCount: document.querySelector("#sharedMealSelectionCount"),
    sharedFoodSelectionCount: document.querySelector("#sharedFoodSelectionCount"),
    removeSelectedSharedMeals: document.querySelector("#removeSelectedSharedMeals"),
    removeSelectedSharedFoods: document.querySelector("#removeSelectedSharedFoods"),
    bookmarkedFoodCount: document.querySelector("#bookmarkedFoodCount"),
    bookmarkedMealCount: document.querySelector("#bookmarkedMealCount"),
    clearAllBookmarks: document.querySelector("#clearAllBookmarks"),
    searchInput: document.querySelector("#searchInput"),
    allFoodsButton: document.querySelector("#allFoodsButton"),
    mealSearchInput: document.querySelector("#mealSearchInput"),
    mealTypeFilter: document.querySelector("#mealTypeFilter"),
    mealCategoryFilter: document.querySelector("#mealCategoryFilter"),
    resetMealFilters: document.querySelector("#resetMealFilters"),
    categoryFilter: document.querySelector("#categoryFilter"),
    scoreFilter: document.querySelector("#scoreFilter"),
    priorityFilter: document.querySelector("#priorityFilter"),
    resetFilters: document.querySelector("#resetFilters"),
    activeFilters: document.querySelector("#activeFilters"),
    mealResultCount: document.querySelector("#mealResultCount"),
    recipeSearchInput: document.querySelector("#recipeSearchInput"),
    recipeCategoryFilter: document.querySelector("#recipeCategoryFilter"),
    recipeItemFilter: document.querySelector("#recipeItemFilter"),
    resetRecipeFilters: document.querySelector("#resetRecipeFilters"),
    recipeResultCount: document.querySelector("#recipeResultCount"),
    recipeMealCount: document.querySelector("#recipeMealCount"),
    recipeFoodCount: document.querySelector("#recipeFoodCount"),
    recipeMealGrid: document.querySelector("#recipeMealGrid"),
    recipeFoodGrid: document.querySelector("#recipeFoodGrid"),
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
    syncButton: document.querySelector("#syncButton"),
    syncButtonLabel: document.querySelector("#syncButtonLabel"),
    syncDot: document.querySelector("#syncDot"),
    syncMenu: document.querySelector("#syncMenu"),
    syncMenuTitle: document.querySelector("#syncMenuTitle"),
    syncMenuText: document.querySelector("#syncMenuText"),
    syncMenuPrimary: document.querySelector("#syncMenuPrimary"),
    syncMenuRenew: document.querySelector("#syncMenuRenew"),
    syncMenuLogout: document.querySelector("#syncMenuLogout"),
    syncPanel: document.querySelector("#syncPanel"),
    syncStatusTitle: document.querySelector("#syncStatusTitle"),
    syncStatusText: document.querySelector("#syncStatusText"),
    syncSecondary: document.querySelector("#syncSecondary"),
    syncRenew: document.querySelector("#syncRenew"),
    syncLogout: document.querySelector("#syncLogout"),
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
  const shoppingPanelHome = dom.shoppingPanel.parentElement;
  let isBrandHomeNavigation = false;

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

  function normalizeFoodName(value) {
    return String(value || "")
      .toLocaleLowerCase("de")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function normalizeMealType(value) {
    const mealType = String(value || "").toLocaleLowerCase("de");
    return mealType === "warm" || mealType === "kalt" ? mealType : "";
  }

  function cleanIds(value, validIds) {
    return [...new Set((Array.isArray(value) ? value : []).map(Number).filter((id) => validIds.has(id)))];
  }

  function cleanSelectedIds(value) {
    return cleanIds(value, validFoodIds);
  }

  function cleanBookmarkedFoodIds(value) {
    return cleanIds(value, validFoodIds);
  }

  function cleanBookmarkedMealIds(value) {
    return cleanIds(value, validMealIds);
  }

  function cleanSharedMealIds(value) {
    return cleanIds(value, validMealIds);
  }

  function cleanSharedFoodIds(value) {
    return cleanIds(value, validFoodIds);
  }

  function cleanRecipeText(value, maxLength) {
    return typeof value === "string" ? value.trim().slice(0, maxLength) : "";
  }

  function cleanRecipes(value) {
    if (!Array.isArray(value)) return [];
    const recipesById = new Map();
    value.forEach((recipe) => {
      if (!recipe || typeof recipe !== "object") return;
      const itemType = recipe.itemType === "meal" ? "meal" : "food";
      const itemId = Number(recipe.itemId ?? recipe.foodId);
      const validItemIds = itemType === "meal" ? validMealIds : validFoodIds;
      const id = cleanRecipeText(recipe.id, 120);
      const title = cleanRecipeText(recipe.title, 160);
      const url = cleanRecipeText(recipe.url, 2048);
      const notes = cleanRecipeText(recipe.notes, 20000);
      if (!validItemIds.has(itemId) || !id || (!title && !url && !notes)) return;
      recipesById.set(id, {
        id,
        itemType,
        itemId,
        ...(itemType === "food" ? { foodId: itemId } : {}),
        title,
        url,
        notes,
        createdAt: cleanRecipeText(recipe.createdAt, 40),
        updatedAt: cleanRecipeText(recipe.updatedAt, 40),
      });
    });
    return [...recipesById.values()];
  }

  function recipesSignature(value) {
    return JSON.stringify(cleanRecipes(value).sort((left, right) => left.id.localeCompare(right.id)));
  }

  function loadSharedMealIds() {
    try {
      return cleanSharedMealIds(JSON.parse(localStorage.getItem(sharedMealStorageKey) || "[]"));
    } catch {
      return [];
    }
  }

  function loadSharedFoodIds() {
    try {
      return cleanSharedFoodIds(JSON.parse(localStorage.getItem(sharedFoodStorageKey) || "[]"));
    } catch {
      return [];
    }
  }

  function loadRecipes() {
    try {
      return cleanRecipes(JSON.parse(localStorage.getItem(recipeStorageKey) || "[]"));
    } catch {
      return [];
    }
  }

  function nowIso() {
    return new Date().toISOString();
  }

  function loadSelectionData() {
    try {
      const selected = cleanSelectedIds(JSON.parse(localStorage.getItem(storageKey) || "[]"));
      const bookmarks = JSON.parse(localStorage.getItem(bookmarkStorageKey) || "{}");
      const bookmarkedFoodIds = cleanBookmarkedFoodIds(bookmarks.foodIds || []);
      const bookmarkedMealIds = cleanBookmarkedMealIds(bookmarks.mealIds || []);
      const sharedMealIds = loadSharedMealIds();
      const sharedFoodIds = loadSharedFoodIds();
      const recipes = loadRecipes();
      const meta = JSON.parse(localStorage.getItem(storageMetaKey) || "{}");
      const hasLocalData = selected.length || bookmarkedFoodIds.length || bookmarkedMealIds.length || sharedMealIds.length || sharedFoodIds.length || recipes.length;
      const updatedAt = typeof meta.updatedAt === "string" ? meta.updatedAt : hasLocalData ? nowIso() : "";
      return { selected, bookmarkedFoodIds, bookmarkedMealIds, sharedMealIds, sharedFoodIds, recipes, updatedAt };
    } catch {
      return { selected: [], bookmarkedFoodIds: [], bookmarkedMealIds: [], sharedMealIds: [], sharedFoodIds: [], recipes: [], updatedAt: "" };
    }
  }

  function saveSelectionLocally(updatedAt = nowIso()) {
    const selected = cleanSelectedIds([...state.selected]);
    const bookmarkedFoodIds = cleanBookmarkedFoodIds([...state.bookmarkedFoods]);
    const bookmarkedMealIds = cleanBookmarkedMealIds([...state.bookmarkedMeals]);
    const sharedMealIds = cleanSharedMealIds([...state.sharedMeals]);
    const sharedFoodIds = cleanSharedFoodIds([...state.sharedFoods]);
    const recipes = cleanRecipes(state.recipes);
    localStorage.setItem(storageKey, JSON.stringify(selected));
    localStorage.setItem(bookmarkStorageKey, JSON.stringify({
      foodIds: bookmarkedFoodIds,
      mealIds: bookmarkedMealIds,
    }));
    localStorage.setItem(sharedMealStorageKey, JSON.stringify(sharedMealIds));
    localStorage.setItem(sharedFoodStorageKey, JSON.stringify(sharedFoodIds));
    localStorage.setItem(recipeStorageKey, JSON.stringify(recipes));
    localStorage.setItem(storageMetaKey, JSON.stringify({ updatedAt }));
    state.localUpdatedAt = updatedAt;
    return { selected, bookmarkedFoodIds, bookmarkedMealIds, sharedMealIds, sharedFoodIds, recipes, updatedAt };
  }

  function persistSelection() {
    saveSelectionLocally();
    queueOneDriveSave();
  }

  function selectionPayload(updatedAt = state.localUpdatedAt || nowIso()) {
    return {
      app: "lebensmitteleinkauf",
      version: 5,
      updatedAt,
      selectedIds: cleanSelectedIds([...state.selected]),
      bookmarkedFoodIds: cleanBookmarkedFoodIds([...state.bookmarkedFoods]),
      bookmarkedMealIds: cleanBookmarkedMealIds([...state.bookmarkedMeals]),
      sharedMealIds: cleanSharedMealIds([...state.sharedMeals]),
      sharedFoodIds: cleanSharedFoodIds([...state.sharedFoods]),
      recipes: cleanRecipes(state.recipes),
    };
  }

  function parseRemoteData(data) {
    if (!data || typeof data !== "object") return null;
    const selectedIds = cleanSelectedIds(data.selectedIds || data.selected || []);
    return {
      app: "lebensmitteleinkauf",
      version: Number(data.version) || 1,
      updatedAt: typeof data.updatedAt === "string" ? data.updatedAt : "",
      selectedIds,
      bookmarkedFoodIds: cleanBookmarkedFoodIds(data.bookmarkedFoodIds || data.bookmarkFoodIds || []),
      bookmarkedMealIds: cleanBookmarkedMealIds(data.bookmarkedMealIds || data.bookmarkMealIds || []),
      sharedMealIds: cleanSharedMealIds(data.sharedMealIds || data.sharedMeals || []),
      hasSharedMealIds: Array.isArray(data.sharedMealIds) || Array.isArray(data.sharedMeals),
      sharedFoodIds: cleanSharedFoodIds(data.sharedFoodIds || data.sharedFoods || []),
      hasSharedFoodIds: Array.isArray(data.sharedFoodIds) || Array.isArray(data.sharedFoods),
      recipes: cleanRecipes(data.recipes || data.foodRecipes || []),
      hasRecipes: Array.isArray(data.recipes) || Array.isArray(data.foodRecipes),
    };
  }

  function sameUserData(left, right) {
    const leftData = Array.isArray(left) ? { selectedIds: left } : left || {};
    const rightData = Array.isArray(right) ? { selectedIds: right } : right || {};
    return sameCoreUserData(leftData, rightData)
      && cleanSharedMealIds(leftData.sharedMealIds).sort((a, b) => a - b).join(",") === cleanSharedMealIds(rightData.sharedMealIds).sort((a, b) => a - b).join(",")
      && cleanSharedFoodIds(leftData.sharedFoodIds).sort((a, b) => a - b).join(",") === cleanSharedFoodIds(rightData.sharedFoodIds).sort((a, b) => a - b).join(",")
      && recipesSignature(leftData.recipes) === recipesSignature(rightData.recipes);
  }

  function sameCoreUserData(left, right) {
    const leftData = Array.isArray(left) ? { selectedIds: left } : left || {};
    const rightData = Array.isArray(right) ? { selectedIds: right } : right || {};
    return cleanSelectedIds(leftData.selectedIds).sort((a, b) => a - b).join(",") === cleanSelectedIds(rightData.selectedIds).sort((a, b) => a - b).join(",")
      && cleanBookmarkedFoodIds(leftData.bookmarkedFoodIds).sort((a, b) => a - b).join(",") === cleanBookmarkedFoodIds(rightData.bookmarkedFoodIds).sort((a, b) => a - b).join(",")
      && cleanBookmarkedMealIds(leftData.bookmarkedMealIds).sort((a, b) => a - b).join(",") === cleanBookmarkedMealIds(rightData.bookmarkedMealIds).sort((a, b) => a - b).join(",");
  }

  function remoteIsNewer(remoteUpdatedAt, knownUpdatedAt) {
    if (!remoteUpdatedAt || !knownUpdatedAt) return false;
    return Date.parse(remoteUpdatedAt) > Date.parse(knownUpdatedAt);
  }

  function setSyncStatus(status, title, message) {
    if (hasOneDriveManualLogout() && status !== "local" && title !== "OneDrive-Abmeldung") {
      state.sync.status = "local";
      state.sync.title = "Nicht angemeldet";
      state.sync.message = "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert.";
      renderSyncStatus();
      return;
    }
    state.sync.status = status;
    state.sync.title = title;
    state.sync.message = message;
    renderSyncStatus();
  }

  function getSyncMenuText() {
    if (!state.sync.account) {
      const helpText = "Zuerst auf OneDrive anmelden klicken.\nNur bei zu langer Anmeldedauer auf Anmeldung erneuern klicken.";
      if (state.sync.title === "Nicht angemeldet") {
        return `Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert.\n${helpText}`;
      }
      return `${state.sync.message}\n${helpText}`;
    }
    return state.sync.message;
  }

  function renderSyncStatus() {
    if (!dom.syncButton) return;
    const manuallyLoggedOut = hasOneDriveManualLogout();
    dom.syncButton.dataset.syncStatus = state.sync.status;
    if (dom.syncPanel) dom.syncPanel.dataset.syncStatus = state.sync.status;
    if (dom.syncMenu) dom.syncMenu.dataset.syncStatus = state.sync.status;
    if (dom.syncStatusTitle) dom.syncStatusTitle.textContent = state.sync.title;
    if (dom.syncStatusText) dom.syncStatusText.textContent = state.sync.message;
    if (dom.syncMenuTitle) dom.syncMenuTitle.textContent = state.sync.title;
    if (dom.syncMenuText) dom.syncMenuText.textContent = getSyncMenuText();
    dom.syncButton.disabled = false;
    if (dom.syncSecondary) dom.syncSecondary.disabled = state.sync.busy;
    if (dom.syncRenew) dom.syncRenew.disabled = false;
    if (dom.syncLogout) {
      dom.syncLogout.disabled = false;
      dom.syncLogout.hidden = !state.sync.account;
    }
    if (dom.syncMenuPrimary) dom.syncMenuPrimary.disabled = state.sync.busy;
    if (dom.syncMenuRenew) dom.syncMenuRenew.disabled = false;
    if (dom.syncMenuLogout) {
      dom.syncMenuLogout.disabled = false;
      dom.syncMenuLogout.hidden = false;
    }
    dom.syncButtonLabel.textContent = state.sync.account && !manuallyLoggedOut ? "OneDrive" : "Anmelden";

    let actionText;
    if (manuallyLoggedOut) {
      actionText = "Mit OneDrive anmelden";
    } else if (state.sync.status === "conflict") {
      actionText = "OneDrive laden";
    } else if (state.sync.needsInteractiveToken) {
      actionText = "OneDrive bestätigen";
    } else if (state.sync.account) {
      actionText = state.sync.busy ? "Synchronisiert ..." : "Jetzt synchronisieren";
    } else {
      actionText = state.sync.busy ? "Anmeldung ..." : "Mit OneDrive anmelden";
    }
    if (dom.syncSecondary) dom.syncSecondary.textContent = actionText;
    if (dom.syncMenuPrimary) dom.syncMenuPrimary.textContent = actionText;
  }

  function explainAuthError(error) {
    const text = `${error?.errorCode || ""} ${error?.message || ""}`.toLowerCase();
    const code = error?.errorCode || error?.error || error?.name || "";
    const suffix = code ? ` (${code})` : "";
    if (text.includes("redirect-started")) return "Du wirst zu Microsoft weitergeleitet.";
    if (text.includes("interaction_in_progress")) return "Die Microsoft-Anmeldung ist noch blockiert. Starte die OneDrive-Anmeldung bitte erneut.";
    if (text.includes("popup") || text.includes("block")) return "Die Microsoft-Anmeldung wurde vom Browser blockiert. Starte die Anmeldung bitte erneut.";
    if (text.includes("user_cancelled") || text.includes("cancel")) return "Anmeldung oder Zustimmung wurde abgebrochen.";
    if (text.includes("consent") || text.includes("access_denied")) return "Zustimmung verweigert. OneDrive-Sync bleibt ausgeschaltet.";
    if (text.includes("interaction_required")) return "Bitte melde dich erneut an, damit OneDrive verwendet werden darf.";
    return `Microsoft-Anmeldung fehlgeschlagen${suffix}. Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte bleiben lokal gespeichert.`;
  }

  function clearLoginPending() {
    localStorage.removeItem(pendingLoginKey);
    sessionStorage.removeItem(authReloadKey);
    sessionStorage.removeItem(authConfirmAfterReloadKey);
  }

  function markOneDriveManualLogout() {
    localStorage.setItem(manualLogoutKey, String(Date.now()));
  }

  function clearOneDriveManualLogout() {
    localStorage.removeItem(manualLogoutKey);
  }

  function hasOneDriveManualLogout() {
    return Boolean(localStorage.getItem(manualLogoutKey));
  }

  function clearStaleMsalInteractionStatus() {
    [localStorage, sessionStorage].forEach((storage) => {
      for (let index = storage.length - 1; index >= 0; index -= 1) {
        const key = storage.key(index) || "";
        const isMsalInteractionKey = key.includes("interaction.status") && (key.includes(msalConfig.auth.clientId) || key.startsWith("msal."));
        if (isMsalInteractionKey) storage.removeItem(key);
      }
    });
  }

  function clearOneDriveMsalCache() {
    const clientId = msalConfig.auth.clientId.toLowerCase();
    const looksLikeMsalValue = (value) => {
      if (!value || value[0] !== "{") return false;
      try {
        const parsed = JSON.parse(value);
        return Boolean(
          parsed.clientId === msalConfig.auth.clientId
          || parsed.homeAccountId
          || parsed.localAccountId
          || parsed.credentialType
          || parsed.authorityType
          || parsed.environment === "login.windows.net"
          || parsed.environment === "login.microsoftonline.com"
        );
      } catch (error) {
        return false;
      }
    };
    [localStorage, sessionStorage].forEach((storage) => {
      for (let index = storage.length - 1; index >= 0; index -= 1) {
        const key = storage.key(index) || "";
        const normalizedKey = key.toLowerCase();
        const value = storage.getItem(key) || "";
        const isMsalKey = normalizedKey.startsWith("msal.")
          || normalizedKey.includes(clientId)
          || normalizedKey.includes("login.microsoftonline.com")
          || normalizedKey.includes("login.windows.net");
        if (isMsalKey || looksLikeMsalValue(value)) storage.removeItem(key);
      }
    });
  }

  function clearOneDriveMsalCookies() {
    const clientId = msalConfig.auth.clientId.toLowerCase();
    document.cookie.split(";").forEach((cookie) => {
      const name = cookie.split("=")[0]?.trim();
      if (!name) return;
      const normalizedName = name.toLowerCase();
      if (!normalizedName.startsWith("msal.") && !normalizedName.includes(clientId)) return;
      document.cookie = `${name}=; expires=Thu, 01 Jan 1970 00:00:00 GMT; path=/; SameSite=Lax`;
      document.cookie = `${name}=; max-age=0; path=/; SameSite=Lax`;
    });
  }

  function isInteractionInProgressError(error) {
    const text = `${error?.errorCode || ""} ${error?.message || ""}`.toLowerCase();
    return text.includes("interaction_in_progress");
  }

  async function redirectWithFreshInteraction(startRedirect) {
    try {
      await startRedirect();
    } catch (error) {
      if (!isInteractionInProgressError(error)) throw error;
      clearStaleMsalInteractionStatus();
      await startRedirect();
    }
  }

  function hasOneDriveRedirectResponse() {
    const queryParams = new URLSearchParams(window.location.search);
    const hashText = window.location.hash.startsWith("#") ? window.location.hash.slice(1) : window.location.hash;
    const hashParams = new URLSearchParams(hashText);
    return ["code", "error", "state", "client_info"].some((key) => queryParams.has(key) || hashParams.has(key));
  }

  function clearOneDriveRedirectResponseUrl() {
    if (!hasOneDriveRedirectResponse()) return;
    const url = new URL(window.location.href);
    ["code", "error", "error_description", "state", "client_info", "session_state"].forEach((key) => url.searchParams.delete(key));
    const hashText = url.hash.startsWith("#") ? url.hash.slice(1) : url.hash;
    const hashParams = new URLSearchParams(hashText);
    const hasAuthHash = ["code", "error", "state", "client_info"].some((key) => hashParams.has(key));
    if (hasAuthHash) url.hash = "";
    const cleanUrl = `${url.pathname}${url.search}${url.hash}`;
    if (window.history?.replaceState) window.history.replaceState(null, document.title, cleanUrl);
  }

  function clearAuthReloadParam() {
    const url = new URL(window.location.href);
    const hasAuthParam = url.searchParams.has(authReloadParam) || url.searchParams.has(authLogoutParam) || url.searchParams.has(appRefreshParam);
    if (!hasAuthParam) return;
    url.searchParams.delete(authReloadParam);
    url.searchParams.delete(authLogoutParam);
    url.searchParams.delete(appRefreshParam);
    const cleanUrl = `${url.pathname}${url.search}${url.hash}`;
    if (window.history?.replaceState) window.history.replaceState(null, document.title, cleanUrl);
  }

  function reloadForFreshAppVersion(latestVersion) {
    const refreshUrl = new URL(window.location.href);
    refreshUrl.hash = "";
    refreshUrl.searchParams.set(appRefreshParam, latestVersion);
    window.location.replace(refreshUrl.toString());
    setTimeout(() => {
      if (window.location.href !== refreshUrl.toString()) window.location.href = refreshUrl.toString();
    }, 250);
  }

  async function fetchLatestAppVersion() {
    const controller = typeof AbortController === "function" ? new AbortController() : null;
    const timeoutId = controller ? setTimeout(() => controller.abort(), 1200) : null;
    try {
      return await fetch(`${appVersionFile}?ts=${Date.now()}`, {
        cache: "no-store",
        ...(controller ? { signal: controller.signal } : {}),
      });
    } finally {
      if (timeoutId) clearTimeout(timeoutId);
    }
  }

  async function ensureLatestAppVersion() {
    if (hasOneDriveRedirectResponse()) return false;
    try {
      const response = await fetchLatestAppVersion();
      if (!response.ok) return false;
      const data = await response.json();
      const latestVersion = String(data?.appVersion || "").trim();
      if (!latestVersion || latestVersion === appVersion) {
        if (sessionStorage.getItem(appRefreshSessionKey) === latestVersion) {
          sessionStorage.removeItem(appRefreshSessionKey);
        }
        return false;
      }
      if (sessionStorage.getItem(appRefreshSessionKey) === latestVersion) return false;
      sessionStorage.setItem(appRefreshSessionKey, latestVersion);
      reloadForFreshAppVersion(latestVersion);
      return true;
    } catch (error) {
      return false;
    }
  }

  function forceOneDriveLogoutReload() {
    const refreshUrl = new URL(window.location.href);
    refreshUrl.hash = "";
    refreshUrl.searchParams.delete(authReloadParam);
    refreshUrl.searchParams.set(authLogoutParam, String(Date.now()));
    window.location.replace(refreshUrl.toString());
    setTimeout(() => {
      if (window.location.href !== refreshUrl.toString()) window.location.href = refreshUrl.toString();
    }, 250);
  }

  function rememberRedirectToken(response) {
    if (!response?.accessToken) return;
    state.sync.redirectAccessToken = response.accessToken;
    state.sync.redirectAccessTokenExpiresAt = response.expiresOn instanceof Date
      ? response.expiresOn.getTime()
      : Date.now() + 45 * 60 * 1000;
  }

  function getRememberedRedirectToken() {
    if (!state.sync.redirectAccessToken) return "";
    if (Date.now() > state.sync.redirectAccessTokenExpiresAt - 60 * 1000) {
      state.sync.redirectAccessToken = "";
      state.sync.redirectAccessTokenExpiresAt = 0;
      return "";
    }
    return state.sync.redirectAccessToken;
  }

  async function getGraphToken() {
    if (hasOneDriveManualLogout()) throw new Error("not-signed-in");
    if (!state.sync.account) throw new Error("not-signed-in");
    const request = { ...loginRequest, account: state.sync.account };
    const redirectToken = getRememberedRedirectToken();
    if (redirectToken) return redirectToken;
    try {
      const response = await state.sync.msal.acquireTokenSilent(request);
      return response.accessToken;
    } catch (error) {
      if (!state.sync.allowInteractiveTokenRedirect) {
        const tokenError = new Error("interactive-token-required");
        tokenError.cause = error;
        throw tokenError;
      }
      setSyncStatus("loading", "Microsoft-Anmeldung", "Du wirst zu Microsoft weitergeleitet.");
      await redirectWithFreshInteraction(() => state.sync.msal.acquireTokenRedirect({ ...request, redirectStartPage: window.location.href }));
      throw new Error("redirect-started");
    }
  }

  async function graphFetch(path, options = {}) {
    const token = await getGraphToken();
    const response = await fetch(`${graphBaseUrl}${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(options.headers || {}),
      },
    });
    if (response.status === 404) return null;
    if (!response.ok) {
      const detail = await response.text().catch(() => "");
      const error = new Error(detail || `Graph request failed: ${response.status}`);
      error.status = response.status;
      throw error;
    }
    return response;
  }

  async function loadRemoteSelection() {
    const metadataResponse = await graphFetch(graphFilePath);
    if (!metadataResponse) return { exists: false, data: null, etag: "" };
    const metadata = await metadataResponse.json();
    const contentResponse = await graphFetch(graphContentPath, { cache: "no-store" });
    if (!contentResponse) return { exists: false, data: null, etag: metadata.eTag || "" };
    const data = parseRemoteData(await contentResponse.json());
    return {
      exists: Boolean(data),
      data,
      etag: metadata.eTag || "",
    };
  }

  async function uploadRemoteSelection(payload) {
    const response = await graphFetch(graphContentPath, {
      method: "PUT",
      headers: { "Content-Type": "application/json;charset=utf-8" },
      body: JSON.stringify(payload, null, 2),
    });
    return response ? response.json() : null;
  }

  function applyRemoteSelection(remoteData, remoteEtag = "") {
    state.selected = new Set(remoteData.selectedIds);
    state.bookmarkedFoods = new Set(remoteData.bookmarkedFoodIds);
    state.bookmarkedMeals = new Set(remoteData.bookmarkedMealIds);
    state.sharedMeals = new Set(remoteData.sharedMealIds || []);
    state.sharedFoods = new Set(remoteData.sharedFoodIds || []);
    state.recipes = cleanRecipes(remoteData.recipes || []);
    saveSelectionLocally(remoteData.updatedAt || nowIso());
    state.sync.lastRemoteUpdatedAt = remoteData.updatedAt || state.localUpdatedAt;
    state.sync.lastRemoteEtag = remoteEtag;
    state.sync.hasRemoteData = true;
    state.sync.conflictData = null;
    renderFoods();
    renderShoppingList();
    renderMeals();
    renderBookmarks();
    renderRecipeOverview();
    renderShares();
  }

  function completeRemoteSave(payload, metadata) {
    state.sync.lastRemoteUpdatedAt = payload.updatedAt;
    state.sync.lastRemoteEtag = metadata?.eTag || state.sync.lastRemoteEtag;
    state.sync.hasRemoteData = true;
    state.sync.conflictData = null;
    state.sync.needsInteractiveToken = false;
    state.pendingSharedMeals.clear();
    state.pendingSharedFoods.clear();
    setSyncStatus("synced", "Mit OneDrive synchronisiert", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte sind im OneDrive-App-Ordner gespeichert.");
  }

  function handleOneDriveError(error, fallbackTitle = "OneDrive nicht verfügbar") {
    if (error?.message === "redirect-started") return;
    if (error?.message === "not-signed-in") {
      state.sync.needsInteractiveToken = false;
      setSyncStatus("local", "Nicht angemeldet", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert.");
      return;
    }
    if (error?.message === "interactive-token-required") {
      state.sync.needsInteractiveToken = true;
      const message = "OneDrive ist angemeldet, braucht aber noch eine aktive Bestätigung. Tippe auf \"OneDrive bestätigen\", um die Berechtigung zu erteilen.";
      setSyncStatus("error", "OneDrive-Bestätigung nötig", message);
      showToast("OneDrive-Bestätigung nötig.");
      return;
    }
    const message = error?.status === 401 || error?.status === 403
      ? "Zugriff auf OneDrive wurde nicht erlaubt. Bitte erneut anmelden."
      : "OneDrive konnte nicht erreicht werden. Lokale Daten bleiben erhalten.";
    setSyncStatus("error", fallbackTitle, message);
    showToast(message);
  }

  async function confirmOneDriveAccess() {
    if (!state.sync.msal || !state.sync.account || state.sync.busy) return;
    clearLoginPending();
    clearStaleMsalInteractionStatus();
    state.sync.busy = true;
    state.sync.needsInteractiveToken = false;
    setSyncStatus("loading", "Microsoft-Anmeldung", "Du wirst zu Microsoft weitergeleitet, um OneDrive zu bestätigen.");
    try {
      await redirectWithFreshInteraction(() => state.sync.msal.acquireTokenRedirect({
        ...loginRequest,
        account: state.sync.account,
        redirectStartPage: window.location.href,
      }));
    } catch (error) {
      state.sync.needsInteractiveToken = true;
      setSyncStatus("error", "Bestätigung fehlgeschlagen", explainAuthError(error));
      showToast(explainAuthError(error));
    } finally {
      state.sync.busy = false;
      renderSyncStatus();
    }
  }

  function remoteShouldReplaceLocal(remoteData, localData) {
    if (!remoteData) return false;
    if (!localData.selected.length && !localData.bookmarkedFoodIds.length && !localData.bookmarkedMealIds.length && !localData.sharedMealIds.length && !localData.sharedFoodIds.length && !localData.recipes.length) return true;
    if (remoteIsNewer(remoteData.updatedAt, localData.updatedAt)) return true;
    return Boolean(remoteData.updatedAt && !localData.updatedAt);
  }

  async function migrateLegacyRemoteSharedMeals(remote, local, preferRemoteData) {
    const mergedRemoteData = preferRemoteData
      ? { ...remote.data }
      : {
          ...remote.data,
          updatedAt: local.updatedAt,
          selectedIds: local.selected,
          bookmarkedFoodIds: local.bookmarkedFoodIds,
          bookmarkedMealIds: local.bookmarkedMealIds,
        };
    mergedRemoteData.sharedMealIds = cleanSharedMealIds(local.sharedMealIds);
    mergedRemoteData.hasSharedMealIds = true;
    if (!remote.data.hasSharedFoodIds) {
      mergedRemoteData.sharedFoodIds = cleanSharedFoodIds(local.sharedFoodIds);
      mergedRemoteData.hasSharedFoodIds = true;
    }
    if (!remote.data.hasRecipes) {
      mergedRemoteData.recipes = cleanRecipes(local.recipes);
      mergedRemoteData.hasRecipes = true;
    }
    applyRemoteSelection(mergedRemoteData, remote.etag);
    const payload = selectionPayload(nowIso());
    saveSelectionLocally(payload.updatedAt);
    const metadata = await uploadRemoteSelection(payload);
    completeRemoteSave(payload, metadata);
  }

  async function migrateLegacyRemoteSharedFoods(remote, local, preferRemoteData) {
    const mergedRemoteData = preferRemoteData
      ? { ...remote.data }
      : {
          ...remote.data,
          updatedAt: local.updatedAt,
          selectedIds: local.selected,
          bookmarkedFoodIds: local.bookmarkedFoodIds,
          bookmarkedMealIds: local.bookmarkedMealIds,
          sharedMealIds: local.sharedMealIds,
        };
    mergedRemoteData.sharedFoodIds = cleanSharedFoodIds(local.sharedFoodIds);
    mergedRemoteData.hasSharedFoodIds = true;
    if (!remote.data.hasRecipes) {
      mergedRemoteData.recipes = cleanRecipes(local.recipes);
      mergedRemoteData.hasRecipes = true;
    }
    applyRemoteSelection(mergedRemoteData, remote.etag);
    const payload = selectionPayload(nowIso());
    saveSelectionLocally(payload.updatedAt);
    const metadata = await uploadRemoteSelection(payload);
    completeRemoteSave(payload, metadata);
  }

  async function migrateLegacyRemoteRecipes(remote, local, preferRemoteData) {
    const mergedRemoteData = preferRemoteData
      ? { ...remote.data }
      : {
          ...remote.data,
          updatedAt: local.updatedAt,
          selectedIds: local.selected,
          bookmarkedFoodIds: local.bookmarkedFoodIds,
          bookmarkedMealIds: local.bookmarkedMealIds,
          sharedMealIds: local.sharedMealIds,
          sharedFoodIds: local.sharedFoodIds,
        };
    mergedRemoteData.recipes = cleanRecipes(local.recipes);
    mergedRemoteData.hasRecipes = true;
    applyRemoteSelection(mergedRemoteData, remote.etag);
    const payload = selectionPayload(nowIso());
    saveSelectionLocally(payload.updatedAt);
    const metadata = await uploadRemoteSelection(payload);
    completeRemoteSave(payload, metadata);
  }

  async function addPendingSharedEntriesToRemote(remote) {
    const sharedMealIds = cleanSharedMealIds([
      ...(remote.data.sharedMealIds || []),
      ...state.pendingSharedMeals,
    ]);
    const sharedFoodIds = cleanSharedFoodIds([
      ...(remote.data.sharedFoodIds || []),
      ...state.pendingSharedFoods,
    ]);
    applyRemoteSelection({
      ...remote.data,
      sharedMealIds,
      hasSharedMealIds: true,
      sharedFoodIds,
      hasSharedFoodIds: true,
    }, remote.etag);
    const payload = selectionPayload(nowIso());
    saveSelectionLocally(payload.updatedAt);
    const metadata = await uploadRemoteSelection(payload);
    completeRemoteSave(payload, metadata);
  }

  async function saveSelectionToOneDrive() {
    if (!state.sync.account || state.sync.busy) return;
    state.sync.busy = true;
    setSyncStatus("saving", "Speichere in OneDrive", "Prüfe zuerst, ob dort neuere Daten liegen.");
    try {
      const latest = await loadRemoteSelection();
      if (
        latest.exists
        && state.sync.lastRemoteUpdatedAt
        && remoteIsNewer(latest.data.updatedAt, state.sync.lastRemoteUpdatedAt)
        && !sameUserData(latest.data, selectionPayload())
      ) {
        state.sync.conflictData = latest;
        state.sync.lastRemoteUpdatedAt = latest.data.updatedAt || state.sync.lastRemoteUpdatedAt;
        state.sync.lastRemoteEtag = latest.etag || state.sync.lastRemoteEtag;
        setSyncStatus("conflict", "Konflikt erkannt", "OneDrive enthält neuere Daten. Es wurde nichts überschrieben.");
        showToast("Konflikt erkannt: OneDrive wurde nicht überschrieben.");
        return;
      }

      const payload = selectionPayload(nowIso());
      saveSelectionLocally(payload.updatedAt);
      const metadata = await uploadRemoteSelection(payload);
      completeRemoteSave(payload, metadata);
    } catch (error) {
      handleOneDriveError(error, "Speichern fehlgeschlagen");
    } finally {
      state.sync.busy = false;
      renderSyncStatus();
    }
  }

  async function manualSyncOneDrive() {
    if (!state.sync.account || state.sync.busy) return;
    state.sync.busy = true;
    const previousInteractiveTokenRedirect = state.sync.allowInteractiveTokenRedirect;
    state.sync.allowInteractiveTokenRedirect = true;
    setSyncStatus("loading", "Prüfe OneDrive", "Einkaufsliste, Lesezeichen, geteilte Mahlzeiten und Lebensmittel sowie Rezepte werden mit OneDrive abgeglichen.");
    try {
      const remote = await loadRemoteSelection();
      const local = loadSelectionData();

      if (remote.exists && !remote.data.hasSharedMealIds && local.sharedMealIds.length) {
        const preferRemoteData = remoteShouldReplaceLocal(remote.data, local)
          || (!remoteIsNewer(local.updatedAt, remote.data.updatedAt) && !sameCoreUserData(remote.data, {
            selectedIds: local.selected,
            bookmarkedFoodIds: local.bookmarkedFoodIds,
            bookmarkedMealIds: local.bookmarkedMealIds,
          }));
        await migrateLegacyRemoteSharedMeals(remote, local, preferRemoteData);
        showToast("Geteilte Mahlzeiten wurden nach OneDrive übernommen.");
        return;
      }

      if (remote.exists && !remote.data.hasSharedFoodIds && local.sharedFoodIds.length) {
        const preferRemoteData = remoteShouldReplaceLocal(remote.data, local)
          || (!remoteIsNewer(local.updatedAt, remote.data.updatedAt) && !sameCoreUserData(remote.data, {
            selectedIds: local.selected,
            bookmarkedFoodIds: local.bookmarkedFoodIds,
            bookmarkedMealIds: local.bookmarkedMealIds,
          }));
        await migrateLegacyRemoteSharedFoods(remote, local, preferRemoteData);
        showToast("Geteilte Lebensmittel wurden nach OneDrive übernommen.");
        return;
      }

      if (remote.exists && !remote.data.hasRecipes && local.recipes.length) {
        const preferRemoteData = remoteShouldReplaceLocal(remote.data, local)
          || (!remoteIsNewer(local.updatedAt, remote.data.updatedAt) && !sameCoreUserData(remote.data, {
            selectedIds: local.selected,
            bookmarkedFoodIds: local.bookmarkedFoodIds,
            bookmarkedMealIds: local.bookmarkedMealIds,
          }));
        await migrateLegacyRemoteRecipes(remote, local, preferRemoteData);
        showToast("Deine Rezepte wurden nach OneDrive übernommen.");
        return;
      }

      if (remote.exists && (state.pendingSharedMeals.size || state.pendingSharedFoods.size)) {
        await addPendingSharedEntriesToRemote(remote);
        showToast("Geöffneter Teilen-Link wurde nach OneDrive übernommen.");
        return;
      }

      if (remote.exists && remoteShouldReplaceLocal(remote.data, local)) {
        applyRemoteSelection(remote.data, remote.etag);
        state.sync.needsInteractiveToken = false;
        setSyncStatus("synced", "Mit OneDrive synchronisiert", "Neuere Daten aus OneDrive wurden übernommen.");
        showToast("Neuere Daten aus OneDrive wurden übernommen.");
        return;
      }

      if (remote.exists && sameUserData(remote.data, selectionPayload())) {
        state.sync.lastRemoteUpdatedAt = remote.data.updatedAt || state.sync.lastRemoteUpdatedAt;
        state.sync.lastRemoteEtag = remote.etag || state.sync.lastRemoteEtag;
        state.sync.needsInteractiveToken = false;
        setSyncStatus("synced", "Mit OneDrive synchronisiert", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte sind aktuell.");
        return;
      }

      if (!remote.exists || remoteIsNewer(local.updatedAt, remote.data?.updatedAt)) {
        const payload = selectionPayload(nowIso());
        saveSelectionLocally(payload.updatedAt);
        const metadata = await uploadRemoteSelection(payload);
        completeRemoteSave(payload, metadata);
        showToast("Einkaufsliste, Lesezeichen, geteilte Mahlzeiten und Lebensmittel sowie Rezepte wurden nach OneDrive gespeichert.");
        return;
      }

      if (remote.exists) {
        applyRemoteSelection(remote.data, remote.etag);
        state.sync.needsInteractiveToken = false;
        setSyncStatus("synced", "Mit OneDrive synchronisiert", "OneDrive-Daten wurden übernommen.");
      }
    } catch (error) {
      handleOneDriveError(error, "Synchronisieren fehlgeschlagen");
    } finally {
      state.sync.allowInteractiveTokenRedirect = previousInteractiveTokenRedirect;
      state.sync.busy = false;
      renderSyncStatus();
    }
  }

  let oneDriveSaveTimer;

  function queueOneDriveSave() {
    if (hasOneDriveManualLogout() || !state.sync.account) {
      setSyncStatus("local", "Nicht angemeldet", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert.");
      return;
    }
    clearTimeout(oneDriveSaveTimer);
    setSyncStatus("saving", "Speichern vorbereitet", "Die Änderung wird gleich mit OneDrive synchronisiert.");
    oneDriveSaveTimer = setTimeout(() => { void saveSelectionToOneDrive(); }, 650);
  }

  async function syncFromOneDrive({ forceRemote = false, allowInteractiveTokenRedirect = false } = {}) {
    if (hasOneDriveManualLogout()) return;
    if (!state.sync.account || state.sync.busy) return;
    state.sync.busy = true;
    const previousInteractiveTokenRedirect = state.sync.allowInteractiveTokenRedirect;
    state.sync.allowInteractiveTokenRedirect = allowInteractiveTokenRedirect;
    setSyncStatus("loading", "Prüfe OneDrive", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden geladen.");
    try {
      const remote = forceRemote && state.sync.conflictData ? state.sync.conflictData : await loadRemoteSelection();
      const local = loadSelectionData();

      if (!remote.exists) {
        if (local.selected.length || local.bookmarkedFoodIds.length || local.bookmarkedMealIds.length || local.sharedMealIds.length || local.sharedFoodIds.length || local.recipes.length) {
          const payload = selectionPayload(local.updatedAt || nowIso());
          const metadata = await uploadRemoteSelection(payload);
          completeRemoteSave(payload, metadata);
          showToast("Lokale Einkaufsliste, Lesezeichen, geteilte Mahlzeiten und Lebensmittel sowie Rezepte wurden nach OneDrive übernommen.");
        } else {
          state.sync.hasRemoteData = false;
          state.sync.needsInteractiveToken = false;
          setSyncStatus("synced", "Mit OneDrive verbunden", "Noch keine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel oder Rezepte im OneDrive-App-Ordner.");
        }
        return;
      }

      const remoteMatchesLocalCore = sameCoreUserData(remote.data, {
        selectedIds: local.selected,
        bookmarkedFoodIds: local.bookmarkedFoodIds,
        bookmarkedMealIds: local.bookmarkedMealIds,
      });
      if (
        !remote.data.hasSharedMealIds
        && local.sharedMealIds.length
        && (forceRemote || remoteShouldReplaceLocal(remote.data, local) || remoteMatchesLocalCore)
      ) {
        await migrateLegacyRemoteSharedMeals(remote, local, forceRemote || remoteShouldReplaceLocal(remote.data, local));
        showToast("Geteilte Mahlzeiten wurden nach OneDrive übernommen.");
        return;
      }

      if (
        !remote.data.hasSharedFoodIds
        && local.sharedFoodIds.length
        && (forceRemote || remoteShouldReplaceLocal(remote.data, local) || remoteMatchesLocalCore)
      ) {
        await migrateLegacyRemoteSharedFoods(remote, local, forceRemote || remoteShouldReplaceLocal(remote.data, local));
        showToast("Geteilte Lebensmittel wurden nach OneDrive übernommen.");
        return;
      }

      if (
        !remote.data.hasRecipes
        && local.recipes.length
        && (forceRemote || remoteShouldReplaceLocal(remote.data, local) || remoteMatchesLocalCore)
      ) {
        await migrateLegacyRemoteRecipes(remote, local, forceRemote || remoteShouldReplaceLocal(remote.data, local));
        showToast("Deine Rezepte wurden nach OneDrive übernommen.");
        return;
      }

      if (state.pendingSharedMeals.size || state.pendingSharedFoods.size) {
        await addPendingSharedEntriesToRemote(remote);
        showToast("Geöffneter Teilen-Link wurde nach OneDrive übernommen.");
        return;
      }

      if (
        forceRemote
        || (!local.selected.length && !local.bookmarkedFoodIds.length && !local.bookmarkedMealIds.length && !local.sharedMealIds.length && !local.sharedFoodIds.length && !local.recipes.length)
        || remoteIsNewer(remote.data.updatedAt, local.updatedAt)
        || sameUserData(remote.data, {
          selectedIds: local.selected,
          bookmarkedFoodIds: local.bookmarkedFoodIds,
          bookmarkedMealIds: local.bookmarkedMealIds,
          sharedMealIds: local.sharedMealIds,
          sharedFoodIds: local.sharedFoodIds,
          recipes: local.recipes,
        })
      ) {
        applyRemoteSelection(remote.data, remote.etag);
        state.sync.needsInteractiveToken = false;
        setSyncStatus("synced", "Mit OneDrive synchronisiert", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte wurden aus OneDrive geladen.");
        return;
      }

      state.sync.lastRemoteUpdatedAt = remote.data.updatedAt || "";
      state.sync.lastRemoteEtag = remote.etag || "";
      state.sync.hasRemoteData = true;
      state.sync.conflictData = remote;
      state.sync.needsInteractiveToken = false;
      setSyncStatus("conflict", "Unterschiedliche Daten", "Lokale Daten und OneDrive-Daten unterscheiden sich. Es wurde nichts überschrieben.");
    } catch (error) {
      handleOneDriveError(error);
    } finally {
      state.sync.allowInteractiveTokenRedirect = previousInteractiveTokenRedirect;
      state.sync.busy = false;
      renderSyncStatus();
    }
  }

  async function loginToOneDrive() {
    if (!state.sync.msal || state.sync.busy) return;
    clearOneDriveManualLogout();
    clearLoginPending();
    clearStaleMsalInteractionStatus();
    state.sync.busy = true;
    setSyncStatus("loading", "Microsoft-Anmeldung", "Du wirst zu Microsoft weitergeleitet.");
    try {
      await redirectWithFreshInteraction(() => state.sync.msal.loginRedirect({ ...loginRequest, redirectStartPage: window.location.href }));
    } catch (error) {
      setSyncStatus("error", "Anmeldung fehlgeschlagen", explainAuthError(error));
      showToast(explainAuthError(error));
    } finally {
      state.sync.busy = false;
      renderSyncStatus();
    }
  }

  function logoutFromOneDrive(event) {
    event?.preventDefault();
    event?.stopPropagation();
    markOneDriveManualLogout();
    clearLoginPending();
    clearTimeout(oneDriveSaveTimer);
    if (state.sync.busy) state.sync.busy = false;
    state.sync.busy = true;
    setSyncStatus("loading", "OneDrive-Abmeldung", "Die Verbindung zu OneDrive wird lokal getrennt.");

    try {
      const account = state.sync.account || state.sync.msal?.getActiveAccount?.() || state.sync.msal?.getAllAccounts?.()?.[0];
      clearLoginPending();
      clearStaleMsalInteractionStatus();
      clearOneDriveMsalCache();
      clearOneDriveMsalCookies();
      state.sync.account = null;
      state.sync.lastRemoteUpdatedAt = "";
      state.sync.lastRemoteEtag = "";
      state.sync.hasRemoteData = false;
      state.sync.conflictData = null;
      state.sync.needsInteractiveToken = false;
      state.sync.redirectAccessToken = "";
      state.sync.redirectAccessTokenExpiresAt = 0;
      state.sync.allowInteractiveTokenRedirect = false;
      state.sync.msal?.setActiveAccount?.(null);
      if (state.sync.msal?.clearCache) {
        Promise.resolve(state.sync.msal.clearCache(account ? { account } : {})).catch(() => {});
      }
      setSyncStatus("local", "Nicht angemeldet", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert.");
      state.sync.busy = false;
      renderSyncStatus();
      showToast("OneDrive wurde abgemeldet.");
      forceOneDriveLogoutReload();
    } catch (error) {
      clearLoginPending();
      clearOneDriveMsalCache();
      clearOneDriveMsalCookies();
      state.sync.account = null;
      state.sync.busy = false;
      setSyncStatus("local", "Nicht angemeldet", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert.");
      showToast("OneDrive wurde lokal getrennt.");
      forceOneDriveLogoutReload();
    }
  }

  function closeSyncMenu() {
    if (!dom.syncMenu || !dom.syncButton) return;
    dom.syncMenu.hidden = true;
    dom.syncButton.setAttribute("aria-expanded", "false");
  }

  function toggleSyncMenu() {
    if (!dom.syncMenu || !dom.syncButton) return;
    const open = dom.syncMenu.hidden;
    dom.syncMenu.hidden = !open;
    dom.syncButton.setAttribute("aria-expanded", open ? "true" : "false");
  }

  function runOneDrivePrimaryAction(event) {
    event?.preventDefault();
    event?.stopPropagation();
    if (state.sync.needsInteractiveToken) confirmOneDriveAccess();
    else if (state.sync.conflictData) syncFromOneDrive({ forceRemote: true });
    else if (state.sync.account) manualSyncOneDrive();
    else loginToOneDrive();
  }

  function renewOneDriveLogin(event) {
    event?.preventDefault();
    event?.stopPropagation();
    location.reload();
  }

  async function initializeOneDrive() {
    renderSyncStatus();
    if (!window.msal?.PublicClientApplication) {
      setSyncStatus("error", "OneDrive nicht verfügbar", "MSAL.js konnte nicht geladen werden. Lokale Speicherung bleibt aktiv.");
      return;
    }

    try {
      state.sync.msal = new window.msal.PublicClientApplication(msalConfig);
      if (typeof state.sync.msal.initialize === "function") {
        await state.sync.msal.initialize();
      }
      if (hasOneDriveManualLogout()) {
        clearLoginPending();
        clearStaleMsalInteractionStatus();
        clearOneDriveMsalCache();
        clearOneDriveMsalCookies();
        state.sync.account = null;
        state.sync.allowInteractiveTokenRedirect = false;
        state.sync.msal.setActiveAccount?.(null);
        setSyncStatus("local", "Nicht angemeldet", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert.");
        return;
      }
      const redirectResponse = await state.sync.msal.handleRedirectPromise();
      rememberRedirectToken(redirectResponse);
      clearOneDriveRedirectResponseUrl();
      const accounts = state.sync.msal.getAllAccounts();
      state.sync.account = redirectResponse?.account || accounts[0] || null;
      if (state.sync.account) {
        clearLoginPending();
        state.sync.msal.setActiveAccount(state.sync.account);
        await syncFromOneDrive();
      } else {
        setSyncStatus("local", "Nicht angemeldet", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert. Melde dich an, um OneDrive zu nutzen.");
      }
    } catch (error) {
      if (hasOneDriveManualLogout()) {
        clearOneDriveRedirectResponseUrl();
        clearOneDriveMsalCache();
        clearOneDriveMsalCookies();
        state.sync.account = null;
        setSyncStatus("local", "Nicht angemeldet", "Deine Einkaufsliste, Lesezeichen, geteilten Mahlzeiten und Lebensmittel sowie Rezepte werden lokal auf diesem Gerät gespeichert.");
        return;
      }
      const accounts = state.sync.msal?.getAllAccounts?.() || [];
      state.sync.account = accounts[0] || null;
      if (state.sync.account) {
        clearLoginPending();
        state.sync.msal.setActiveAccount(state.sync.account);
        setSyncStatus("loading", "Microsoft-Anmeldung erkannt", "OneDrive wird erneut geprüft.");
        await syncFromOneDrive();
      } else {
        clearOneDriveRedirectResponseUrl();
        clearLoginPending();
        setSyncStatus("error", "Anmeldung fehlgeschlagen", explainAuthError(error));
      }
    } finally {
      state.sync.initialized = true;
      renderSyncStatus();
    }
  }

  function populateFilters() {
    const categories = [...new Set(foods.map((food) => food.category))]
      .sort((a, b) => a.localeCompare(b, "de", { sensitivity: "base" }));
    const categoryOptions = categories.map((category) => `<option value="${escapeHtml(category)}">${escapeHtml(category)}</option>`).join("");
    dom.categoryFilter.insertAdjacentHTML("beforeend", categoryOptions);
    dom.mealCategoryFilter.insertAdjacentHTML("beforeend", categoryOptions);
    dom.recipeCategoryFilter.insertAdjacentHTML("beforeend", categoryOptions);
  }

  function populateRecipeItemFilter() {
    const currentValue = state.recipeItemKey;
    const recipes = cleanRecipes(state.recipes);
    const mealsWithRecipes = [...new Set(recipes.filter((recipe) => recipe.itemType === "meal").map((recipe) => recipe.itemId))]
      .map((mealId) => meals.find((meal) => meal.id === mealId))
      .filter(Boolean)
      .sort((left, right) => left.situation.localeCompare(right.situation, "de", { sensitivity: "base" }));
    const foodsWithRecipes = [...new Set(recipes.filter((recipe) => recipe.itemType === "food").map((recipe) => recipe.itemId))]
      .map((foodId) => foodById.get(foodId))
      .filter(Boolean)
      .sort((left, right) => left.name.localeCompare(right.name, "de", { sensitivity: "base" }));
    dom.recipeItemFilter.innerHTML = '<option value="">Alle Mahlzeiten und Lebensmittel</option>'
      + (mealsWithRecipes.length ? `<optgroup label="Mahlzeiten">${mealsWithRecipes.map((meal) => `<option value="meal:${meal.id}">${escapeHtml(meal.situation)}</option>`).join("")}</optgroup>` : "")
      + (foodsWithRecipes.length ? `<optgroup label="Lebensmittel">${foodsWithRecipes.map((food) => `<option value="food:${food.id}">${escapeHtml(food.name)}</option>`).join("")}</optgroup>` : "");
    const validValues = new Set([
      ...mealsWithRecipes.map((meal) => `meal:${meal.id}`),
      ...foodsWithRecipes.map((food) => `food:${food.id}`),
    ]);
    if (currentValue && validValues.has(currentValue)) {
      dom.recipeItemFilter.value = currentValue;
    } else {
      state.recipeItemKey = "";
      dom.recipeItemFilter.value = "";
    }
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

  function renderBookmarkButton(kind, item) {
    const isMeal = kind === "meal";
    const bookmarked = isMeal ? state.bookmarkedMeals.has(item.id) : state.bookmarkedFoods.has(item.id);
    const itemType = isMeal ? "Mahlzeit" : "Lebensmittel";
    const itemName = isMeal ? item.situation : item.name;
    const action = bookmarked ? "Lesezeichen entfernen" : "Lesezeichen setzen";
    return `
      <button class="bookmark-button${bookmarked ? " is-bookmarked" : ""}" type="button"
        data-bookmark-kind="${kind}" data-bookmark-id="${item.id}" aria-pressed="${bookmarked}"
        aria-label="${escapeHtml(itemType)} ${escapeHtml(itemName)}: ${action}" title="${action}">
        <span class="button-icon">${icon("bookmark")}</span>
        <span class="bookmark-button-label">${bookmarked ? "Entfernen" : "Merken"}</span>
      </button>`;
  }

  function renderFoodCard(food, options = {}) {
    const selected = state.selected.has(food.id);
    return `
      <article class="food-card${selected ? " is-selected" : ""}${options.isShareTarget ? " is-share-target" : ""}" data-id="${food.id}"${options.isShareTarget ? ' tabindex="-1"' : ""}>
        ${options.isShareView ? `
          <label class="share-select-control food-share-select-control">
            <input type="checkbox" data-shared-food-select value="${food.id}"${options.isShareSelected ? " checked" : ""} aria-label="${escapeHtml(food.name)} zum Entfernen auswählen" />
            <span>${options.isShareSelected ? "Ausgewählt" : "Zum Entfernen auswählen"}</span>
          </label>` : ""}
        <button class="select-food" type="button" aria-label="${escapeHtml(food.name)} ${selected ? "von der Einkaufsliste entfernen" : "zur Einkaufsliste hinzufügen"}" aria-pressed="${selected}"></button>
        <div class="food-card-body">
          <div class="food-card-top">
            <div>
              <span class="category-label"><span class="category-symbol">${categoryIcon(food.category)}</span>${escapeHtml(food.subcategory)}</span>
              <h3>${escapeHtml(food.name)}</h3>
            </div>
            <div class="food-card-tools">
              <span class="priority-dot"><i></i>${escapeHtml(food.priority)}</span>
              ${renderBookmarkButton("food", food)}
            </div>
          </div>
          <div class="food-meta">
            <span class="meta-pill score">Sättigung ${food.score}/5 · ${escapeHtml(food.satiety)}</span>
            <span class="meta-pill">KH ${escapeHtml(food.carbs)}</span>
          </div>
          <p class="food-benefit">${escapeHtml(food.benefit)}</p>
          <div class="food-card-footer">
            <span class="compound">${escapeHtml(food.compounds)}</span>
            <div class="food-card-actions">
              <button class="details-button" type="button">Details →</button>
              <button class="food-meals-button" type="button" aria-label="Mahlzeiten mit ${escapeHtml(food.name)} anzeigen">Mahlzeiten →</button>
            </div>
          </div>
          <div class="food-personal-actions">
            <button class="food-share-button" type="button" data-share-food-id="${food.id}" aria-label="Teilen-Link für ${escapeHtml(food.name)} kopieren">zum Teilen →</button>
            <button class="food-recipes-button" type="button" aria-label="Eigene Rezepte für ${escapeHtml(food.name)} verwalten">meine Rezepte →</button>
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
    dom.loadMore.hidden = visible.length >= result.length;
    renderActiveFilters(result.length);
  }

  function renderActiveFilters(resultCount) {
    const chips = [];
    if (state.search) chips.push(`Suche: ${state.search}`);
    if (state.category) chips.push(state.category);
    if (state.score) chips.push(`Score ${state.score}`);
    if (state.priority) chips.push(`Priorität: ${state.priority}`);
    dom.activeFilters.innerHTML = chips.length
      ? `${chips.map((chip) => `<span class="filter-chip">${escapeHtml(chip)}</span>`).join("")}<span class="filter-result-count">${resultCount} ${resultCount === 1 ? "Eintrag" : "Einträge"}</span>`
      : "";
  }

  function resetFoodFilters() {
    state.search = "";
    state.category = "";
    state.score = "";
    state.priority = "";
    state.limit = window.innerWidth < 680 ? 18 : 28;

    dom.searchInput.value = "";
    dom.categoryFilter.selectedIndex = 0;
    dom.scoreFilter.selectedIndex = 0;
    dom.priorityFilter.selectedIndex = 0;

    renderFoods();
  }

  function resetMealFilters() {
    state.mealSearch = "";
    state.mealType = "";
    state.mealCategory = "";

    dom.mealSearchInput.value = "";
    dom.mealTypeFilter.selectedIndex = 0;
    dom.mealCategoryFilter.selectedIndex = 0;

    renderMeals();
  }

  function resetRecipeFilters() {
    state.recipeSearch = "";
    state.recipeCategory = "";
    state.recipeItemKey = "";

    dom.recipeSearchInput.value = "";
    dom.recipeCategoryFilter.selectedIndex = 0;
    dom.recipeItemFilter.selectedIndex = 0;

    renderRecipeOverview();
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
          <p>Tippe bei einem Lebensmittel auf das Kästchen. Mit OneDrive-Anmeldung bleibt deine Auswahl geräteübergreifend synchron.</p>
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
    renderMeals();
    renderBookmarks();
    renderShares();
  }

  function toggleBookmark(kind, id) {
    const isMeal = kind === "meal";
    const target = isMeal ? state.bookmarkedMeals : state.bookmarkedFoods;
    const validIds = isMeal ? validMealIds : validFoodIds;
    if (!validIds.has(id)) return;
    const wasBookmarked = target.has(id);
    if (wasBookmarked) target.delete(id);
    else target.add(id);
    persistSelection();
    renderFoods();
    renderMeals();
    renderBookmarks();
    renderShares();
    showToast(wasBookmarked ? "Lesezeichen wurde entfernt." : "Lesezeichen wurde gesetzt.");
  }

  function mealIngredientFoods(meal) {
    return (meal.ingredients || []).map((ingredient) => foodByName.get(normalizeFoodName(ingredient))).filter(Boolean);
  }

  function mealVariantFoods(meal) {
    const variantText = ` ${normalizeFoodName(meal.variants)} `;
    const ingredientNames = new Set((meal.ingredients || []).map(normalizeFoodName));
    return foods
      .map((food) => ({
        food,
        position: variantText.indexOf(` ${normalizeFoodName(food.name)} `),
      }))
      .filter(({ food, position }) => position >= 0 && !ingredientNames.has(normalizeFoodName(food.name)))
      .sort((left, right) => left.position - right.position)
      .map(({ food }) => food);
  }

  function addMealIngredients(mealId, ingredientNames) {
    const meal = meals.find((item) => item.id === mealId);
    if (!meal) return;
    const names = ingredientNames || meal.ingredients || [];
    const ingredients = names.map((name) => foodByName.get(normalizeFoodName(name))).filter(Boolean);
    const previousSize = state.selected.size;
    ingredients.forEach((food) => state.selected.add(food.id));
    const addedCount = state.selected.size - previousSize;
    persistSelection();
    renderFoods();
    renderShoppingList();
    renderMeals();
    renderShares();
    showToast(addedCount ? `${addedCount} Zutaten wurden auf die Liste gesetzt.` : "Alle Zutaten sind bereits auf der Liste.");
  }

  function mealChoice(name, checked) {
    return `
      <label class="meal-choice">
        <input type="checkbox" name="meal-food" value="${escapeHtml(name)}"${checked ? " checked" : ""} />
        <span>${escapeHtml(name)}</span>
      </label>`;
  }

  function openMealVariation(mealId, action) {
    const meal = meals.find((item) => item.id === mealId);
    if (!meal) return;
    const ingredients = meal.ingredients || [];
    const variants = mealVariantFoods(meal).map((food) => food.name);
    const actionLabel = action === "recipe"
      ? "für Rezeptsuche →"
      : action === "offers" ? "Suche Sonderangebote →" : "Auf die Liste →";
    dom.detailDialog.classList.remove("is-image-dialog");
    dom.detailDialog.classList.add("is-meal-variation-dialog");
    dom.detailDialog.classList.remove("is-recipe-dialog");
    dom.detailContent.innerHTML = `
      <div class="detail-content meal-variation-content">
        <div class="detail-icon">${icon("meal")}</div>
        <p class="eyebrow">Mahlzeit variieren</p>
        <h2>${escapeHtml(meal.situation)}</h2>
        <p class="detail-subtitle">Wähle die Lebensmittel aus, die berücksichtigt werden sollen.</p>
        <form class="meal-variation-form" data-meal-variation-form data-meal-id="${meal.id}" data-meal-action="${action}">
          <section class="detail-section meal-choice-section">
            <h3>Zutaten</h3>
            <div class="meal-choice-list">${ingredients.map((name) => mealChoice(name, true)).join("")}</div>
          </section>
          <section class="detail-section meal-choice-section">
            <h3>Varianten</h3>
            ${variants.length
              ? `<div class="meal-choice-list">${variants.map((name) => mealChoice(name, false)).join("")}</div>`
              : '<p class="meal-choice-empty">Für diese Mahlzeit sind keine konkreten Lebensmittel als Varianten hinterlegt.</p>'}
          </section>
          <div class="meal-variation-footer">
            <p class="meal-selection-count" aria-live="polite">${ingredients.length} Lebensmittel ausgewählt</p>
            <button class="primary-button meal-variation-action" type="submit">${actionLabel}</button>
          </div>
        </form>
      </div>`;
    dom.detailDialog.showModal();
  }

  function openMealOffersPostalCode(mealId, ingredientNames) {
    const meal = meals.find((item) => item.id === mealId);
    const names = ingredientNames || meal?.ingredients || [];
    if (!meal || !names.length) return;
    dom.detailDialog.classList.remove("is-image-dialog");
    dom.detailDialog.classList.remove("is-meal-variation-dialog");
    dom.detailDialog.classList.remove("is-recipe-dialog");
    dom.detailContent.innerHTML = `
      <div class="detail-content meal-offers-postal-content">
        <div class="detail-icon">${icon("meal")}</div>
        <p class="eyebrow">Suche Sonderangebote</p>
        <h2>Postleitzahl eingeben</h2>
        <p class="detail-subtitle">Damit die Suche auf lokale Geschäfte in deiner Nähe eingegrenzt werden kann.</p>
        <form class="meal-offers-postal-form" data-meal-offers-postal-form data-meal-id="${meal.id}">
          ${names.map((name) => `<input type="hidden" name="meal-offers-food" value="${escapeHtml(name)}" />`).join("")}
          <label class="meal-offers-postal-field" for="mealOffersPostalCode">
            <span>Deine Postleitzahl</span>
            <input id="mealOffersPostalCode" name="postal-code" type="text" inputmode="numeric" autocomplete="postal-code" pattern="[0-9]{5}" minlength="5" maxlength="5" placeholder="z. B. 10115" required autofocus />
          </label>
          <p class="meal-offers-postal-hint">Bitte gib eine fünfstellige deutsche Postleitzahl ein. Sie wird nur in den kopierten Suchtext eingefügt.</p>
          <section class="detail-section meal-offers-food-summary">
            <h3>Ausgewählte Lebensmittel</h3>
            <p>${names.map((name) => escapeHtml(name)).join(", ")}</p>
          </section>
          <div class="meal-variation-footer">
            <p class="meal-selection-count">${names.length} Lebensmittel werden berücksichtigt</p>
            <button class="primary-button meal-variation-action" type="submit">Text kopieren →</button>
          </div>
        </form>
      </div>`;
    dom.detailDialog.showModal();
  }

  function runMealAction(mealId, action, ingredientNames) {
    if (action === "recipe") void copyMealRecipeSearch(mealId, ingredientNames);
    else if (action === "offers") openMealOffersPostalCode(mealId, ingredientNames);
    else addMealIngredients(mealId, ingredientNames);
  }

  function requestMealAction(mealId, action) {
    const meal = meals.find((item) => item.id === mealId);
    if (!meal) return;
    openConfirm({
      title: "Mahlzeit variieren?",
      text: "Möchtest du die Lebensmittel für diese Aktion selbst auswählen?",
      cancel: "Nein",
      accept: "Ja",
      tone: "cancel-primary",
      cancelAction: () => runMealAction(mealId, action, meal.ingredients || []),
      action: () => openMealVariation(mealId, action),
    });
  }

  function openDetails(id) {
    const food = foods.find((item) => item.id === id);
    if (!food) return;
    const links = String(food.sources).split(";").map((url) => url.trim()).filter(Boolean);
    dom.detailDialog.classList.remove("is-image-dialog");
    dom.detailDialog.classList.remove("is-meal-variation-dialog");
    dom.detailDialog.classList.remove("is-recipe-dialog");
    dom.detailContent.innerHTML = `
      <div class="detail-content">
        <div class="detail-icon">${categoryIcon(food.category)}</div>
        <p class="eyebrow">${escapeHtml(food.category)}</p>
        <h2>${escapeHtml(food.name)}</h2>
        <p class="detail-subtitle">${escapeHtml(food.subcategory)}</p>
        <div class="detail-badges">
          <span class="meta-pill score">Sättigung ${food.score}/5 · ${escapeHtml(food.satiety)}</span>
          <span class="meta-pill">Kohlenhydrate ${escapeHtml(food.carbs)}</span>
          <span class="meta-pill">Empfehlung ${escapeHtml(food.priority)}</span>
        </div>
        ${detailSection("Warum sinnvoll", food.benefit)}
        ${detailSection("Aktive Stoffe", food.compounds)}
        ${detailSection("Praktische Verwendung", food.use)}
        ${detailSection("Hinweise / Einschränkungen", food.caution)}
        ${detailSection("Nährwerte", food.nutrition, "ohne Gewähr")}
        ${detailSection("Menge pro Portion", food.portion, "Durchschnittswerte - individueller Bedarf oder Kombinationen mit weiteren Lebensmitteln können zu anderen Mengen führen.")}
        ${links.length ? `<section class="detail-section"><h3>Quellen</h3><div class="detail-sources">${links.map((url, index) => `<a href="${escapeHtml(url)}" target="_blank" rel="noreferrer">Quelle ${index + 1} ↗</a>`).join("")}</div></section>` : ""}
      </div>`;
    dom.detailDialog.showModal();
  }

  function createRecipeId() {
    if (window.crypto?.randomUUID) return window.crypto.randomUUID();
    return `recipe-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
  }

  function safeRecipeUrl(value) {
    try {
      const url = new URL(String(value || "").trim());
      return url.protocol === "http:" || url.protocol === "https:" ? url.href : "";
    } catch {
      return "";
    }
  }

  function recipeDisplayTitle(recipe) {
    if (recipe.title) return recipe.title;
    const url = safeRecipeUrl(recipe.url);
    if (url) {
      try {
        return new URL(url).hostname.replace(/^www\./, "");
      } catch {
        return "Gespeichertes Rezept";
      }
    }
    return "Rezept ohne Überschrift";
  }

  function renderSavedRecipe(recipe) {
    const url = safeRecipeUrl(recipe.url);
    return `
      <article class="saved-recipe-card">
        <div class="saved-recipe-head">
          <h3>${escapeHtml(recipeDisplayTitle(recipe))}</h3>
          <div class="saved-recipe-actions">
            <button class="recipe-edit-button" type="button" data-edit-recipe-id="${escapeHtml(recipe.id)}">Bearbeiten</button>
            <button class="recipe-delete-button" type="button" data-delete-recipe-id="${escapeHtml(recipe.id)}">Löschen</button>
          </div>
        </div>
        ${url ? `<a class="saved-recipe-link" href="${escapeHtml(url)}" target="_blank" rel="noreferrer">Rezept öffnen ↗</a>` : recipe.url ? `<p class="saved-recipe-url">${escapeHtml(recipe.url)}</p>` : ""}
        ${recipe.notes ? `<p class="saved-recipe-notes">${escapeHtml(recipe.notes).replace(/\n/g, "<br>")}</p>` : ""}
      </article>`;
  }

  function recipeTarget(itemType, itemId) {
    if (itemType === "meal") {
      const meal = meals.find((item) => item.id === itemId);
      return meal ? {
        itemType: "meal",
        itemId: meal.id,
        name: meal.situation,
        subtitle: "Speichere Links, eigene Notizen oder ganze Rezepte passend zu dieser Mahlzeit.",
        detail: meal.date ? `Mahlzeit vom ${formatMealDate(meal.date)}` : "Mahlzeit",
        icon: icon("meal"),
      } : null;
    }
    const food = foodById.get(itemId);
    return food ? {
      itemType: "food",
      itemId: food.id,
      name: food.name,
      subtitle: "Speichere Links, eigene Notizen oder ganze Rezepte passend zu diesem Lebensmittel.",
      detail: food.subcategory,
      icon: categoryIcon(food.category),
    } : null;
  }

  function openRecipes(itemType, itemId, editRecipeId = "") {
    const target = recipeTarget(itemType, itemId);
    if (!target) return;
    const recipes = cleanRecipes(state.recipes)
      .filter((recipe) => recipe.itemType === target.itemType && recipe.itemId === target.itemId)
      .sort((left, right) => String(right.updatedAt).localeCompare(String(left.updatedAt)));
    const recipeToEdit = recipes.find((recipe) => recipe.id === editRecipeId) || null;
    dom.detailDialog.classList.remove("is-image-dialog");
    dom.detailDialog.classList.remove("is-meal-variation-dialog");
    dom.detailDialog.classList.add("is-recipe-dialog");
    dom.detailContent.innerHTML = `
      <div class="detail-content recipe-content">
        <div class="detail-icon">${target.icon}</div>
        <p class="eyebrow">${escapeHtml(target.name)}</p>
        <h2>Meine Rezepte</h2>
        <p class="detail-subtitle">${escapeHtml(target.subtitle)}</p>
        <form class="recipe-form" data-recipe-form data-recipe-item-type="${target.itemType}" data-recipe-item-id="${target.itemId}">
          <input type="hidden" name="recipe-id" value="${escapeHtml(recipeToEdit?.id || "")}" />
          <label class="recipe-field">
            <span>Überschrift:</span>
            <input name="recipe-title" type="text" maxlength="160" autocomplete="off" value="${escapeHtml(recipeToEdit?.title || "")}" placeholder="z. B. Ofenlachs mit Kräutern" />
          </label>
          <label class="recipe-field">
            <span>URL:</span>
            <input name="recipe-url" type="url" maxlength="2048" inputmode="url" autocomplete="url" value="${escapeHtml(recipeToEdit?.url || "")}" placeholder="https://…" />
          </label>
          <label class="recipe-field">
            <span>Anmerkungen:</span>
            <textarea name="recipe-notes" maxlength="20000" rows="6" placeholder="Eigene Hinweise oder das vollständige Rezept …">${escapeHtml(recipeToEdit?.notes || "")}</textarea>
          </label>
          <p class="recipe-form-hint">Mindestens ein Feld muss ausgefüllt sein. Ohne OneDrive-Anmeldung bleiben deine Rezepte nur in diesem Browser gespeichert.</p>
          <div class="recipe-form-actions">
            ${recipeToEdit ? '<button class="secondary-button recipe-cancel-edit" type="button">Bearbeiten abbrechen</button>' : ""}
            <button class="primary-button recipe-save-button" type="submit"${recipeToEdit ? "" : " disabled"}>${recipeToEdit ? "Änderungen speichern" : "Rezept speichern"}</button>
          </div>
        </form>
        <section class="recipe-library" aria-labelledby="savedRecipesTitle">
          <div class="recipe-library-heading">
            <h3 id="savedRecipesTitle">Gespeicherte Rezepte</h3>
            <span>${recipes.length}</span>
          </div>
          <div class="saved-recipe-list">
            ${recipes.length ? recipes.map(renderSavedRecipe).join("") : '<div class="recipe-empty"><strong>Noch keine Rezepte gespeichert.</strong><p>Fülle oben mindestens eines der drei Felder aus.</p></div>'}
          </div>
        </section>
      </div>`;
    if (!dom.detailDialog.open) dom.detailDialog.showModal();
    if (recipeToEdit) dom.detailContent.querySelector('[name="recipe-title"]')?.focus();
  }

  function openFoodRecipes(foodId, editRecipeId = "") {
    openRecipes("food", foodId, editRecipeId);
  }

  function openMealRecipes(mealId, editRecipeId = "") {
    openRecipes("meal", mealId, editRecipeId);
  }

  function saveRecipe(form) {
    const itemType = form.dataset.recipeItemType === "meal" ? "meal" : "food";
    const itemId = Number(form.dataset.recipeItemId);
    const validItemIds = itemType === "meal" ? validMealIds : validFoodIds;
    if (!validItemIds.has(itemId)) return;
    const recipeId = String(form.elements["recipe-id"].value || "");
    const title = cleanRecipeText(form.elements["recipe-title"].value, 160);
    const url = cleanRecipeText(form.elements["recipe-url"].value, 2048);
    const notes = cleanRecipeText(form.elements["recipe-notes"].value, 20000);
    if (!title && !url && !notes) return;
    if (!form.checkValidity()) {
      form.reportValidity();
      return;
    }
    const existing = state.recipes.find((recipe) => recipe.id === recipeId && recipe.itemType === itemType && recipe.itemId === itemId);
    const timestamp = nowIso();
    const savedRecipe = {
      id: existing?.id || createRecipeId(),
      itemType,
      itemId,
      title,
      url,
      notes,
      createdAt: existing?.createdAt || timestamp,
      updatedAt: timestamp,
    };
    state.recipes = cleanRecipes(existing
      ? state.recipes.map((recipe) => recipe.id === existing.id ? savedRecipe : recipe)
      : [...state.recipes, savedRecipe]);
    persistSelection();
    renderRecipeOverview();
    openRecipes(itemType, itemId);
    showToast(existing ? "Rezept wurde geändert." : "Rezept wurde gespeichert.");
  }

  function requestRecipeDelete(recipeId, { reopenManager = true } = {}) {
    const recipe = state.recipes.find((item) => item.id === recipeId);
    if (!recipe) return;
    openConfirm({
      title: "Rezept löschen?",
      text: `„${recipeDisplayTitle(recipe)}“ wird dauerhaft gelöscht${state.sync.account ? " und die Änderung mit OneDrive synchronisiert" : " und aus diesem Browser entfernt"}.`,
      accept: "Rezept löschen",
      action: () => {
        state.recipes = state.recipes.filter((item) => item.id !== recipe.id);
        persistSelection();
        renderRecipeOverview();
        if (reopenManager) openRecipes(recipe.itemType, recipe.itemId);
        showToast("Rezept wurde gelöscht.");
      },
    });
  }

  function openAllFoods() {
    const names = Array.isArray(foodNames) && foodNames.length
      ? foodNames
      : [...new Set(foods.map((food) => food.name).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, "de", { sensitivity: "base" }));
    dom.detailDialog.classList.remove("is-image-dialog");
    dom.detailDialog.classList.remove("is-meal-variation-dialog");
    dom.detailContent.innerHTML = `
      <div class="detail-content">
        <div class="detail-icon">${icon("leaf")}</div>
        <p class="eyebrow">Alle Lebensmittel</p>
        <h2>Alle Lebensmittel</h2>
        <p class="detail-subtitle">${names.length} Lebensmittel alphabetisch sortiert.</p>
        <section class="detail-section all-foods-section">
          <h3>Alphabetische Liste</h3>
          <ul class="all-foods-list">${names.map((name) => `<li><button class="all-foods-list-button" type="button" data-food-search="${escapeHtml(name)}">${escapeHtml(name)}</button></li>`).join("")}</ul>
        </section>
      </div>`;
    dom.detailDialog.showModal();
  }

  function applyFoodSearchFromList(name) {
    const term = String(name || "").trim();
    if (!term) return;

    state.search = term;
    state.mealSearch = term;
    state.category = "";
    state.score = "";
    state.priority = "";
    state.mealType = "";
    state.mealCategory = "";
    state.limit = window.innerWidth < 680 ? 18 : 28;

    dom.searchInput.value = term;
    dom.mealSearchInput.value = term;
    dom.categoryFilter.selectedIndex = 0;
    dom.scoreFilter.selectedIndex = 0;
    dom.priorityFilter.selectedIndex = 0;
    dom.mealTypeFilter.selectedIndex = 0;
    dom.mealCategoryFilter.selectedIndex = 0;

    renderFoods();
    renderMeals();
    dom.detailDialog.close();
    dom.searchInput.focus({ preventScroll: true });
  }

  function applyMealSearchFromFood(name) {
    const term = String(name || "").trim();
    if (!term) return;

    state.mealSearch = term;
    state.mealType = "";
    state.mealCategory = "";

    dom.mealSearchInput.value = term;
    dom.mealTypeFilter.selectedIndex = 0;
    dom.mealCategoryFilter.selectedIndex = 0;

    renderMeals();
    setView("meals");
    dom.mealSearchInput.focus({ preventScroll: true });
  }

  function openMealGuideImage(step) {
    const guideImage = mealGuideImages[step];
    if (!guideImage) return;
    dom.detailDialog.classList.add("is-image-dialog");
    dom.detailDialog.classList.remove("is-meal-variation-dialog");
    dom.detailDialog.classList.remove("is-recipe-dialog");
    dom.detailContent.innerHTML = `
      <div class="meal-guide-image-content">
        <h2 class="sr-only">Bildanleitung zu Schritt ${step}</h2>
        <img class="meal-guide-dialog-image" src="${escapeHtml(guideImage.src)}" alt="${escapeHtml(guideImage.alt)}" />
      </div>`;
    dom.detailDialog.showModal();
  }

  function detailSection(title, value, note = "") {
    if (!value) return "";
    return `<section class="detail-section"><h3>${escapeHtml(title)}</h3>${note ? `<p class="detail-note">${escapeHtml(note)}</p>` : ""}<p>${escapeHtml(value)}</p></section>`;
  }

  function formatMealDate(value) {
    const [year, month, day] = String(value || "").split("-");
    return year && month && day ? `${day}.${month}.${year}` : String(value || "");
  }

  function renderMealCard(meal, options = {}) {
    const mealDate = formatMealDate(meal.date);
    const mealNumber = (mealIndexById.get(meal.id) ?? 0) + 1;
    const ingredientFoods = mealIngredientFoods(meal);
    const allIngredientsSelected = ingredientFoods.length > 0 && ingredientFoods.every((food) => state.selected.has(food.id));
    return `
      <article class="meal-card${options.isShareTarget ? " is-share-target" : ""}" data-id="${meal.id}"${options.isShareTarget ? ' tabindex="-1"' : ""}>
        ${options.isShareView ? `
          <label class="share-select-control">
            <input type="checkbox" data-shared-meal-select value="${meal.id}"${options.isShareSelected ? " checked" : ""} aria-label="${escapeHtml(meal.situation)} zum Entfernen auswählen" />
            <span>${options.isShareSelected ? "Ausgewählt" : "Zum Entfernen auswählen"}</span>
          </label>` : ""}
        <div class="meal-meta">
          ${renderBookmarkButton("meal", meal)}
          ${mealDate ? `<time class="meal-date" datetime="${escapeHtml(meal.date)}">${escapeHtml(mealDate)}</time>` : ""}
          <span class="meal-number">${String(mealNumber).padStart(2, "0")}</span>
        </div>
        <p class="eyebrow">${escapeHtml(meal.satiety)} sättigend</p>
        <h2>${escapeHtml(meal.situation)}</h2>
        <p class="meal-ingredients"><strong>Zutaten:</strong> ${(meal.ingredients || []).map(escapeHtml).join(", ")}</p>
        ${meal.variants ? `<div class="meal-variants"><span>Varianten</span><p>${escapeHtml(meal.variants)}</p></div>` : ""}
        <div class="meal-details">
          <div><span>Warum sinnvoll</span><p>${escapeHtml(meal.reason)}</p></div>
        </div>
        <div class="meal-card-footer">
          <button class="meal-list-button${allIngredientsSelected ? " is-added" : ""}" type="button" data-meal-id="${meal.id}" aria-pressed="${allIngredientsSelected}">${allIngredientsSelected ? "Auf der Liste ✓" : "Auf die Liste →"}</button>
        </div>
        <div class="meal-share-row">
          <button class="meal-offers-button" type="button" data-offers-meal-id="${meal.id}" aria-label="Sonderangebotssuchtext für ${escapeHtml(meal.situation)} kopieren">Suche Sonderangebote →</button>
          <button class="meal-recipe-button" type="button" data-recipe-meal-id="${meal.id}" aria-label="Rezeptsuchtext für ${escapeHtml(meal.situation)} kopieren">für Rezeptsuche →</button>
        </div>
        <div class="meal-personal-actions">
          <button class="meal-share-button" type="button" data-share-meal-id="${meal.id}" aria-label="Teilen-Link für ${escapeHtml(meal.situation)} kopieren">zum Teilen →</button>
          <button class="meal-personal-recipes-button" type="button" data-open-meal-recipes="${meal.id}" aria-label="Eigene Rezepte für ${escapeHtml(meal.situation)} verwalten">meine Rezepte →</button>
        </div>
      </article>`;
  }

  function renderMeals() {
    const term = state.mealSearch.trim().toLocaleLowerCase("de");
    const visibleMeals = meals
      .filter((meal) => {
        const searchable = `${meal.situation} ${meal.satiety} ${(meal.ingredients || []).join(" ")} ${meal.reason} ${meal.variants} ${formatMealDate(meal.date)}`.toLocaleLowerCase("de");
        const mealType = normalizeMealType(meal.mealType);
        const ingredientFoods = mealIngredientFoods(meal);
        return (!term || searchable.includes(term))
          && (!state.mealType || mealType === state.mealType)
          && (!state.mealCategory || ingredientFoods.some((food) => food.category === state.mealCategory));
      })
      .reverse();

    const hasActiveMealFilter = Boolean(state.mealSearch || state.mealType || state.mealCategory);
    dom.mealResultCount.hidden = !hasActiveMealFilter;
    dom.mealResultCount.textContent = hasActiveMealFilter
      ? `${visibleMeals.length} ${visibleMeals.length === 1 ? "Mahlzeit" : "Mahlzeiten"}`
      : "";

    dom.mealGrid.innerHTML = visibleMeals.length
      ? visibleMeals.map(renderMealCard).join("")
      : '<div class="empty-results">Keine passenden Empfehlungen gefunden.</div>';
  }

  function renderBookmarks() {
    const bookmarkedMeals = meals
      .filter((meal) => state.bookmarkedMeals.has(meal.id))
      .sort((left, right) => right.id - left.id);
    const bookmarkedFoods = foods
      .filter((food) => state.bookmarkedFoods.has(food.id))
      .sort((left, right) => right.id - left.id);

    dom.bookmarkedMealCount.textContent = `${bookmarkedMeals.length} ${bookmarkedMeals.length === 1 ? "Lesezeichen" : "Lesezeichen"}`;
    dom.bookmarkedFoodCount.textContent = `${bookmarkedFoods.length} ${bookmarkedFoods.length === 1 ? "Lesezeichen" : "Lesezeichen"}`;
    dom.clearAllBookmarks.disabled = !bookmarkedMeals.length && !bookmarkedFoods.length;
    dom.bookmarkedMealGrid.innerHTML = bookmarkedMeals.length
      ? bookmarkedMeals.map(renderMealCard).join("")
      : '<div class="bookmark-empty"><span class="button-icon">' + icon("bookmark") + '</span><strong>Noch keine Mahlzeit gespeichert</strong><p>Setze bei einer Mahlzeit ein Lesezeichen, um sie hier wiederzufinden.</p></div>';
    dom.bookmarkedFoodGrid.innerHTML = bookmarkedFoods.length
      ? bookmarkedFoods.map(renderFoodCard).join("")
      : '<div class="bookmark-empty"><span class="button-icon">' + icon("bookmark") + '</span><strong>Noch kein Lebensmittel gespeichert</strong><p>Setze bei einem Lebensmittel ein Lesezeichen, um es hier wiederzufinden.</p></div>';
  }

  function recipeCreatedTimestamp(recipe) {
    const timestamp = Date.parse(recipe.createdAt || "");
    return Number.isFinite(timestamp) ? timestamp : 0;
  }

  function formatRecipeCreatedAt(recipe) {
    const timestamp = recipeCreatedTimestamp(recipe);
    if (!timestamp) return "Erfassungsdatum nicht verfügbar";
    return `Erfasst am ${new Intl.DateTimeFormat("de-DE", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    }).format(new Date(timestamp))}`;
  }

  function renderRecipeOverviewCard(recipe, target) {
    const url = safeRecipeUrl(recipe.url);
    return `
      <article class="recipe-overview-card" data-recipe-id="${escapeHtml(recipe.id)}">
        <div class="recipe-overview-card-head">
          <div>
            <button class="recipe-food-link" type="button" data-open-recipe-item-type="${target.itemType}" data-open-recipe-item-id="${target.itemId}">${escapeHtml(target.name)}</button>
            <h3>${escapeHtml(recipeDisplayTitle(recipe))}</h3>
          </div>
          <span class="recipe-created-date">${escapeHtml(formatRecipeCreatedAt(recipe))}</span>
        </div>
        ${target.detail ? `<p class="recipe-food-subcategory">${escapeHtml(target.detail)}</p>` : ""}
        ${url ? `<a class="recipe-overview-link" href="${escapeHtml(url)}" target="_blank" rel="noreferrer">Rezept öffnen ↗</a>` : recipe.url ? `<p class="recipe-overview-url">${escapeHtml(recipe.url)}</p>` : ""}
        ${recipe.notes ? `<div class="recipe-overview-notes">${escapeHtml(recipe.notes).replace(/\n/g, "<br>")}</div>` : ""}
        <div class="recipe-overview-actions">
          <button class="recipe-edit-button" type="button" data-edit-recipe-id="${escapeHtml(recipe.id)}">Bearbeiten</button>
          <button class="recipe-delete-button" type="button" data-delete-recipe-id="${escapeHtml(recipe.id)}">Löschen</button>
        </div>
      </article>`;
  }

  function recipeTargetCategories(recipe) {
    if (recipe.itemType === "meal") {
      const meal = meals.find((item) => item.id === recipe.itemId);
      return meal ? new Set(mealIngredientFoods(meal).map((food) => food.category)) : new Set();
    }
    const food = foodById.get(recipe.itemId);
    return new Set(food ? [food.category] : []);
  }

  function recipeOverviewEmpty(itemType, hasMatchingFilters) {
    const itemLabel = itemType === "meal" ? "Mahlzeit" : "Lebensmittel";
    return `
      <div class="recipe-overview-empty">
        <span class="button-icon">${icon("recipe")}</span>
        <strong>${hasMatchingFilters ? `Keine passenden Rezepte für ${itemLabel === "Mahlzeit" ? "Mahlzeiten" : "Lebensmittel"} gefunden.` : `Noch keine Rezepte für ${itemLabel === "Mahlzeit" ? "Mahlzeiten" : "Lebensmittel"} gespeichert.`}</strong>
        <p>${hasMatchingFilters ? "Passe die Suche oder Auswahl an und versuche es erneut." : `Klicke bei ${itemLabel === "Mahlzeit" ? "einer Mahlzeit" : "einem Lebensmittel"} auf „meine Rezepte →“, um das erste Rezept hinzuzufügen.`}</p>
      </div>`;
  }

  function renderRecipeOverview() {
    populateRecipeItemFilter();
    const term = state.recipeSearch.trim().toLocaleLowerCase("de");
    const matchingRecipes = cleanRecipes(state.recipes)
      .map((recipe) => ({ recipe, target: recipeTarget(recipe.itemType, recipe.itemId) }))
      .filter(({ recipe, target }) => {
        if (!target) return false;
        const searchable = `${recipe.title} ${recipe.url} ${recipe.notes} ${target.name} ${target.detail}`.toLocaleLowerCase("de");
        return (!term || searchable.includes(term))
          && (!state.recipeCategory || recipeTargetCategories(recipe).has(state.recipeCategory))
          && (!state.recipeItemKey || `${recipe.itemType}:${recipe.itemId}` === state.recipeItemKey);
      })
      .sort((left, right) => {
        const itemOrder = right.recipe.itemId - left.recipe.itemId;
        if (itemOrder) return itemOrder;
        const dateOrder = recipeCreatedTimestamp(right.recipe) - recipeCreatedTimestamp(left.recipe);
        if (dateOrder) return dateOrder;
        return left.target.name.localeCompare(right.target.name, "de", { sensitivity: "base" });
      });

    dom.recipeResultCount.textContent = `${matchingRecipes.length} ${matchingRecipes.length === 1 ? "Rezept" : "Rezepte"}`;
    const mealRecipes = matchingRecipes.filter(({ recipe }) => recipe.itemType === "meal");
    const foodRecipes = matchingRecipes.filter(({ recipe }) => recipe.itemType === "food");
    dom.recipeMealCount.textContent = `${mealRecipes.length} ${mealRecipes.length === 1 ? "Rezept" : "Rezepte"}`;
    dom.recipeFoodCount.textContent = `${foodRecipes.length} ${foodRecipes.length === 1 ? "Rezept" : "Rezepte"}`;
    dom.recipeMealGrid.innerHTML = mealRecipes.length
      ? mealRecipes.map(({ recipe, target }) => renderRecipeOverviewCard(recipe, target)).join("")
      : recipeOverviewEmpty("meal", Boolean(term || state.recipeCategory || state.recipeItemKey));
    dom.recipeFoodGrid.innerHTML = foodRecipes.length
      ? foodRecipes.map(({ recipe, target }) => renderRecipeOverviewCard(recipe, target)).join("")
      : recipeOverviewEmpty("food", Boolean(term || state.recipeCategory || state.recipeItemKey));
  }

  function renderShares() {
    state.selectedSharedMeals = new Set([...state.selectedSharedMeals].filter((id) => state.sharedMeals.has(id)));
    state.selectedSharedFoods = new Set([...state.selectedSharedFoods].filter((id) => state.sharedFoods.has(id)));
    const sharedMeals = meals
      .filter((meal) => state.sharedMeals.has(meal.id))
      .sort((left, right) => right.id - left.id);
    const sharedFoods = foods
      .filter((food) => state.sharedFoods.has(food.id))
      .sort((left, right) => right.id - left.id);

    dom.sharedMealCount.textContent = `${sharedMeals.length} ${sharedMeals.length === 1 ? "Mahlzeit" : "Mahlzeiten"}`;
    dom.sharedFoodCount.textContent = `${sharedFoods.length} ${sharedFoods.length === 1 ? "Lebensmittel" : "Lebensmittel"}`;
    dom.sharedMealGrid.innerHTML = sharedMeals.length
      ? sharedMeals.map((meal) => renderMealCard(meal, {
        isShareView: true,
        isShareSelected: state.selectedSharedMeals.has(meal.id),
        isShareTarget: meal.id === state.sharedTargetMealId,
      })).join("")
      : '<div class="bookmark-empty share-empty"><span class="button-icon">' + icon("share") + '</span><strong>Noch keine Mahlzeit geteilt</strong><p>Klicke bei einer Mahlzeit auf „zum Teilen →“. Der direkte Link wird kopiert und die Mahlzeit erscheint hier.</p></div>';
    dom.sharedFoodGrid.innerHTML = sharedFoods.length
      ? sharedFoods.map((food) => renderFoodCard(food, {
        isShareView: true,
        isShareSelected: state.selectedSharedFoods.has(food.id),
        isShareTarget: food.id === state.sharedTargetFoodId,
      })).join("")
      : '<div class="bookmark-empty share-empty"><span class="button-icon">' + icon("share") + '</span><strong>Noch kein Lebensmittel geteilt</strong><p>Klicke bei einem Lebensmittel auf „zum Teilen →“. Der direkte Link wird kopiert und das Lebensmittel erscheint hier.</p></div>';
    renderSharedSelections();
  }

  function renderSharedSelections() {
    const mealCount = state.selectedSharedMeals.size;
    const foodCount = state.selectedSharedFoods.size;
    dom.sharedMealSelectionCount.textContent = `${mealCount} ausgewählt`;
    dom.sharedFoodSelectionCount.textContent = `${foodCount} ausgewählt`;
    dom.removeSelectedSharedMeals.disabled = mealCount === 0;
    dom.removeSelectedSharedFoods.disabled = foodCount === 0;
  }

  function renderInsights() {
    const categoryCounts = [...foods.reduce((map, food) => map.set(food.category, (map.get(food.category) || 0) + 1), new Map())];
    const scoreCounts = [5, 4, 3, 2, 1].map((score) => ({ score, count: foods.filter((food) => food.score === score).length }));
    const average = foods.reduce((sum, food) => sum + food.score, 0) / foods.length;
    const topRated = scoreCounts[0].count + scoreCounts[1].count;
    const metrics = [
      [foods.length, "Lebensmittel gesamt", ""],
      [meals.length, "Mahlzeiten gesamt", ""],
      [topRated, "Score 4 oder 5", `${Math.round(topRated / foods.length * 100)} % der Auswahl`],
      [categoryCounts.length, "Oberkategorien", "klar gegliedert"],
      [average.toLocaleString("de-DE", { maximumFractionDigits: 1 }), "Ø Sättigungs-Score", "von maximal 5"],
    ];
    document.querySelector("#metricGrid").innerHTML = metrics.map(([value, label, note]) => `
      <article class="metric-card"><span>${label}</span><strong>${value}</strong>${note ? `<small>${note}</small>` : ""}</article>`).join("");
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

  function clearAppUrlHash() {
    const url = new URL(window.location.href);
    if (!url.hash && !url.searchParams.has(sharedMealUrlParam) && !url.searchParams.has(sharedFoodUrlParam)) return;
    url.hash = "";
    url.searchParams.delete(sharedMealUrlParam);
    url.searchParams.delete(sharedFoodUrlParam);
    const cleanUrl = `${url.pathname}${url.search}`;
    if (window.history?.replaceState) {
      window.history.replaceState(null, document.title, cleanUrl);
    } else {
      window.location.hash = "";
    }
  }

  function isValidView(view) {
    return Boolean(view && document.querySelector(`[data-view-panel="${view}"]`));
  }

  function loadSavedView() {
    try {
      const view = sessionStorage.getItem(currentViewStorageKey);
      return isValidView(view) ? view : "";
    } catch {
      return "";
    }
  }

  function saveCurrentView() {
    try {
      sessionStorage.setItem(currentViewStorageKey, state.view);
    } catch {
      // Session storage can be unavailable in strict privacy modes.
    }
  }

  function saveCurrentViewPosition() {
    saveCurrentView();
    try {
      sessionStorage.setItem(currentViewScrollStorageKey, JSON.stringify({
        view: state.view,
        scrollY: Math.max(0, Math.round(window.scrollY || 0)),
      }));
    } catch {
      // Session storage can be unavailable in strict privacy modes.
    }
  }

  function restoreSavedViewPosition(view) {
    let scrollY = 0;
    try {
      const payload = JSON.parse(sessionStorage.getItem(currentViewScrollStorageKey) || "{}");
      if (payload.view !== view || !Number.isFinite(payload.scrollY)) return;
      scrollY = Math.max(0, Math.round(payload.scrollY));
    } catch {
      return;
    }

    const restore = () => window.scrollTo({ top: scrollY, left: 0, behavior: "auto" });
    requestAnimationFrame(() => {
      restore();
      setTimeout(restore, 120);
    });
  }

  function setView(view, options = {}) {
    if (!isValidView(view)) view = "foods";
    state.view = view;
    saveCurrentView();
    document.querySelectorAll("[data-view-panel]").forEach((panel) => panel.classList.toggle("is-active", panel.dataset.viewPanel === view));
    document.querySelectorAll("[data-view]").forEach((button) => button.classList.toggle("is-active", button.dataset.view === view));
    if (options.scroll !== false) window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function viewFromHash() {
    const viewByHash = { "#lebensmittel": "foods", "#mahlzeiten": "meals", "#tagesbaukasten": "meals", "#rezepte": "recipes", "#meine-rezepte": "recipes", "#lesezeichen": "bookmarks", "#auswertung": "insights", "#teilen": "shares", "#so-funktionierts": "help", "#hilfe-oben": "help", "#hilfe-lebensmittel": "help", "#hilfe-mahlzeiten": "help", "#hilfe-rezepte": "help", "#hilfe-lesezeichen": "help", "#hilfe-teilen": "help", "#hilfe-auswertung": "help" };
    return viewByHash[window.location.hash] || "foods";
  }

  function scrollHelpSectionFromHash(hash) {
    if (!hash.startsWith("#hilfe-")) return false;
    const target = document.querySelector(hash);
    if (!target) return false;
    requestAnimationFrame(() => target.scrollIntoView({ behavior: "auto", block: "start" }));
    return true;
  }

  function prepareBrandHomeNavigation() {
    isBrandHomeNavigation = true;
    try {
      sessionStorage.setItem(currentViewStorageKey, "foods");
      sessionStorage.removeItem(currentViewScrollStorageKey);
    } catch {
      // Session storage can be unavailable in strict privacy modes.
    }
  }

  function sharedMealIdFromUrl() {
    const id = Number(new URL(window.location.href).searchParams.get(sharedMealUrlParam));
    return validMealIds.has(id) ? id : null;
  }

  function sharedFoodIdFromUrl() {
    const id = Number(new URL(window.location.href).searchParams.get(sharedFoodUrlParam));
    return validFoodIds.has(id) ? id : null;
  }

  function scrollSharedMealIntoView(mealId) {
    requestAnimationFrame(() => {
      const target = dom.sharedMealGrid.querySelector(`.meal-card[data-id="${mealId}"]`);
      if (!target) return;
      target.scrollIntoView({ behavior: "auto", block: "start" });
      target.focus({ preventScroll: true });
    });
  }

  function scrollSharedFoodIntoView(foodId) {
    requestAnimationFrame(() => {
      const target = dom.sharedFoodGrid.querySelector(`.food-card[data-id="${foodId}"]`);
      if (!target) return;
      target.scrollIntoView({ behavior: "auto", block: "start" });
      target.focus({ preventScroll: true });
    });
  }

  function initializeViewFromUrl() {
    const sharedMealId = sharedMealIdFromUrl();
    if (sharedMealId) {
      state.sharedMeals.add(sharedMealId);
      state.pendingSharedMeals.add(sharedMealId);
      state.sharedTargetMealId = sharedMealId;
      persistSelection();
      renderShares();
      setView("shares", { scroll: false });
      scrollSharedMealIntoView(sharedMealId);
      return;
    }

    const sharedFoodId = sharedFoodIdFromUrl();
    if (sharedFoodId) {
      state.sharedFoods.add(sharedFoodId);
      state.pendingSharedFoods.add(sharedFoodId);
      state.sharedTargetFoodId = sharedFoodId;
      persistSelection();
      renderShares();
      setView("shares", { scroll: false });
      scrollSharedFoodIntoView(sharedFoodId);
      return;
    }

    if (window.location.hash) {
      const initialHash = window.location.hash;
      const view = viewFromHash();
      if (view !== "help") clearAppUrlHash();
      setView(view, { scroll: false });
      if (view === "help" && scrollHelpSectionFromHash(initialHash)) return;
      restoreSavedViewPosition(view);
      return;
    }

    const savedView = loadSavedView();
    if (savedView) {
      setView(savedView, { scroll: false });
      restoreSavedViewPosition(savedView);
    }
  }

  function closeLegalModal() {
    const modal = document.querySelector("#legal-modal");
    if (modal) modal.hidden = true;
  }

  async function openLegalModal(path, title) {
    const modal = document.querySelector("#legal-modal");
    const titleEl = document.querySelector("#legal-modal-title");
    const bodyEl = document.querySelector("#legal-modal-body");
    if (!modal || !titleEl || !bodyEl || !path) return;

    titleEl.textContent = title || "Rechtliches";
    bodyEl.textContent = "Wird geladen...";
    modal.hidden = false;
    clearAppUrlHash();

    try {
      const response = await fetch(path, { cache: "no-cache" });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      bodyEl.textContent = await response.text();
    } catch {
      bodyEl.textContent = "Der rechtliche Text konnte nicht geladen werden.";
    }
  }

  function syncShoppingPanelPlacement() {
    const target = window.innerWidth <= 900 ? document.body : shoppingPanelHome;
    if (dom.shoppingPanel.parentElement !== target) target.append(dom.shoppingPanel);
  }

  function openShopping() {
    syncShoppingPanelPlacement();
    if (window.innerWidth > 900) {
      if (state.view !== "foods") setView("foods");
      requestAnimationFrame(() => dom.shoppingPanel.scrollIntoView({ behavior: "smooth", block: "start" }));
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

  async function copyText(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      try {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        document.body.append(textArea);
        textArea.select();
        const copied = document.execCommand("copy");
        textArea.remove();
        return copied;
      } catch {
        return false;
      }
    }
  }

  async function copyList() {
    if (!state.selected.size) return;
    await copyText(listText());
    showToast("Einkaufsliste kopiert.");
  }

  async function copyMealRecipeSearch(mealId, ingredientNames) {
    const meal = meals.find((item) => item.id === mealId);
    const names = ingredientNames || meal?.ingredients || [];
    if (!meal || !names.length) return;
    const recipeSearchText = `Suche mir Rezepte mit genau diesen Zutaten und füge keine weiteren Zutaten hinzu: ${names.join(", ")}`;
    await copyText(recipeSearchText);
    showToast("Text für die Rezeptsuche wurde kopiert.");
  }

  async function copyMealOffersSearch(mealId, ingredientNames, postalCode) {
    const meal = meals.find((item) => item.id === mealId);
    const names = ingredientNames || meal?.ingredients || [];
    if (!meal || !names.length || !postalCode) return;
    const offersSearchText = `Bitte sage mir wo die folgenden Lebensmittel gekauft werden können und zusätzlich auch wo die folgenden Lebensmittel im Angebot sind. Wenn möglich prüfe bitte auch die digitalen Prospekte der Anbieter. Dann suche mir bitte wo genau diese Zutaten in meiner Nähe gekauft werden können und wenn vorhanden auch wo die folgenden Lebensmittel im Angebot sind. Meine Postleitzahl ist ${postalCode}. Füge bitte keine Zutaten hinzu. Bitte gebe mir nur Anbieter die lokale Läden haben, also keine reinen online Händler. Bitte gebe mir wenn möglich zu den Anbietern auch die URL mit an. Bitte gebe mir das Ergebnis als eine klare, realistische und lokal gültige Liste der günstigsten Preise. Bitte gebe mir ca. alle 30 Sekunden einen Status damit ich weiß ob du noch arbeitest oder ob du fertig bist. Die gesuchten Zutaten sind: ${names.join(", ")}`;
    await copyText(offersSearchText);
    showToast("Text für die Sonderangebotssuche wurde kopiert.");
  }

  function mealShareUrl(mealId) {
    const url = new URL(window.location.href);
    url.search = "";
    url.searchParams.set(sharedMealUrlParam, String(mealId));
    url.hash = "teilen";
    return url.toString();
  }

  function foodShareUrl(foodId) {
    const url = new URL(window.location.href);
    url.search = "";
    url.searchParams.set(sharedFoodUrlParam, String(foodId));
    url.hash = "teilen";
    return url.toString();
  }

  async function shareMeal(mealId) {
    const meal = meals.find((item) => item.id === mealId);
    if (!meal) return;
    state.sharedMeals.add(mealId);
    state.sharedTargetMealId = mealId;
    persistSelection();
    renderShares();
    const shareUrl = mealShareUrl(mealId);
    const copied = await copyText(shareUrl);
    if (window.history?.replaceState) window.history.replaceState(null, document.title, shareUrl);
    showToast(copied ? "Teilen-Link wurde kopiert." : "Teilen-Link steht jetzt in der Adresszeile.");
  }

  async function shareFood(foodId) {
    const food = foodById.get(foodId);
    if (!food) return;
    state.sharedFoods.add(foodId);
    state.sharedTargetFoodId = foodId;
    persistSelection();
    renderShares();
    const shareUrl = foodShareUrl(foodId);
    const copied = await copyText(shareUrl);
    if (window.history?.replaceState) window.history.replaceState(null, document.title, shareUrl);
    showToast(copied ? "Teilen-Link wurde kopiert." : "Teilen-Link steht jetzt in der Adresszeile.");
  }

  function removeSelectedSharedMealEntries() {
    const ids = [...state.selectedSharedMeals].filter((id) => state.sharedMeals.has(id));
    if (!ids.length) return;
    ids.forEach((id) => state.sharedMeals.delete(id));
    state.selectedSharedMeals.clear();
    if (ids.includes(state.sharedTargetMealId)) state.sharedTargetMealId = null;
    if (ids.includes(sharedMealIdFromUrl())) clearAppUrlHash();
    persistSelection();
    renderShares();
    showToast(ids.length === 1 ? "Mahlzeit wurde aus Teilen entfernt." : `${ids.length} Mahlzeiten wurden aus Teilen entfernt.`);
  }

  function removeSelectedSharedFoodEntries() {
    const ids = [...state.selectedSharedFoods].filter((id) => state.sharedFoods.has(id));
    if (!ids.length) return;
    ids.forEach((id) => state.sharedFoods.delete(id));
    state.selectedSharedFoods.clear();
    if (ids.includes(state.sharedTargetFoodId)) state.sharedTargetFoodId = null;
    if (ids.includes(sharedFoodIdFromUrl())) clearAppUrlHash();
    persistSelection();
    renderShares();
    showToast(ids.length === 1 ? "Lebensmittel wurde aus Teilen entfernt." : `${ids.length} Lebensmittel wurden aus Teilen entfernt.`);
  }

  function openConfirm({ title, text, cancel = "Abbrechen", accept = "Liste leeren", tone = "danger", cancelAction = null, action }) {
    state.confirmAction = action;
    state.confirmCancelAction = cancelAction;
    dom.confirmTitle.textContent = title;
    dom.confirmText.textContent = text;
    dom.cancelConfirm.textContent = cancel;
    dom.acceptConfirm.textContent = accept;
    dom.cancelConfirm.className = tone === "cancel-primary" ? "primary-button" : "secondary-button";
    dom.acceptConfirm.className = tone === "cancel-primary"
      ? "secondary-button"
      : tone === "primary" ? "primary-button" : "danger-button solid";
    dom.confirmDialog.showModal();
  }

  function clearSelection() {
    state.selected.clear();
    persistSelection();
    renderFoods();
    renderShoppingList();
    renderMeals();
    renderShares();
    closeShopping();
    showToast("Alle Markierungen wurden gelöscht.");
  }

  function clearBookmarks() {
    state.bookmarkedMeals.clear();
    state.bookmarkedFoods.clear();
    persistSelection();
    renderFoods();
    renderMeals();
    renderBookmarks();
    renderRecipeOverview();
    renderShares();
    showToast("Alle Lesezeichen wurden gelöscht.");
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
      const bookmarkButton = event.target.closest("[data-bookmark-kind][data-bookmark-id]");
      if (bookmarkButton) {
        event.preventDefault();
        toggleBookmark(bookmarkButton.dataset.bookmarkKind, Number(bookmarkButton.dataset.bookmarkId));
        return;
      }

      const brandHome = event.target.closest("[data-brand-home]");
      if (brandHome) {
        event.preventDefault();
        prepareBrandHomeNavigation();
        window.location.assign(brandHome.href);
        return;
      }

      const viewButton = event.target.closest("[data-view]");
      if (viewButton) {
        event.preventDefault();
        clearAppUrlHash();
        setView(viewButton.dataset.view);
      }
      const guideImageButton = event.target.closest("[data-meal-guide-step]");
      if (guideImageButton) openMealGuideImage(Number(guideImageButton.dataset.mealGuideStep));
    });
    [dom.foodGrid, dom.bookmarkedFoodGrid, dom.sharedFoodGrid].forEach((grid) => grid.addEventListener("click", (event) => {
      const card = event.target.closest(".food-card");
      if (!card) return;
      const id = Number(card.dataset.id);
      const shareButton = event.target.closest("[data-share-food-id]");
      if (shareButton) {
        void shareFood(Number(shareButton.dataset.shareFoodId));
        return;
      }
      if (event.target.closest(".select-food")) toggleFood(id);
      if (event.target.closest(".details-button")) openDetails(id);
      if (event.target.closest(".food-meals-button")) applyMealSearchFromFood(foods.find((food) => food.id === id)?.name);
      if (event.target.closest(".food-recipes-button")) openFoodRecipes(id);
    }));
    [dom.mealGrid, dom.bookmarkedMealGrid, dom.sharedMealGrid].forEach((grid) => grid.addEventListener("click", (event) => {
      const personalRecipeButton = event.target.closest("[data-open-meal-recipes]");
      if (personalRecipeButton) {
        openMealRecipes(Number(personalRecipeButton.dataset.openMealRecipes));
        return;
      }
      const offersButton = event.target.closest("[data-offers-meal-id]");
      if (offersButton) {
        requestMealAction(Number(offersButton.dataset.offersMealId), "offers");
        return;
      }
      const shareButton = event.target.closest("[data-share-meal-id]");
      if (shareButton) {
        void shareMeal(Number(shareButton.dataset.shareMealId));
        return;
      }
      const recipeButton = event.target.closest("[data-recipe-meal-id]");
      if (recipeButton) {
        requestMealAction(Number(recipeButton.dataset.recipeMealId), "recipe");
        return;
      }
      const button = event.target.closest("[data-meal-id]");
      if (button) requestMealAction(Number(button.dataset.mealId), "list");
    }));
    dom.shoppingItems.addEventListener("click", (event) => {
      const remove = event.target.closest("[data-remove-id]");
      if (remove) toggleFood(Number(remove.dataset.removeId));
    });
    dom.searchInput.addEventListener("input", () => {
      state.search = dom.searchInput.value;
      state.limit = window.innerWidth < 680 ? 18 : 28;
      renderFoods();
    });
    dom.allFoodsButton.addEventListener("click", openAllFoods);
    dom.mealSearchInput.addEventListener("input", () => {
      state.mealSearch = dom.mealSearchInput.value;
      renderMeals();
    });
    dom.mealTypeFilter.addEventListener("change", () => {
      state.mealType = dom.mealTypeFilter.value;
      renderMeals();
    });
    dom.mealCategoryFilter.addEventListener("change", () => {
      state.mealCategory = dom.mealCategoryFilter.value;
      renderMeals();
    });
    dom.resetMealFilters.addEventListener("click", resetMealFilters);
    dom.recipeSearchInput.addEventListener("input", () => {
      state.recipeSearch = dom.recipeSearchInput.value;
      renderRecipeOverview();
    });
    dom.recipeCategoryFilter.addEventListener("change", () => {
      state.recipeCategory = dom.recipeCategoryFilter.value;
      renderRecipeOverview();
    });
    dom.recipeItemFilter.addEventListener("change", () => {
      state.recipeItemKey = dom.recipeItemFilter.value;
      renderRecipeOverview();
    });
    dom.resetRecipeFilters.addEventListener("click", resetRecipeFilters);
    [dom.recipeMealGrid, dom.recipeFoodGrid].forEach((grid) => grid.addEventListener("click", (event) => {
      const itemButton = event.target.closest("[data-open-recipe-item-type][data-open-recipe-item-id]");
      if (itemButton) {
        openRecipes(itemButton.dataset.openRecipeItemType, Number(itemButton.dataset.openRecipeItemId));
        return;
      }
      const editButton = event.target.closest("[data-edit-recipe-id]");
      if (editButton) {
        const recipe = state.recipes.find((item) => item.id === editButton.dataset.editRecipeId);
        if (recipe) openRecipes(recipe.itemType, recipe.itemId, recipe.id);
        return;
      }
      const deleteButton = event.target.closest("[data-delete-recipe-id]");
      if (deleteButton) requestRecipeDelete(deleteButton.dataset.deleteRecipeId, { reopenManager: false });
    }));
    [[dom.categoryFilter, "category"], [dom.scoreFilter, "score"], [dom.priorityFilter, "priority"]].forEach(([element, key]) => {
      element.addEventListener("change", () => {
        state[key] = element.value;
        state.limit = window.innerWidth < 680 ? 18 : 28;
        renderFoods();
      });
    });
    window.addEventListener("beforeunload", () => {
      if (!isBrandHomeNavigation) saveCurrentViewPosition();
    });
    dom.resetFilters.addEventListener("click", resetFoodFilters);
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
    dom.syncButton.addEventListener("click", (event) => {
      event.stopPropagation();
      toggleSyncMenu();
    });
    dom.syncMenu?.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    dom.syncMenuPrimary?.addEventListener("click", runOneDrivePrimaryAction);
    dom.syncMenuRenew?.addEventListener("click", renewOneDriveLogin);
    dom.syncMenuLogout?.addEventListener("click", logoutFromOneDrive);
    dom.syncSecondary?.addEventListener("click", runOneDrivePrimaryAction);
    dom.syncRenew?.addEventListener("click", renewOneDriveLogin);
    dom.syncLogout?.addEventListener("click", logoutFromOneDrive);
    document.addEventListener("click", closeSyncMenu);
    [dom.closeShopping, dom.scrim].forEach((element) => element.addEventListener("click", closeShopping));
    document.querySelector("#downloadList").addEventListener("click", downloadList);
    document.querySelector("#copyList").addEventListener("click", copyList);
    document.querySelector("#clearList").addEventListener("click", () => openConfirm({
      title: "Einkaufsliste leeren?",
      text: `Alle ${state.selected.size} Markierungen werden entfernt.`,
      action: clearSelection,
    }));
    dom.clearAllBookmarks.addEventListener("click", () => {
      const count = state.bookmarkedMeals.size + state.bookmarkedFoods.size;
      if (!count) return;
      openConfirm({
        title: "Alle Lesezeichen löschen?",
        text: `Alle ${count} Lesezeichen für Mahlzeiten und Lebensmittel werden entfernt.`,
        accept: "Alle löschen",
        action: clearBookmarks,
      });
    });
    dom.sharedMealGrid.addEventListener("change", (event) => {
      const checkbox = event.target.closest("[data-shared-meal-select]");
      if (!checkbox) return;
      const mealId = Number(checkbox.value);
      if (!state.sharedMeals.has(mealId)) return;
      if (checkbox.checked) state.selectedSharedMeals.add(mealId);
      else state.selectedSharedMeals.delete(mealId);
      checkbox.nextElementSibling.textContent = checkbox.checked ? "Ausgewählt" : "Zum Entfernen auswählen";
      renderSharedSelections();
    });
    dom.sharedFoodGrid.addEventListener("change", (event) => {
      const checkbox = event.target.closest("[data-shared-food-select]");
      if (!checkbox) return;
      const foodId = Number(checkbox.value);
      if (!state.sharedFoods.has(foodId)) return;
      if (checkbox.checked) state.selectedSharedFoods.add(foodId);
      else state.selectedSharedFoods.delete(foodId);
      checkbox.nextElementSibling.textContent = checkbox.checked ? "Ausgewählt" : "Zum Entfernen auswählen";
      renderSharedSelections();
    });
    dom.removeSelectedSharedMeals.addEventListener("click", () => {
      const count = state.selectedSharedMeals.size;
      if (!count) return;
      openConfirm({
        title: count === 1 ? "Mahlzeit aus Teilen entfernen?" : `${count} Mahlzeiten aus Teilen entfernen?`,
        text: count === 1
          ? "Die ausgewählte Mahlzeit wird aus deinem Bereich Teilen entfernt und bei OneDrive-Anmeldung auf deinen Geräten synchronisiert. Bereits versendete Links bleiben erreichbar."
          : "Die ausgewählten Mahlzeiten werden aus deinem Bereich Teilen entfernt und bei OneDrive-Anmeldung auf deinen Geräten synchronisiert. Bereits versendete Links bleiben erreichbar.",
        accept: count === 1 ? "Mahlzeit entfernen" : "Mahlzeiten entfernen",
        action: removeSelectedSharedMealEntries,
      });
    });
    dom.removeSelectedSharedFoods.addEventListener("click", () => {
      const count = state.selectedSharedFoods.size;
      if (!count) return;
      openConfirm({
        title: count === 1 ? "Lebensmittel aus Teilen entfernen?" : `${count} Lebensmittel aus Teilen entfernen?`,
        text: count === 1
          ? "Das ausgewählte Lebensmittel wird aus deinem Bereich Teilen entfernt und bei OneDrive-Anmeldung auf deinen Geräten synchronisiert. Bereits versendete Links bleiben erreichbar."
          : "Die ausgewählten Lebensmittel werden aus deinem Bereich Teilen entfernt und bei OneDrive-Anmeldung auf deinen Geräten synchronisiert. Bereits versendete Links bleiben erreichbar.",
        accept: "Lebensmittel entfernen",
        action: removeSelectedSharedFoodEntries,
      });
    });
    document.querySelector(".dialog-close").addEventListener("click", () => dom.detailDialog.close());
    dom.detailDialog.addEventListener("click", (event) => {
      const editRecipeButton = event.target.closest("[data-edit-recipe-id]");
      if (editRecipeButton) {
        const recipe = state.recipes.find((item) => item.id === editRecipeButton.dataset.editRecipeId);
        if (recipe) openRecipes(recipe.itemType, recipe.itemId, recipe.id);
        return;
      }
      const deleteRecipeButton = event.target.closest("[data-delete-recipe-id]");
      if (deleteRecipeButton) {
        requestRecipeDelete(deleteRecipeButton.dataset.deleteRecipeId);
        return;
      }
      const cancelRecipeEdit = event.target.closest(".recipe-cancel-edit");
      if (cancelRecipeEdit) {
        const form = cancelRecipeEdit.closest("[data-recipe-form]");
        if (form) openRecipes(form.dataset.recipeItemType, Number(form.dataset.recipeItemId));
        return;
      }
      const foodSearchButton = event.target.closest("[data-food-search]");
      if (foodSearchButton) {
        applyFoodSearchFromList(foodSearchButton.dataset.foodSearch);
        return;
      }
      if (event.target === dom.detailDialog) dom.detailDialog.close();
    });
    dom.detailDialog.addEventListener("input", (event) => {
      const form = event.target.closest("[data-recipe-form]");
      if (!form) return;
      const hasContent = ["recipe-title", "recipe-url", "recipe-notes"]
        .some((name) => String(form.elements[name].value || "").trim());
      form.querySelector(".recipe-save-button").disabled = !hasContent;
    });
    dom.detailDialog.addEventListener("change", (event) => {
      const form = event.target.closest("[data-meal-variation-form]");
      if (!form || !event.target.matches('input[name="meal-food"]')) return;
      const count = form.querySelectorAll('input[name="meal-food"]:checked').length;
      form.querySelector(".meal-selection-count").textContent = `${count} Lebensmittel ausgewählt`;
      form.querySelector(".meal-variation-action").disabled = count === 0;
    });
    dom.detailDialog.addEventListener("submit", (event) => {
      const recipeForm = event.target.closest("[data-recipe-form]");
      if (recipeForm) {
        event.preventDefault();
        saveRecipe(recipeForm);
        return;
      }
      const postalForm = event.target.closest("[data-meal-offers-postal-form]");
      if (postalForm) {
        event.preventDefault();
        const postalCode = postalForm.elements["postal-code"].value.trim();
        if (!/^[0-9]{5}$/.test(postalCode)) return;
        const names = [...postalForm.querySelectorAll('input[name="meal-offers-food"]')].map((input) => input.value);
        if (!names.length) return;
        dom.detailDialog.close();
        void copyMealOffersSearch(Number(postalForm.dataset.mealId), names, postalCode);
        return;
      }
      const form = event.target.closest("[data-meal-variation-form]");
      if (!form) return;
      event.preventDefault();
      const names = [...form.querySelectorAll('input[name="meal-food"]:checked')].map((input) => input.value);
      if (!names.length) return;
      dom.detailDialog.close();
      runMealAction(Number(form.dataset.mealId), form.dataset.mealAction, names);
    });
    dom.cancelConfirm.addEventListener("click", () => {
      const cancelAction = state.confirmCancelAction;
      dom.confirmDialog.close();
      state.confirmAction = null;
      state.confirmCancelAction = null;
      if (cancelAction) cancelAction();
    });
    dom.acceptConfirm.addEventListener("click", () => {
      const action = state.confirmAction;
      dom.confirmDialog.close();
      state.confirmAction = null;
      state.confirmCancelAction = null;
      if (action) action();
    });
    document.querySelector("#legal-modal-close")?.addEventListener("click", closeLegalModal);
    document.querySelector("#legal-modal")?.addEventListener("click", (event) => {
      if (event.target === document.querySelector("#legal-modal")) closeLegalModal();
    });
    document.querySelectorAll(".app-footer a[data-legal-path]").forEach((link) => {
      link.addEventListener("click", (event) => {
        event.preventDefault();
        void openLegalModal(link.dataset.legalPath, link.dataset.legalTitle);
      });
    });
    document.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        if (state.view === "meals") dom.mealSearchInput.focus();
        else if (state.view === "recipes") dom.recipeSearchInput.focus();
        else {
          setView("foods");
          dom.searchInput.focus();
        }
      }
      if (event.key === "Escape") {
        closeShopping();
        closeSyncMenu();
      }
    });
    window.addEventListener("resize", () => {
      const wasMobile = dom.shoppingPanel.parentElement === document.body;
      syncShoppingPanelPlacement();
      if (wasMobile && window.innerWidth > 900) closeShopping();
    });
  }

  async function initialize() {
    if (await ensureLatestAppVersion()) return;
    clearAuthReloadParam();
    if (hasOneDriveManualLogout()) clearOneDriveRedirectResponseUrl();
    const hasAuthRedirect = !hasOneDriveManualLogout() && hasOneDriveRedirectResponse();
    if (hasAuthRedirect) {
      setSyncStatus("loading", "Microsoft-Anmeldung", "Microsoft-Rückkehr wird verarbeitet.");
      await initializeOneDrive();
    }

    syncShoppingPanelPlacement();
    document.querySelectorAll("[data-icon]").forEach((slot) => { slot.innerHTML = icon(slot.dataset.icon); });
    populateFilters();
    renderFoods();
    renderShoppingList();
    renderSyncStatus();
    renderMeals();
    renderBookmarks();
    renderRecipeOverview();
    renderShares();
    renderInsights();
    bindEvents();

    if (!hasAuthRedirect) {
      initializeViewFromUrl();
      await initializeOneDrive();
    }
  }

  void initialize();
})();
