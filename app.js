(() => {
  "use strict";

  const { foods, meals, sources, foodNames = [] } = window.APP_DATA;
  const appVersion = "meal-reset-link-20260711-1";
  const appVersionFile = "app-version.json";
  const appRefreshParam = "appRefresh";
  const appRefreshSessionKey = "lebensmitteleinkauf:app-refresh-version:v1";
  const currentViewStorageKey = "lebensmitteleinkauf:current-view:v1";
  const currentViewScrollStorageKey = "lebensmitteleinkauf:current-view-scroll:v1";
  const storageKey = "lebensmitteleinkauf:selected:v1";
  const storageMetaKey = "lebensmitteleinkauf:selected:meta:v1";
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
  const localSnapshot = loadSelectionData();
  const foodByName = new Map(foods.map((food) => [normalizeFoodName(food.name), food]));
  const mealGuideImages = {
    1: { src: "assets/meal-guide/step-1.png", alt: "Bildanleitung zu Schritt 1: Eine Mahlzeit auswählen" },
    2: { src: "assets/meal-guide/step-2.png", alt: "Bildanleitung zu Schritt 2: Text für die Rezeptsuche kopieren" },
    3: { src: "assets/meal-guide/step-3.png", alt: "Bildanleitung zu Schritt 3: Den kopierten Text in eine KI einfügen" },
    4: { src: "assets/meal-guide/step-4.png", alt: "Bildanleitung zu Schritt 4: Zutaten auf die Einkaufsliste setzen" },
  };

  const iconPaths = {
    basket: '<path d="M7 10 10 4M17 10l-3-6M4 10h16l-1.4 9H5.4L4 10Z"/><path d="M9 13v3M15 13v3"/>',
    download: '<path d="M12 3v12M7.5 10.5 12 15l4.5-4.5"/><path d="M4 18v2h16v-2"/>',
    search: '<circle cx="11" cy="11" r="6.5"/><path d="m16 16 4 4"/>',
    leaf: '<path d="M19 4C11 4 6 8 6 15c0 3 2 5 5 5 7 0 9-8 8-16Z"/><path d="M5 21c2-5 6-9 11-12"/>',
    meal: '<path d="M7 3v8M4.5 3v5c0 2 1 3 2.5 3s2.5-1 2.5-3V3M7 11v10"/><path d="M16 3c2 2 3 5 3 8v2h-5V9c0-3 1-5 2-6Zm0 10v8"/>',
    chart: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
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
    category: "",
    score: "",
    priority: "",
    limit: window.innerWidth < 680 ? 18 : 28,
    selected: new Set(localSnapshot.selected),
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
      message: "Deine Liste wird lokal auf diesem Gerät gespeichert.",
      lastRemoteUpdatedAt: "",
      lastRemoteEtag: "",
      hasRemoteData: false,
      conflictData: null,
    },
    confirmAction: null,
  };

  const dom = {
    foodGrid: document.querySelector("#foodGrid"),
    mealGrid: document.querySelector("#mealGrid"),
    resultCount: document.querySelector("#resultCount"),
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

  function cleanSelectedIds(value) {
    return [...new Set((Array.isArray(value) ? value : []).map(Number).filter((id) => validFoodIds.has(id)))];
  }

  function nowIso() {
    return new Date().toISOString();
  }

  function loadSelectionData() {
    try {
      const selected = cleanSelectedIds(JSON.parse(localStorage.getItem(storageKey) || "[]"));
      const meta = JSON.parse(localStorage.getItem(storageMetaKey) || "{}");
      const updatedAt = typeof meta.updatedAt === "string" ? meta.updatedAt : selected.length ? nowIso() : "";
      return { selected, updatedAt };
    } catch {
      return { selected: [], updatedAt: "" };
    }
  }

  function saveSelectionLocally(updatedAt = nowIso()) {
    const selected = cleanSelectedIds([...state.selected]);
    localStorage.setItem(storageKey, JSON.stringify(selected));
    localStorage.setItem(storageMetaKey, JSON.stringify({ updatedAt }));
    state.localUpdatedAt = updatedAt;
    return { selected, updatedAt };
  }

  function persistSelection() {
    saveSelectionLocally();
    queueOneDriveSave();
  }

  function selectionPayload(updatedAt = state.localUpdatedAt || nowIso()) {
    return {
      app: "lebensmitteleinkauf",
      version: 1,
      updatedAt,
      selectedIds: cleanSelectedIds([...state.selected]),
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
    };
  }

  function sameSelection(left, right) {
    return cleanSelectedIds(left).join(",") === cleanSelectedIds(right).join(",");
  }

  function remoteIsNewer(remoteUpdatedAt, knownUpdatedAt) {
    if (!remoteUpdatedAt || !knownUpdatedAt) return false;
    return Date.parse(remoteUpdatedAt) > Date.parse(knownUpdatedAt);
  }

  function setSyncStatus(status, title, message) {
    if (hasOneDriveManualLogout() && status !== "local" && title !== "OneDrive-Abmeldung") {
      state.sync.status = "local";
      state.sync.title = "Nicht angemeldet";
      state.sync.message = "Deine Liste wird lokal auf diesem Gerät gespeichert.";
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
        return `Deine Liste wird lokal auf diesem Gerät gespeichert.\n${helpText}`;
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
    return `Microsoft-Anmeldung fehlgeschlagen${suffix}. Deine Liste bleibt lokal gespeichert.`;
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
    saveSelectionLocally(remoteData.updatedAt || nowIso());
    state.sync.lastRemoteUpdatedAt = remoteData.updatedAt || state.localUpdatedAt;
    state.sync.lastRemoteEtag = remoteEtag;
    state.sync.hasRemoteData = true;
    state.sync.conflictData = null;
    renderFoods();
    renderShoppingList();
    renderMeals();
  }

  function completeRemoteSave(payload, metadata) {
    state.sync.lastRemoteUpdatedAt = payload.updatedAt;
    state.sync.lastRemoteEtag = metadata?.eTag || state.sync.lastRemoteEtag;
    state.sync.hasRemoteData = true;
    state.sync.conflictData = null;
    state.sync.needsInteractiveToken = false;
    setSyncStatus("synced", "Mit OneDrive synchronisiert", "Deine Einkaufsliste ist im OneDrive-App-Ordner gespeichert.");
  }

  function handleOneDriveError(error, fallbackTitle = "OneDrive nicht verfügbar") {
    if (error?.message === "redirect-started") return;
    if (error?.message === "not-signed-in") {
      state.sync.needsInteractiveToken = false;
      setSyncStatus("local", "Nicht angemeldet", "Deine Liste wird lokal auf diesem Gerät gespeichert.");
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
    if (!localData.selected.length) return true;
    if (remoteIsNewer(remoteData.updatedAt, localData.updatedAt)) return true;
    return Boolean(remoteData.updatedAt && !localData.updatedAt);
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
        && !sameSelection(latest.data.selectedIds, state.selected)
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
    setSyncStatus("loading", "Prüfe OneDrive", "Aktuelle Einkaufsliste wird aus OneDrive abgeglichen.");
    try {
      const remote = await loadRemoteSelection();
      const local = loadSelectionData();

      if (remote.exists && remoteShouldReplaceLocal(remote.data, local)) {
        applyRemoteSelection(remote.data, remote.etag);
        state.sync.needsInteractiveToken = false;
        setSyncStatus("synced", "Mit OneDrive synchronisiert", "Neuere OneDrive-Liste wurde übernommen.");
        showToast("Neuere OneDrive-Liste wurde übernommen.");
        return;
      }

      if (remote.exists && sameSelection(remote.data.selectedIds, state.selected)) {
        state.sync.lastRemoteUpdatedAt = remote.data.updatedAt || state.sync.lastRemoteUpdatedAt;
        state.sync.lastRemoteEtag = remote.etag || state.sync.lastRemoteEtag;
        state.sync.needsInteractiveToken = false;
        setSyncStatus("synced", "Mit OneDrive synchronisiert", "Deine Einkaufsliste ist aktuell.");
        return;
      }

      if (!remote.exists || remoteIsNewer(local.updatedAt, remote.data?.updatedAt)) {
        const payload = selectionPayload(nowIso());
        saveSelectionLocally(payload.updatedAt);
        const metadata = await uploadRemoteSelection(payload);
        completeRemoteSave(payload, metadata);
        showToast("Einkaufsliste wurde nach OneDrive gespeichert.");
        return;
      }

      if (remote.exists) {
        applyRemoteSelection(remote.data, remote.etag);
        state.sync.needsInteractiveToken = false;
        setSyncStatus("synced", "Mit OneDrive synchronisiert", "OneDrive-Liste wurde übernommen.");
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
      setSyncStatus("local", "Nicht angemeldet", "Deine Liste wird lokal auf diesem Gerät gespeichert.");
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
    setSyncStatus("loading", "Prüfe OneDrive", "Deine Einkaufsliste wird geladen.");
    try {
      const remote = forceRemote && state.sync.conflictData ? state.sync.conflictData : await loadRemoteSelection();
      const local = loadSelectionData();

      if (!remote.exists) {
        if (local.selected.length) {
          const payload = selectionPayload(local.updatedAt || nowIso());
          const metadata = await uploadRemoteSelection(payload);
          completeRemoteSave(payload, metadata);
          showToast("Lokale Einkaufsliste wurde nach OneDrive übernommen.");
        } else {
          state.sync.hasRemoteData = false;
          state.sync.needsInteractiveToken = false;
          setSyncStatus("synced", "Mit OneDrive verbunden", "Noch keine Einkaufsliste im OneDrive-App-Ordner.");
        }
        return;
      }

      if (forceRemote || !local.selected.length || remoteIsNewer(remote.data.updatedAt, local.updatedAt) || sameSelection(remote.data.selectedIds, local.selected)) {
        applyRemoteSelection(remote.data, remote.etag);
        state.sync.needsInteractiveToken = false;
        setSyncStatus("synced", "Mit OneDrive synchronisiert", "Deine Einkaufsliste wurde aus OneDrive geladen.");
        return;
      }

      state.sync.lastRemoteUpdatedAt = remote.data.updatedAt || "";
      state.sync.lastRemoteEtag = remote.etag || "";
      state.sync.hasRemoteData = true;
      state.sync.conflictData = remote;
      state.sync.needsInteractiveToken = false;
      setSyncStatus("conflict", "Unterschiedliche Listen", "Lokale und OneDrive-Liste unterscheiden sich. Es wurde nichts überschrieben.");
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
      setSyncStatus("local", "Nicht angemeldet", "Deine Liste wird lokal auf diesem Gerät gespeichert.");
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
      setSyncStatus("local", "Nicht angemeldet", "Deine Liste wird lokal auf diesem Gerät gespeichert.");
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
        setSyncStatus("local", "Nicht angemeldet", "Deine Liste wird lokal auf diesem Gerät gespeichert.");
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
        setSyncStatus("local", "Nicht angemeldet", "Deine Liste wird lokal auf diesem Gerät gespeichert. Melde dich an, um OneDrive zu nutzen.");
      }
    } catch (error) {
      if (hasOneDriveManualLogout()) {
        clearOneDriveRedirectResponseUrl();
        clearOneDriveMsalCache();
        clearOneDriveMsalCookies();
        state.sync.account = null;
        setSyncStatus("local", "Nicht angemeldet", "Deine Liste wird lokal auf diesem Gerät gespeichert.");
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
  }

  function mealIngredientFoods(meal) {
    return (meal.ingredients || []).map((ingredient) => foodByName.get(normalizeFoodName(ingredient))).filter(Boolean);
  }

  function addMealIngredients(mealId) {
    const meal = meals.find((item) => item.id === mealId);
    if (!meal) return;
    const ingredients = mealIngredientFoods(meal);
    const previousSize = state.selected.size;
    ingredients.forEach((food) => state.selected.add(food.id));
    const addedCount = state.selected.size - previousSize;
    persistSelection();
    renderFoods();
    renderShoppingList();
    renderMeals();
    showToast(addedCount ? `${addedCount} Zutaten wurden auf die Liste gesetzt.` : "Alle Zutaten sind bereits auf der Liste.");
  }

  function openDetails(id) {
    const food = foods.find((item) => item.id === id);
    if (!food) return;
    const links = String(food.sources).split(";").map((url) => url.trim()).filter(Boolean);
    dom.detailDialog.classList.remove("is-image-dialog");
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

  function openAllFoods() {
    const names = Array.isArray(foodNames) && foodNames.length
      ? foodNames
      : [...new Set(foods.map((food) => food.name).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, "de", { sensitivity: "base" }));
    dom.detailDialog.classList.remove("is-image-dialog");
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

  function renderMeals() {
    const term = state.mealSearch.trim().toLocaleLowerCase("de");
    const visibleMeals = meals
      .map((meal, index) => ({ meal, index }))
      .filter(({ meal }) => {
        const searchable = `${meal.situation} ${meal.satiety} ${(meal.ingredients || []).join(" ")} ${meal.reason} ${meal.variants} ${formatMealDate(meal.date)}`.toLocaleLowerCase("de");
        const mealType = normalizeMealType(meal.mealType);
        const ingredientFoods = mealIngredientFoods(meal);
        return (!term || searchable.includes(term))
          && (!state.mealType || mealType === state.mealType)
          && (!state.mealCategory || ingredientFoods.some((food) => food.category === state.mealCategory));
      });

    dom.mealGrid.innerHTML = visibleMeals.length ? visibleMeals.map(({ meal, index }) => {
      const mealDate = formatMealDate(meal.date);
      const mealNumber = meals.length - index;
      const ingredientFoods = mealIngredientFoods(meal);
      const allIngredientsSelected = ingredientFoods.length > 0 && ingredientFoods.every((food) => state.selected.has(food.id));
      return `
      <article class="meal-card">
        <div class="meal-meta">
          ${mealDate ? `<time class="meal-date" datetime="${escapeHtml(meal.date)}">${escapeHtml(mealDate)}</time>` : ""}
          <span class="meal-number">${String(mealNumber).padStart(2, "0")}</span>
        </div>
        <p class="eyebrow">${escapeHtml(meal.satiety)} sättigend</p>
        <h2>${escapeHtml(meal.situation)}</h2>
        <p class="meal-ingredients"><strong>Zutaten:</strong> ${(meal.ingredients || []).map(escapeHtml).join(", ")}</p>
        <div class="meal-details">
          <div><span>Warum sinnvoll</span><p>${escapeHtml(meal.reason)}</p></div>
          <div><span>Varianten</span><p>${escapeHtml(meal.variants)}</p></div>
        </div>
        <div class="meal-card-footer">
          <button class="meal-recipe-button" type="button" data-recipe-meal-id="${meal.id}" aria-label="Rezeptsuchtext für ${escapeHtml(meal.situation)} kopieren">für Rezeptsuche →</button>
          <button class="meal-list-button${allIngredientsSelected ? " is-added" : ""}" type="button" data-meal-id="${meal.id}" aria-pressed="${allIngredientsSelected}">${allIngredientsSelected ? "Auf der Liste ✓" : "Auf die Liste →"}</button>
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
      [foods.length, "Lebensmittel gesamt", ""],
      [meals.length, "Rezepte gesamt", ""],
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
    if (!window.location.hash) return;
    const cleanUrl = `${window.location.pathname}${window.location.search}`;
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
    const viewByHash = { "#lebensmittel": "foods", "#mahlzeiten": "meals", "#tagesbaukasten": "meals", "#auswertung": "insights" };
    return viewByHash[window.location.hash] || "foods";
  }

  function initializeViewFromUrl() {
    if (window.location.hash) {
      const view = viewFromHash();
      clearAppUrlHash();
      setView(view, { scroll: false });
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
    } catch {
      const textArea = document.createElement("textarea");
      textArea.value = text;
      document.body.append(textArea);
      textArea.select();
      document.execCommand("copy");
      textArea.remove();
    }
  }

  async function copyList() {
    if (!state.selected.size) return;
    await copyText(listText());
    showToast("Einkaufsliste kopiert.");
  }

  async function copyMealRecipeSearch(mealId) {
    const meal = meals.find((item) => item.id === mealId);
    if (!meal || !meal.ingredients?.length) return;
    const recipeSearchText = `Suche mir Rezepte mit genau diesen Zutaten und füge keine weiteren Zutaten hinzu: ${meal.ingredients.join(", ")}`;
    await copyText(recipeSearchText);
    showToast("Text für die Rezeptsuche wurde kopiert.");
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
    renderMeals();
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
      const brandReload = event.target.closest("[data-brand-reload]");
      if (brandReload) {
        event.preventDefault();
        saveCurrentViewPosition();
        window.location.reload();
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
    dom.foodGrid.addEventListener("click", (event) => {
      const card = event.target.closest(".food-card");
      if (!card) return;
      const id = Number(card.dataset.id);
      if (event.target.closest(".select-food")) toggleFood(id);
      if (event.target.closest(".details-button")) openDetails(id);
      if (event.target.closest(".food-meals-button")) applyMealSearchFromFood(foods.find((food) => food.id === id)?.name);
    });
    dom.mealGrid.addEventListener("click", (event) => {
      const recipeButton = event.target.closest("[data-recipe-meal-id]");
      if (recipeButton) {
        copyMealRecipeSearch(Number(recipeButton.dataset.recipeMealId));
        return;
      }
      const button = event.target.closest("[data-meal-id]");
      if (button) addMealIngredients(Number(button.dataset.mealId));
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
    [[dom.categoryFilter, "category"], [dom.scoreFilter, "score"], [dom.priorityFilter, "priority"]].forEach(([element, key]) => {
      element.addEventListener("change", () => {
        state[key] = element.value;
        state.limit = window.innerWidth < 680 ? 18 : 28;
        renderFoods();
      });
    });
    window.addEventListener("beforeunload", saveCurrentViewPosition);
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
    document.querySelector(".dialog-close").addEventListener("click", () => dom.detailDialog.close());
    dom.detailDialog.addEventListener("click", (event) => {
      const foodSearchButton = event.target.closest("[data-food-search]");
      if (foodSearchButton) {
        applyFoodSearchFromList(foodSearchButton.dataset.foodSearch);
        return;
      }
      if (event.target === dom.detailDialog) dom.detailDialog.close();
    });
    dom.cancelConfirm.addEventListener("click", () => dom.confirmDialog.close());
    dom.acceptConfirm.addEventListener("click", () => {
      dom.confirmDialog.close();
      if (state.confirmAction) state.confirmAction();
      state.confirmAction = null;
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
    renderInsights();
    bindEvents();

    if (!hasAuthRedirect) {
      initializeViewFromUrl();
      await initializeOneDrive();
    }
  }

  void initialize();
})();
