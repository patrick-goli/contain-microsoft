// Param values from https://developer.mozilla.org/Add-ons/WebExtensions/API/contextualIdentities/create
const MICROSOFT_CONTAINER_NAME = "Microsoft";
const MICROSOFT_CONTAINER_COLOR = "blue";
const MICROSOFT_CONTAINER_ICON = "briefcase";

let MICROSOFT_DOMAINS = ["microsoft.com", "microsoft.org", "microsoft.net"];
const MICROSOFT_SERVICES = ["bing.com", "bing.net", "bingtoolbar.com", "outlook.com", "skype.com", "hotmail.com", "live.com",
  "onedrive.com", "office.com", "xbox.com", "visualstudio.com", "microsoftvisualstudio.com"];
///
MICROSOFT_DOMAINS = MICROSOFT_DOMAINS.concat(MICROSOFT_SERVICES);

const MAC_ADDON_ID = "@testpilot-containers";

let macAddonEnabled = false;
let microsoftCookieStoreId = null;
let microsoftCookiesCleared = false;

const canceledRequests = {};
const microsoftHostREs = [];

async function isMACAddonEnabled() {
  try {
    const macAddonInfo = await browser.management.get(MAC_ADDON_ID);
    if (macAddonInfo.enabled) {
      return true;
    }
  } catch (e) {
    return false;
  }
  return false;
}

async function setupMACAddonManagementListeners() {
  browser.management.onInstalled.addListener(info => {
    if (info.id === MAC_ADDON_ID) {
      macAddonEnabled = true;
    }
  });
  browser.management.onUninstalled.addListener(info => {
    if (info.id === MAC_ADDON_ID) {
      macAddonEnabled = false;
    }
  })
  browser.management.onEnabled.addListener(info => {
    if (info.id === MAC_ADDON_ID) {
      macAddonEnabled = true;
    }
  })
  browser.management.onDisabled.addListener(info => {
    if (info.id === MAC_ADDON_ID) {
      macAddonEnabled = false;
    }
  })
}

async function getMACAssignment(url) {
  try {
    const assignment = await browser.runtime.sendMessage(MAC_ADDON_ID, {
      method: "getAssignment",
      url
    });
    return assignment;
  } catch (e) {
    return false;
  }
}

function cancelRequest(tab, options) {
  // we decided to cancel the request at this point, register canceled request
  canceledRequests[tab.id] = {
    requestIds: {
      [options.requestId]: true
    },
    urls: {
      [options.url]: true
    }
  };

  // since webRequest onCompleted and onErrorOccurred are not 100% reliable
  // we register a timer here to cleanup canceled requests, just to make sure we don't
  // end up in a situation where certain urls in a tab.id stay canceled
  setTimeout(() => {
    if (canceledRequests[tab.id]) {
      delete canceledRequests[tab.id];
    }
  }, 2000);
}

function shouldCancelEarly(tab, options) {
  // we decided to cancel the request at this point
  if (!canceledRequests[tab.id]) {
    cancelRequest(tab, options);
  } else {
    let cancelEarly = false;
    if (canceledRequests[tab.id].requestIds[options.requestId] ||
      canceledRequests[tab.id].urls[options.url]) {
      // same requestId or url from the same tab
      // this is a redirect that we have to cancel early to prevent opening two tabs
      cancelEarly = true;
    }
    // register this requestId and url as canceled too
    canceledRequests[tab.id].requestIds[options.requestId] = true;
    canceledRequests[tab.id].urls[options.url] = true;
    if (cancelEarly) {
      return true;
    }
  }
  return false;
}

function generateMicrosoftHostREs() {
  for (let microsoftDomain of MICROSOFT_DOMAINS) {
    microsoftHostREs.push(new RegExp(`^(.*\\.)?${microsoftDomain}$`));
  }
}

async function clearMicrosoftCookies() {
  // Clear all microsoft cookies
  const containers = await browser.contextualIdentities.query({});
  containers.push({
    cookieStoreId: 'firefox-default'
  });
  containers.map(container => {
    const storeId = container.cookieStoreId;
    if (storeId === microsoftCookieStoreId) {
      // Don't clear cookies in the Microsoft Container
      return;
    }

    MICROSOFT_DOMAINS.map(async microsoftDomain => {
      const microsoftCookieUrl = `https://${microsoftDomain}/`;

      const cookies = await browser.cookies.getAll({
        domain: microsoftDomain,
        storeId
      });

      cookies.map(cookie => {
        browser.cookies.remove({
          name: cookie.name,
          url: microsoftCookieUrl,
          storeId
        });
      });
    });
  });
}

async function setupContainer() {
  // Use existing Microsoft container, or create one
  const contexts = await browser.contextualIdentities.query({ name: MICROSOFT_CONTAINER_NAME })
  if (contexts.length > 0) {
    microsoftCookieStoreId = contexts[0].cookieStoreId;
  } else {
    const context = await browser.contextualIdentities.create({
      name: MICROSOFT_CONTAINER_NAME,
      color: MICROSOFT_CONTAINER_COLOR,
      icon: MICROSOFT_CONTAINER_ICON
    })
    microsoftCookieStoreId = context.cookieStoreId;
  }
}

async function containMicrosoft(options) {
  // Listen to requests and open Microsoft into its Container,
  // open other sites into the default tab context
  const requestUrl = new URL(options.url);

  let isMicrosoft = false;
  for (let microsoftHostRE of microsoftHostREs) {
    if (microsoftHostRE.test(requestUrl.host)) {
      isMicrosoft = true;
      break;
    }
  }

  // We have to check with every request if the requested URL is assigned with MAC
  // because the user can assign URLs at any given time (needs MAC Events)
  if (macAddonEnabled) {
    const macAssigned = await getMACAssignment(options.url);
    if (macAssigned) {
      // This URL is assigned with MAC, so we don't handle this request
      return;
    }
  }

  const tab = await browser.tabs.get(options.tabId);
  const tabCookieStoreId = tab.cookieStoreId;
  if (isMicrosoft) {
    if (tabCookieStoreId !== microsoftCookieStoreId && !tab.incognito) {
      // See https://github.com/mozilla/contain-microsoft/issues/23
      // Sometimes this add-on is installed but doesn't get a microsoftCookieStoreId ?
      if (microsoftCookieStoreId) {
        if (shouldCancelEarly(tab, options)) {
          return { cancel: true };
        }
        browser.tabs.create({
          url: requestUrl.toString(),
          cookieStoreId: microsoftCookieStoreId,
          active: tab.active,
          index: tab.index,
          windowId: tab.windowId
        });
        browser.tabs.remove(options.tabId);
        return { cancel: true };
      }
    }
  } else {
    if (tabCookieStoreId === microsoftCookieStoreId) {
      if (shouldCancelEarly(tab, options)) {
        return { cancel: true };
      }
      browser.tabs.create({
        url: requestUrl.toString(),
        active: tab.active,
        index: tab.index,
        windowId: tab.windowId
      });
      browser.tabs.remove(options.tabId);
      return { cancel: true };
    }
  }
}

(async function init() {
  await setupMACAddonManagementListeners();
  macAddonEnabled = await isMACAddonEnabled();

  await setupContainer();
  clearMicrosoftCookies();
  generateMicrosoftHostREs();

  // Add the request listener
  browser.webRequest.onBeforeRequest.addListener(containMicrosoft, { urls: ["<all_urls>"], types: ["main_frame"] }, ["blocking"]);

  // Clean up canceled requests
  browser.webRequest.onCompleted.addListener((options) => {
    if (canceledRequests[options.tabId]) {
      delete canceledRequests[options.tabId];
    }
  }, { urls: ["<all_urls>"], types: ["main_frame"] });
  browser.webRequest.onErrorOccurred.addListener((options) => {
    if (canceledRequests[options.tabId]) {
      delete canceledRequests[options.tabId];
    }
  }, { urls: ["<all_urls>"], types: ["main_frame"] });
})();
