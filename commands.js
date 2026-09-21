/*
 * Teltherm "Quote Required" Outlook add-in: function command
 *
 * One click on the button:
 *   1. makes sure the "Quote Required" category exists in the user's mailbox
 *   2. adds that category to the open email (keeping any existing categories)
 *   3. flags the email
 * Flow flw_QuoteReq_Flag then fires on the flag, creates the list record,
 * completes the flag, swaps the category to "Quote Logged" and moves the email.
 *
 * Office.js can add categories but cannot set a flag, so both are done in a
 * single Microsoft Graph PATCH using Nested App Authentication (NAA).
 *
 * Placeholders replaced by configure.ps1: 968071bf-b520-4ebf-b679-5f0ecb50fc7b, 914aa421-b6cb-4176-9642-b43451b7e386
 */

const CONFIG = {
  clientId: "968071bf-b520-4ebf-b679-5f0ecb50fc7b",
  authority: "https://login.microsoftonline.com/914aa421-b6cb-4176-9642-b43451b7e386",
  scopes: ["Mail.ReadWrite", "MailboxSettings.ReadWrite"],
  categoryRequired: "Quote Required",   // must match the flow's condition exactly
  categoryLogged: "Quote Logged",       // set by the flow once the record exists
  newCategoryColour: "preset0",         // Red, used only if the category doesn't exist yet
  graph: "https://graph.microsoft.com/v1.0"
};

let msalInstance = null;

async function getToken() {
  if (!msalInstance) {
    msalInstance = await msal.createNestablePublicClientApplication({
      auth: { clientId: CONFIG.clientId, authority: CONFIG.authority },
      cache: { cacheLocation: "localStorage" }
    });
  }
  const request = { scopes: CONFIG.scopes };
  try {
    return (await msalInstance.acquireTokenSilent(request)).accessToken;
  } catch (err) {
    if (err instanceof msal.InteractionRequiredAuthError) {
      // Only expected on first use if admin consent hasn't been granted.
      return (await msalInstance.acquireTokenPopup(request)).accessToken;
    }
    throw err;
  }
}

async function graph(token, method, path, body) {
  const res = await fetch(CONFIG.graph + path, {
    method,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json"
    },
    body: body ? JSON.stringify(body) : undefined
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph ${method} ${res.status}: ${text.slice(0, 200)}`);
  }
  return res.status === 204 ? null : res.json();
}

function notify(type, message) {
  const item = Office.context.mailbox.item;
  const details =
    type === "error"
      ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message }
      : {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message,
          icon: "Icon.16",
          persistent: false
        };
  return new Promise((resolve) => item.notificationMessages.replaceAsync("quoteRequired", details, resolve));
}

async function ensureMasterCategory(token) {
  // Non-fatal: Graph accepts a category name that isn't in the master list,
  // it just shows uncoloured. This keeps every rep's mailbox consistent.
  try {
    const list = await graph(token, "GET", "/me/outlook/masterCategories");
    const exists = (list.value || []).some((c) => c.displayName === CONFIG.categoryRequired);
    if (!exists) {
      await graph(token, "POST", "/me/outlook/masterCategories", {
        displayName: CONFIG.categoryRequired,
        color: CONFIG.newCategoryColour
      });
    }
  } catch (e) {
    console.warn("Could not check/create master category:", e);
  }
}

async function markQuoteRequired(event) {
  try {
    if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
      await notify("error", "This version of Outlook can't run the Quote Required button. Use Ctrl+Shift+5 instead.");
      return;
    }

    const mailbox = Office.context.mailbox;
    const restId = mailbox.convertToRestId(mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    if (!restId) throw new Error("Couldn't read the message ID.");
    const msgPath = "/me/messages/" + encodeURIComponent(restId);

    const token = await getToken();
    const msg = await graph(token, "GET", msgPath + "?$select=categories,flag");
    const categories = msg.categories || [];

    if (categories.includes(CONFIG.categoryLogged)) {
      await notify("info", "Already logged as a quote request.");
      return;
    }
    if (categories.includes(CONFIG.categoryRequired) && msg.flag && msg.flag.flagStatus === "flagged") {
      await notify("info", "Already marked Quote Required. It will be logged shortly.");
      return;
    }

    await ensureMasterCategory(token);

    const newCategories = categories.includes(CONFIG.categoryRequired)
      ? categories
      : categories.concat(CONFIG.categoryRequired);

    await graph(token, "PATCH", msgPath, {
      categories: newCategories,
      flag: { flagStatus: "flagged" }
    });

    await notify("info", "Marked Quote Required. It will leave your inbox once it's logged.");
  } catch (err) {
    console.error(err);
    await notify("error", "Quote Required failed: " + String(err.message || err).slice(0, 110));
  } finally {
    event.completed();
  }
}

Office.onReady(() => {});
Office.actions.associate("markQuoteRequired", markQuoteRequired);
