Office.onReady(function() {
    console.log('[EV] Office.onReady fired');
});

let emailData = null;
let fetchPromise = null;
let intervalId = null;
let pollCount = 0;

function loadEmailData() {
    if (emailData !== null) {
        console.log('[EV] loadEmailData: already loaded, size=' + Object.keys(emailData).length);
        return Promise.resolve(emailData);
    }
    if (fetchPromise) {
        console.log('[EV] loadEmailData: fetch already in progress');
        return fetchPromise;
    }

    console.log('[EV] loadEmailData: starting fetch');
    fetchPromise = fetch('https://addin.impuls-leasing.local/api/email/list-v2')
        .then(function(response) { return response.json(); })
        .then(function(response) {
            const rows = response.dataTable || response;
            emailData = {};
            if (Array.isArray(rows)) {
                rows.forEach(function(row) {
                    const e = row.email.toLowerCase();
                    if (!emailData[e]) emailData[e] = [];
                    emailData[e].push(row.client);
                });
            }
            console.log('[EV] loadEmailData: done, size=' + Object.keys(emailData).length);
            fetchPromise = null;
            return emailData;
        })
        .catch(function(error) {
            console.error('[EV] loadEmailData: fetch failed', error);
            fetchPromise = null;
            return Promise.reject(error);
        });

    return fetchPromise;
}

function onMessageComposeHandler(event) {
    console.log('[EV] onMessageComposeHandler called');
    loadEmailData()
        .then(function() {
            console.log('[EV] data ready, running first poll');
            checkCurrentRecipients();
            if (!intervalId) {
                console.log('[EV] registering setInterval');
                intervalId = setInterval(checkCurrentRecipients, 2000);
                console.log('[EV] intervalId=' + intervalId);
            } else {
                console.log('[EV] interval already registered, id=' + intervalId);
            }
            console.log('[EV] calling event.completed()');
            event.completed();
        })
        .catch(function(err) {
            console.error('[EV] error in handler', err);
            event.completed();
        });
}

function checkCurrentRecipients() {
    pollCount++;
    var thisPoll = pollCount;
    console.log('[EV] poll #' + thisPoll + ' fired');
    Office.context.mailbox.item.to.getAsync(function(result) {
        console.log('[EV] poll #' + thisPoll + ' getAsync status=' + result.status);
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error('[EV] poll #' + thisPoll + ' getAsync failed:', result.error && result.error.message);
            return;
        }
        console.log('[EV] poll #' + thisPoll + ' raw recipients count=' + result.value.length);
        result.value.forEach(function(r, i) {
            console.log('[EV] poll #' + thisPoll + ' recipient[' + i + '] display="' + r.displayName + '" email="' + r.emailAddress + '"');
        });
        var resolved = result.value.filter(function(r) { return r.emailAddress; });
        console.log('[EV] poll #' + thisPoll + ' resolved count=' + resolved.length);
        checkRecipients(resolved);
    });
}

function checkRecipients(recipients) {
    const NOTIFICATION_KEY = "emailValidatorWarning";

    if (recipients.length === 0) {
        console.log('[EV] no recipients, removing notification');
        removeNotification(NOTIFICATION_KEY);
        return;
    }

    var verifiedParts = [];
    var invalidEmails = [];

    recipients.forEach(function(recipient) {
        var email = recipient.emailAddress.toLowerCase();
        var username = email.split('@')[0];

        if (emailData[email]) {
            verifiedParts.push(username + ": " + emailData[email].join(", "));
        } else {
            invalidEmails.push(username);
        }
    });

    var unverifiedStr = invalidEmails.length > 0 ? "⚠️ Unverified: " + invalidEmails.join(" | ") : "";
    var message;

    if (verifiedParts.length === 0) {
        message = unverifiedStr;
    } else if (!unverifiedStr) {
        var full = "✅ " + verifiedParts.join(" | ");
        message = full.length > 150 ? full.substring(0, 147) + "..." : full;
    } else {
        var maxVerified = 150 - unverifiedStr.length - 3;
        var verifiedFull = "✅ " + verifiedParts.join(" | ");
        var verifiedSection = maxVerified > 0
            ? (verifiedFull.length <= maxVerified ? verifiedFull : verifiedFull.substring(0, maxVerified - 3) + "...")
            : null;
        message = verifiedSection ? verifiedSection + " | " + unverifiedStr : unverifiedStr;
    }

    console.log('[EV] setting banner: "' + message + '"');
    var details = {
        type: "informationalMessage",
        message: message,
        icon: "Icon.16x16",
        persistent: false
    };
    Office.context.mailbox.item.notificationMessages.removeAsync(NOTIFICATION_KEY, function() {
        Office.context.mailbox.item.notificationMessages.addAsync(NOTIFICATION_KEY, details, function(r) {
            console.log('[EV] addAsync status=' + r.status + (r.error ? ' error=' + r.error.message : ''));
        });
    });
}

function removeNotification(key) {
    Office.context.mailbox.item.notificationMessages.removeAsync(key, function() {});
}

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
