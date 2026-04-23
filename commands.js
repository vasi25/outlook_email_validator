Office.onReady(function() {
    console.log('Email Validator background service started');
});

let emailData = null;
let fetchPromise = null;
let handlerRegistered = false;
let lastBannerMessage = null;

function loadEmailData() {
    if (emailData !== null) return Promise.resolve(emailData);
    if (fetchPromise) return fetchPromise;

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
            console.log('emailData loaded, size=' + Object.keys(emailData).length);
            fetchPromise = null;
            return emailData;
        })
        .catch(function(error) {
            console.error('Failed to fetch valid emails:', error);
            fetchPromise = null;
            return Promise.reject(error);
        });

    return fetchPromise;
}

function onMessageComposeHandler(event) {
    loadEmailData()
        .then(function() {
            checkCurrentRecipients();
            if (!handlerRegistered) {
                Office.context.mailbox.item.addHandlerAsync(
                    Office.EventType.RecipientsChanged,
                    function() { checkCurrentRecipients(); },
                    function(result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            handlerRegistered = true;
                        } else {
                            console.error('addHandlerAsync failed:', result.error);
                        }
                    }
                );
            }
            event.completed();
        })
        .catch(function() {
            event.completed();
        });
}

function checkCurrentRecipients() {
    Office.context.mailbox.item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) return;
        var resolved = result.value.filter(function(r) { return r.emailAddress; });
        checkRecipients(resolved);
    });
}

function checkRecipients(recipients) {
    const NOTIFICATION_KEY = "emailValidatorWarning";

    if (recipients.length === 0) {
        if (lastBannerMessage === null) return;
        lastBannerMessage = null;
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

    if (message === lastBannerMessage) return;
    lastBannerMessage = message;

    Office.context.mailbox.item.notificationMessages.replaceAsync(
        NOTIFICATION_KEY,
        {
            type: "informationalMessage",
            message: message,
            icon: "Icon.16x16",
            persistent: false
        }
    );
}

function removeNotification(key) {
    Office.context.mailbox.item.notificationMessages.removeAsync(key, function() {});
}

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
