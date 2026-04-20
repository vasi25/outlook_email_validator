// Background script that runs automatically when composing emails
Office.onReady(function() {
    console.log('Email Validator background service started');
});

function onMessageComposeHandler(event) {
    console.log('New message compose detected - starting validation');
    startAutoValidation();
    event.completed();
}

let emailData = null;
let lastRecipientList = "";

function startAutoValidation() {
    fetch('https://addin.impuls-leasing.local/api/email/list-v2')
        .then(response => response.json())
        .then(response => {
            const rows = response.dataTable || response;
            if (!Array.isArray(rows)) return;

            // Process rows in batches to avoid freezing Outlook
            emailData = {};
            var i = 0;
            function processBatch() {
                var end = Math.min(i + 1000, rows.length);
                for (; i < end; i++) {
                    var e = rows[i].email.toLowerCase();
                    if (!emailData[e]) emailData[e] = [];
                    emailData[e].push(rows[i].client);
                }
                if (i < rows.length) {
                    setTimeout(processBatch, 0);
                } else {
                    console.log('Batch done, emailData size:', Object.keys(emailData).length);
                    pollRecipients();
                    setInterval(pollRecipients, 2000);
                }
            }
            processBatch();
        })
        .catch(error => {
            console.error('Failed to fetch valid emails:', error);
        });
}

function pollRecipients() {
    Office.context.mailbox.item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.log('to.getAsync failed:', result.error);
            return;
        }
        const resolved = result.value.filter(r => r.emailAddress);
        const key = resolved.map(r => r.emailAddress.toLowerCase()).sort().join(",");
        console.log('poll key:', key, 'last:', lastRecipientList);
        if (key === lastRecipientList) return;
        lastRecipientList = key;
        checkRecipients(resolved);
    });
}

function checkRecipients(recipients) {
    const NOTIFICATION_KEY = "emailValidatorWarning";

    if (recipients.length === 0) {
        removeNotification(NOTIFICATION_KEY);
        return;
    }

    let verifiedParts = [];
    let invalidEmails = [];

    recipients.forEach(function(recipient) {
        const email = recipient.emailAddress.toLowerCase();
        const username = email.split('@')[0];

        if (emailData[email]) {
            verifiedParts.push(username + ": " + emailData[email].join(", "));
        } else {
            invalidEmails.push(email);
        }
    });

    var unverifiedStr = invalidEmails.length > 0 ? "⚠️ Unverified: " + invalidEmails.join(", ") : "";
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

    console.log('Setting banner:', message);
    Office.context.mailbox.item.notificationMessages.replaceAsync(
        NOTIFICATION_KEY,
        {
            type: "informationalMessage",
            message: message,
            icon: "icon1",
            persistent: false
        },
        function(r) { console.log('replaceAsync result:', r.status, r.error); }
    );
}

function removeNotification(key) {
    Office.context.mailbox.item.notificationMessages.removeAsync(key, function(asyncResult) {});
}

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
