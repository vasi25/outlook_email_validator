// Background script that runs automatically when composing emails
Office.onReady(function() {
    console.log('Email Validator background service started');
});

// This function is called automatically when composing a new message
function onMessageComposeHandler(event) {
    console.log('New message compose detected - starting validation');
    startAutoValidation();
    event.completed();
}

let emailData = null;

function startAutoValidation() {
    fetch('https://addin.impuls-leasing.local/api/email/list-v2')
        .then(response => response.json())
        .then(response => {
            const rows = response.dataTable || response;
            if (!Array.isArray(rows)) return;

            emailData = {};
            rows.forEach(function(row) {
                const e = row.email.toLowerCase();
                if (!emailData[e]) emailData[e] = [];
                emailData[e].push(row.client);
            });
            console.log('emailData loaded, size:', Object.keys(emailData).length);

            pollRecipients();
            setInterval(pollRecipients, 2000);
        })
        .catch(error => {
            console.error('Failed to fetch valid emails:', error);
        });
}

function pollRecipients() {
    Office.context.mailbox.item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) return;
        const resolved = result.value.filter(r => r.emailAddress);
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

    console.log('Setting banner:', message);
    Office.context.mailbox.item.notificationMessages.replaceAsync(
        NOTIFICATION_KEY,
        {
            type: "informationalMessage",
            message: message,
            icon: "icon1",
            persistent: false
        },
        function(r) { console.log('replaceAsync:', r.status, r.error && r.error.message); }
    );
}

function removeNotification(key) {
    Office.context.mailbox.item.notificationMessages.removeAsync(key, function(asyncResult) {});
}

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
