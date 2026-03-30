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

let validationInterval;
let lastRecipientList = "";
let emailData = null; // cached email list

function startAutoValidation() {
    // Fetch the email list once, then start polling recipients
    fetch('https://addin.impuls-leasing.local/api/email/list-v2')
        .then(response => response.json())
        .then(response => {
            const rows = response.dataTable || response;
            if (!Array.isArray(rows)) return;

            // Build lookup: { email: [clients] }
            emailData = {};
            rows.forEach(function(row) {
                const e = row.email.toLowerCase();
                if (!emailData[e]) emailData[e] = [];
                emailData[e].push(row.client);
            });

            // Start polling recipients now that we have the list
            pollRecipients();
            if (validationInterval) clearInterval(validationInterval);
            validationInterval = setInterval(pollRecipients, 2000);
        })
        .catch(error => {
            console.error('Failed to fetch valid emails:', error);
        });
}

function pollRecipients() {
    if (!emailData) return;

    Office.context.mailbox.item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) return;
        const key = result.value.map(r => r.emailAddress.toLowerCase()).sort().join(",");
        if (key === lastRecipientList) return;
        lastRecipientList = key;
        checkRecipients(result.value);
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
        // Unverified has priority — fit verified in remaining space
        var maxVerified = 150 - unverifiedStr.length - 3; // 3 for " | "
        var verifiedFull = "✅ " + verifiedParts.join(" | ");
        var verifiedSection = maxVerified > 0
            ? (verifiedFull.length <= maxVerified ? verifiedFull : verifiedFull.substring(0, maxVerified - 3) + "...")
            : null;
        message = verifiedSection ? verifiedSection + " | " + unverifiedStr : unverifiedStr;
    }

    Office.context.mailbox.item.notificationMessages.replaceAsync(
        NOTIFICATION_KEY,
        {
            type: "informationalMessage",
            message: message,
            icon: "icon1",
            persistent: false
        }
    );
}

function removeNotification(key) {
    Office.context.mailbox.item.notificationMessages.removeAsync(key, function(asyncResult) {
        // Ignore errors (notification might not exist)
    });
}

// Register the function for the LaunchEvent
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
