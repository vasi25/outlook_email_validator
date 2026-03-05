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

function startAutoValidation() {
    // Start checking immediately
    checkRecipients();
    
    // Then check every 2 seconds
    if (validationInterval) {
        clearInterval(validationInterval);
    }
    
    validationInterval = setInterval(function() {
        checkRecipients();
    }, 2000);
}

function checkRecipients() {
    const NOTIFICATION_KEY = "emailValidatorWarning";

    try {
        fetch('https://vasilocaladdin:3000/api/email/list')
            .then(response => response.json())
            .then(emailData => {
                // Get recipients from "To" field
                Office.context.mailbox.item.to.getAsync(function(result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const recipients = result.value;
                        let verifiedParts = [];
                        let invalidEmails = [];

                        if (recipients.length === 0) {
                            removeNotification(NOTIFICATION_KEY);
                        } else {
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
                    }
                });
            })
            .catch(error => {
                console.error('Failed to fetch valid emails:', error);
            });
    } catch (error) {
        console.error('Validation error:', error);
    }
}

function removeNotification(key) {
    Office.context.mailbox.item.notificationMessages.removeAsync(key, function(asyncResult) {
        // Ignore errors (notification might not exist)
    });
}

// Register the function for the LaunchEvent
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
