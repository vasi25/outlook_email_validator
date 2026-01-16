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
        // TODO: Replace this with API call later
        // const response = await fetch('https://your-api-url/api/valid-emails');
        // const data = await response.json();
        // const validEmails = data.emails;
        
        // Hardcoded database for now - REPLACE WITH API CALL LATER
        // const validEmails = [
        //     "alexvasilescu25@gmail.com",
        //     "vasilescualex25@gmail.com",
        // ];
        
        // // Get recipients from "To" field
        // Office.context.mailbox.item.to.getAsync(function(result) {
        //     if (result.status === Office.AsyncResultStatus.Succeeded) {
        //         const recipients = result.value;
        //         let invalidEmails = [];
                
        //         if (recipients.length === 0) {
        //             // No recipients - remove notification
        //             removeNotification(NOTIFICATION_KEY);
        //         } else {
        //             // Check each recipient
        //             recipients.forEach(function(recipient) {
        //                 const email = recipient.emailAddress.toLowerCase();
                        
        //                 if (!validEmails.includes(email)) {
        //                     invalidEmails.push(email);
        //                 }
        //             });
                    
        //             if (invalidEmails.length > 0) {
        //                 // Show notification for invalid emails
        //                 const emailList = invalidEmails.join(", ");
        //                 const message = "⚠️ Unverified recipients: " + emailList;
                        
        //                 Office.context.mailbox.item.notificationMessages.replaceAsync(
        //                     NOTIFICATION_KEY,
        //                     {
        //                         type: "informationalMessage",
        //                         message: message,
        //                         icon: "icon1",
        //                         persistent: false
        //                     }
        //                 );
        //             } else {
        //                 // All valid - remove notification
        //                 removeNotification(NOTIFICATION_KEY);
        //             }
        //         }
        //     }
        // });
        // fetch('http://localhost:8000/api/email/list')
        fetch('https://vasiemailvalidator.loca.lt')
            .then(response => response.json())
            .then(validEmails => {
                // validEmails is now the array from your API
                console.log('Fetched emails from API:', validEmails);
                
                // Get recipients from "To" field
                Office.context.mailbox.item.to.getAsync(function(result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const recipients = result.value;
                        let invalidEmails = [];
                        
                        if (recipients.length === 0) {
                            // No recipients - remove notification
                            removeNotification(NOTIFICATION_KEY);
                        } else {
                            // Check each recipient
                            recipients.forEach(function(recipient) {
                                const email = recipient.emailAddress.toLowerCase();
                                
                                if (!validEmails.includes(email)) {
                                    invalidEmails.push(email);
                                }
                            });
                            
                            if (invalidEmails.length > 0) {
                                // Show notification for invalid emails
                                const emailList = invalidEmails.join(", ");
                                const message = "⚠️ Unverified recipients: " + emailList;
                                
                                Office.context.mailbox.item.notificationMessages.replaceAsync(
                                    NOTIFICATION_KEY,
                                    {
                                        type: "informationalMessage",
                                        message: message,
                                        icon: "icon1",
                                        persistent: false
                                    }
                                );
                            } else {
                                // All valid - remove notification
                                removeNotification(NOTIFICATION_KEY);
                            }
                        }
                    }
                });
            })
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
