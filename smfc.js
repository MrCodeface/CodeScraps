Office.onReady(() => {
    if (Office.context.mailbox.item) {
        setFromAddress();
    }
});

function setFromAddress() {
    Office.context.mailbox.getUserIdentityAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            let userEmail = result.value;
            let sharedMailboxes = {
                "shared1@example.com": "shared1@example.com",
                "shared2@example.com": "shared2@example.com"
            };

            let fromAddress = sharedMailboxes[userEmail] || userEmail;

            Office.context.mailbox.item.from.setAsync(fromAddress, function(response) {
                if (response.status === Office.AsyncResultStatus.Failed) {
                    console.error("Failed to set From address:", response.error);
                }
            });
        }
    });
}
