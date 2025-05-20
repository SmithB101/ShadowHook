
Office.actions.associate("forwardToPhishingTeam", function (event) {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    item.forwardAsync({
        toRecipients: ["phishing@kryptokloud.com"],
        body: "This email was reported as phishing."
    }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const metadata = {
                reporter: mailbox.userProfile.emailAddress,
                subject: item.subject,
                timeReported: new Date().toISOString(),
                itemId: item.itemId,
                conversationId: item.conversationId,
                from: item.from ? item.from.emailAddress : "",
                to: item.to ? item.to.map(r => r.emailAddress) : []
            };

            fetch("https://your-backend.kryptokloud.com/api/phishing-report", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(metadata)
            });

            mailbox.item.notificationMessages.addAsync("success", {
                type: "informationalMessage",
                message: "Email forwarded and report logged.",
                icon: "icon16",
                persistent: false
            });
        } else {
            console.error("Forwarding failed", result.error);
        }
        event.completed();
    });
});
