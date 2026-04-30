let mailboxItem;

Office.initialize = function (reason)
{
    console.log("*******************************************************");
    mailboxItem = Office.context.mailbox.item;
}

function validateMessage(event)
{
    console.log("Start validation stream");

    const message =
    {
        type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
        message: 'Message blocked.',
        icon: "Icon.16x16",
        actions:
        [
            {
                actionText: "Show task pane",
                actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
                commandId: "msgComposeOpenPaneButton",
                contextData: "{''}",
            },
        ],
    };

    mailboxItem.notificationMessages.addAsync("action", message);

    event.completed({ allowEvent: false });
    return;
}
