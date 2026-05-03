let mailboxItem = null;

Office.initialize = function (reason)
{
    console.log("*******************************************************");
    mailboxItem = Office.context.mailbox.item;
};

function obtainAttachmentsInfo()
{
    return new Promise(
        (resolve, reject) =>
        {
            mailboxItem.getAttachmentsAsync(
                function(asyncResult)
                {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                    {
                        resolve(asyncResult.value);
                    }
                    else
                    {
                        reject(asyncResult.error);
                    }
                });
        });
}

function obtainAttachmentContent(anId)
{
    return new Promise(
        (resolve, reject) =>
        {
            console.log("Obtain attchament content: " + anId);
            this.mailboxItem.getAttachmentContentAsync(
                anId,
                function(asyncResult)
                {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                    {
                        resolve(asyncResult.value);
                    }
                    else
                    {
                        reject(asyncResult.error);
                    }
                });
        });
}

function validateMessage(event)
{
    console.log("Start validation stream");

    const attachmentsInfo = obtainAttachmentsInfo();

    if (!attachmentsInfo || attachmentsInfo.length == 0)
    {
        console.log("no attachments");
        event.completed({ allowEvent: true, });
    }

    const attachmentsInfoJSON = JSON.stringify(attachmentsInfo);

    let attachmentsPromises = [];

    for (let i = 0; i < attachmentsInfo.length; i++)
    {
        const attachInfo = attachmentsInfo[i];
        if (attachInfo.attachmentType === Office.MailboxEnums.AttachmentType.Cloud)
        {
            continue;
        }

        console.log(`Attachment ID: ${attachInfo.id}. Name: ${attachInfo.name}, size: ${attachInfo.size}, type:${attachInfo.attachmentType}`);

        let attachmentPromise = obtainAttachmentContent(attachInfo.id)
            .then((attachmentContent) =>
            {
                console.log(`Got attachment: name: ${attachmentContent.name} content length: ${attachmentContent.content.length}`);
            })
            .catch((error) =>
            {
                console.error("Obtain attachment error:", error);
            });
        attachmentsPromises.push(attachmentPromise);
    }

    Promise.all(attachmentsPromises);

    event.completed({ allowEvent: true, });
    return;
}
