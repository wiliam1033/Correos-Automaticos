export const sendOutlookEmail = async (
  accessToken: string,
  to: string,
  subject: string,
  body: string,
  attachment?: File | null
) => {
  let attachmentsArray: any[] = [];

  if (attachment) {
    const base64Data = await new Promise<string>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        resolve(result.split(',')[1]);
      };
      reader.readAsDataURL(attachment);
    });

    attachmentsArray.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: attachment.name,
      contentType: attachment.type || 'application/octet-stream',
      contentBytes: base64Data
    });
  }

  const payload = {
    message: {
      subject: subject,
      body: {
        contentType: 'Text',
        content: body
      },
      toRecipients: [
        {
          emailAddress: {
            address: to
          }
        }
      ],
      attachments: attachmentsArray.length > 0 ? attachmentsArray : undefined
    },
    saveToSentItems: 'true'
  };

  const res = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  });

  if (!res.ok) {
    const errorData = await res.json().catch(() => null);
    throw new Error(`Failed to send Outlook email to ${to}: ${errorData?.error?.message || res.statusText}`);
  }
};
