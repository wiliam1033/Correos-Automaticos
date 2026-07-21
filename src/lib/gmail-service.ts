export const sendEmail = async (
  accessToken: string,
  to: string,
  subject: string,
  body: string,
  attachment?: File | null
) => {
  const utf8Subject = `=?utf-8?B?${btoa(unescape(encodeURIComponent(subject)))}?=`;
  const utf8Body = btoa(unescape(encodeURIComponent(body)));
  
  let message = '';

  if (attachment) {
    const boundary = 'foo_bar_baz_boundary';
    
    // Read file as base64
    const base64Data = await new Promise<string>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        resolve(result.split(',')[1]);
      };
      reader.readAsDataURL(attachment);
    });

    const messageParts = [
      `To: ${to}`,
      `Subject: ${utf8Subject}`,
      'MIME-Version: 1.0',
      `Content-Type: multipart/mixed; boundary="${boundary}"`,
      '',
      `--${boundary}`,
      'Content-Type: text/plain; charset="UTF-8"',
      'Content-Transfer-Encoding: base64',
      '',
      utf8Body,
      '',
      `--${boundary}`,
      `Content-Type: ${attachment.type || 'application/octet-stream'}; name="${attachment.name.replace(/"/g, '\\"')}"`,
      'Content-Transfer-Encoding: base64',
      `Content-Disposition: attachment; filename="${attachment.name.replace(/"/g, '\\"')}"`,
      '',
      base64Data,
      '',
      `--${boundary}--`,
    ];
    message = messageParts.join('\r\n');
  } else {
    const messageParts = [
      `To: ${to}`,
      `Subject: ${utf8Subject}`,
      'MIME-Version: 1.0',
      'Content-Type: text/plain; charset="UTF-8"',
      'Content-Transfer-Encoding: base64',
      '',
      utf8Body,
    ];
    message = messageParts.join('\r\n');
  }
  
  // Safely encode the entire message to Base64URL, handling UTF-8 characters natively
  const blob = new Blob([message], { type: 'text/plain' });
  const base64Message = await new Promise<string>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve((reader.result as string).split(',')[1]);
    reader.readAsDataURL(blob);
  });

  const encodedMessage = base64Message
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');

  const res = await fetch(
    'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        raw: encodedMessage,
      }),
    }
  );

  if (!res.ok) {
    const errorData = await res.json();
    throw new Error(`Failed to send email to ${to}: ${errorData.error?.message || res.statusText}`);
  }

  return await res.json();
};
