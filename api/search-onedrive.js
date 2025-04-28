async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed. Use POST.' });
  }

  const { query, folderUrl } = req.body;

  if (!query || !folderUrl) {
    return res.status(400).json({ error: 'Missing query or folderUrl' });
  }

  try {
    const accessToken = await getAccessToken();
    const shareId = encodeShareUrl(folderUrl);

    const response = await fetch(`https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/children`, {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    });

    const data = await response.json();

    if (!data.value) {
      return res.status(500).json({ error: 'Unable to retrieve files from OneDrive.' });
    }

    const matchingFiles = data.value
      .filter(file => file.name.toLowerCase().includes(query.toLowerCase()))
      .map(file => ({
        name: file.name,
        webUrl: file.webUrl
      }));

    res.status(200).json({ files: matchingFiles });
  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ error: 'Unexpected error' });
  }
}

async function getAccessToken() {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;

  const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'https://graph.microsoft.com/.default'
    })
  });

  const tokenData = await tokenRes.json();
  return tokenData.access_token;
}

function encodeShareUrl(url) {
  const base64 = Buffer.from(url).toString('base64');
  return `u!${base64.replace(/\//g, '_').replace(/\+/g, '-')}`;
}

module.exports = handler;
