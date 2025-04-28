async function handler(req, res) {
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed. Use POST.' });
  }

  res.setHeader('Access-Control-Allow-Credentials', true);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  const { query } = req.body;

  if (!query) {
    return res.status(400).json({ error: 'Missing query' });
  }

  try {
    console.log('Starting search for query:', query);
    const accessToken = await getAccessToken();
    console.log('Access Token acquired.');

    // 👇 Path to your secured folder
    const folderPath = 'Marketingas/FOTO'; // <<< CHANGE THIS TO YOUR FOLDER

    const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${folderPath}:/children`, {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    });

    const data = await response.json();
    console.log('Files data:', JSON.stringify(data, null, 2));

    if (!data.value) {
      console.error('Error: No value field in response.');
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
    console.error('Error in handler:', err);
    res.status(500).json({ error: 'Unexpected error' });
  }
}

async function getAccessToken() {
  try {
    const tenantId = 'bfc9924e-c574-4dad-ae2d-a46d1b6f1a1a';
    const clientId = 'f368c58b-2909-46bd-95ae-308e2222d3c8';
    const clientSecret = 'ukT8Q~JIZWAkBNuYIv4KKZVsc6bp0OLNvOXltbak'; // your latest secret

    console.log('Hardcoded credentials being used for token request.');

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
    console.log('Token response:', JSON.stringify(tokenData, null, 2));

    if (!tokenData.access_token) {
      throw new Error('Access token not returned');
    }

    return tokenData.access_token;
  } catch (err) {
    console.error('Error fetching access token:', err);
    throw err;
  }
}

module.exports = handler;
