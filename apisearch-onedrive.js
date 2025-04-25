// /api/search-onedrive.js

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  const { folderName, query } = req.body;

  const token = await getAccessToken();

  // Get folder ID (you can cache this if needed)
  const folderRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${folderName}:/children`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );

  const files = await folderRes.json();

  // Simple name-based search (can be enhanced with file contents later)
  const filteredFiles = files.value.filter(file =>
    file.name.toLowerCase().includes(query.toLowerCase())
  );

  const simplified = filteredFiles.map(file => ({
    name: file.name,
    webUrl: file.webUrl
  }));

  res.status(200).json({ files: simplified });
}

async function getAccessToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        scope: "https://graph.microsoft.com/.default"
      }),
    }
  );

  const data = await res.json();
  return data.access_token;
}
