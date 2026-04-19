import { list, head } from '@vercel/blob';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.method !== 'GET') return res.status(405).end();

  try {
    const { blobs } = await list({ prefix: 'dashboard-data.json' });
    if (!blobs.length) return res.status(404).json({ error: 'No data uploaded yet' });

    // For private blobs, fetch server-side using the token from the environment
    const token = process.env.BLOB_READ_WRITE_TOKEN;
    const response = await fetch(blobs[0].url, {
      headers: token ? { Authorization: `Bearer ${token}` } : {},
    });
    if (!response.ok) return res.status(502).json({ error: 'Failed to fetch blob' });

    const data = await response.json();
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json(data);
  } catch (err) {
    console.error('Data fetch error:', err);
    return res.status(500).json({ error: err.message });
  }
}
