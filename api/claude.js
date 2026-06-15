export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' })
  }

  const apiKey = process.env.ANTHROPIC_API_KEY
  if (!apiKey) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured' })
  }

  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(body),
    })

    const text = await response.text()
    console.log('Anthropic status:', response.status)
    console.log('Anthropic response:', text.slice(0, 500))

    try {
      const data = JSON.parse(text)
      return res.status(response.status).json(data)
    } catch {
      return res.status(500).json({ error: 'Invalid JSON from Anthropic', raw: text.slice(0, 300) })
    }

  } catch (err) {
    console.error('Proxy error:', err.message)
    return res.status(500).json({ error: err.message })
  }
}

export const config = {
  api: { bodyParser: { sizeLimit: '4mb' } },
}
