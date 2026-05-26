export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Cookie');

  if (req.method === 'OPTIONS') return res.status(200).end();

  const apiKey = process.env.GEMINI_API_KEY;
  
  const spCookie = req.headers['x-sp-cookie'];
  let knowledgeContext = '';

  if (spCookie) {
    try {
      const spRes = await fetch(
        `https://jgsoffice.sharepoint.com/sites/URCELTestChannel/_api/web/lists/getbytitle('Chatbot Knowledge')/items?$select=Topic,Content`,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Cookie': spCookie
          }
        }
      );
      const spData = await spRes.json();
      if (spData.value?.length) {
        knowledgeContext = '\n\n## LIVE KNOWLEDGE BASE FROM URCEL SHAREPOINT:\n' +
          spData.value.map(i => `**${i.Topic}**: ${i.Content}`).join('\n');
      }
    } catch (e) {
      console.log('SharePoint fetch failed:', e.message);
    }
  }

  const body = JSON.parse(req.body);
  if (body.system_instruction?.parts?.[0]) {
    body.system_instruction.parts[0].text += knowledgeContext;
  }

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }
  );

  const data = await response.json();
  res.status(200).json(data);
}
