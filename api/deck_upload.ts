import { handleUpload, type HandleUploadBody } from '@vercel/blob/client'
import { readBody } from './_llm.js'

// Browser -> Vercel Blob direct upload (templates are often bigger than the 4.5MB function body limit).
export default async function handler(req: any, res: any) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'POST only' }); return 
  }
  try {
    const body = (await readBody(req)) as HandleUploadBody
    const json = await handleUpload({
      body,
      request: req,
      onBeforeGenerateToken: () => Promise.resolve({ addRandomSuffix: true, maximumSizeInBytes: 50 * 1024 * 1024 }),
      onUploadCompleted: () => Promise.resolve(),
    })
    res.status(200).json(json)
  }
  catch (e: any) {
    res.status(400).json({ error: String(e?.message || e) })
  }
}
