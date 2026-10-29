// Tells the panel where uploads go: Vercel Blob in production, the dev server's local folder otherwise.
export default function handler(_req: any, res: any) {
  res.status(200).json({ blob: process.env.BLOB_READ_WRITE_TOKEN ? 'vercel' : 'local' })
}
