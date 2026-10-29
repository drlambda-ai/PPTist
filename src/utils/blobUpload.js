// Plain JS on purpose: @vercel/blob's types pull in Node's, which changes setTimeout's type across the browser
// code. blobUpload.d.ts gives the one signature the app needs.
export const upload = async (...args) => (await import('@vercel/blob/client')).upload(...args)
