// A real PDF file from JPEG pages (each page is one image, sized to it). JPEG bytes embed as they are (DCTDecode), so a
// PDF needs only these few objects and no library.
export interface JpegPage { bytes: Uint8Array, width: number, height: number }

export const jpegsToPdf = (pages: JpegPage[]): Blob => {
  const enc = new TextEncoder()
  const parts: Uint8Array[] = []
  const offsets: number[] = []
  let size = 0
  const put = (p: string | Uint8Array) => {
    const b = typeof p === 'string' ? enc.encode(p) : p
    parts.push(b)
    size += b.length
  }
  const object = (n: number, head: string, stream?: Uint8Array) => {
    offsets[n] = size
    put(`${n} 0 obj\n${head}\n`)
    if (stream) {
      put('stream\n')
      put(stream)
      put('\nendstream\n')
    }
    put('endobj\n')
  }
  const pageObj = (i: number) => 3 + i * 3
  put('%PDF-1.4\n')
  object(1, '<< /Type /Catalog /Pages 2 0 R >>')
  object(2, `<< /Type /Pages /Count ${pages.length} /Kids [${pages.map((_, i) => `${pageObj(i)} 0 R`).join(' ')}] >>`)
  pages.forEach((p, i) => {
    const n = pageObj(i)
    const draw = enc.encode(`q ${p.width} 0 0 ${p.height} 0 0 cm /Im0 Do Q`)
    object(n, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${p.width} ${p.height}] /Contents ${n + 1} 0 R /Resources << /XObject << /Im0 ${n + 2} 0 R >> >> >>`)
    object(n + 1, `<< /Length ${draw.length} >>`, draw)
    object(n + 2, `<< /Type /XObject /Subtype /Image /Width ${p.width} /Height ${p.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.bytes.length} >>`, p.bytes)
  })
  const total = 3 + pages.length * 3
  const xref = size
  put(`xref\n0 ${total}\n0000000000 65535 f \n`)
  for (let n = 1; n < total; n++) put(`${String(offsets[n]).padStart(10, '0')} 00000 n \n`)
  put(`trailer\n<< /Size ${total} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`)
  return new Blob(parts as BlobPart[], { type: 'application/pdf' })
}
