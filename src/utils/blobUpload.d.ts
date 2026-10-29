export function upload(pathname: string, body: Blob, options: { access: 'public', handleUploadUrl: string }): Promise<{ url: string }>
