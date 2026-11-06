import axios from './axios'
import fetchRequest from './fetch'

// Self-hosted backend: serverless functions live under /api on the same Vercel domain.
export const SERVER_URL = '/api'

interface ImageSearchPayload {
  query: string;
  orientation?: 'landscape' | 'portrait' | 'square' | 'all';
  locale?: 'zh' | 'en';
  order?: 'popular' | 'latest';
  size?: 'large' | 'medium' | 'small';
  image_type?: 'all' | 'photo' | 'illustration' | 'vector';
  page?: number;
  per_page?: number;
}

interface AIPPTOutlinePayload {
  content: string
  language: string
  provider?: string
  model: string
}

interface AIPPTPayload {
  content: string
  language: string
  style: string
  provider?: string
  model: string
}

interface AIWritingPayload {
  content: string
  command: string
}

export default {
  getMockData(filename: string): Promise<any> {
    return axios.get(`./mocks/${filename}.json`)
  },

  searchImage(body: ImageSearchPayload): Promise<any> {
    return axios.post(`${SERVER_URL}/tools/img_search`, body)
  },

  AIPPT_Outline({
    content,
    language,
    provider,
    model,
  }: AIPPTOutlinePayload): Promise<any> {
    return fetchRequest(`${SERVER_URL}/tools/aippt_outline`, {
      method: 'POST',
      body: JSON.stringify({
        content,
        language,
        provider,
        model,
        stream: true,
      }),
    })
  },

  AIPPT({
    content,
    language,
    style,
    provider,
    model,
  }: AIPPTPayload): Promise<any> {
    return fetchRequest(`${SERVER_URL}/tools/aippt`, {
      method: 'POST',
      body: JSON.stringify({
        content,
        language,
        provider,
        model,
        style,
        stream: true,
      }),
    })
  },

  AI_Writing({
    content,
    command,
  }: AIWritingPayload): Promise<any> {
    return fetchRequest(`${SERVER_URL}/tools/ai_writing`, {
      method: 'POST',
      body: JSON.stringify({
        content,
        command,
        stream: true,
      }),
    })
  },

  // Edit chatbot: send current slide elements + an instruction, get updated elements back.
  AI_Edit(body: {
    elements: any[]
    instruction: string
    history?: { role: string, content: string }[]
    slideSize?: { width: number, height: number }
  }): Promise<{ reply: string, elements: any[] | null }> {
    return axios.post(`${SERVER_URL}/tools/chat_edit`, body)
  },

  // AI image generation: returns { src } as a data URL.
  AI_Image(body: { prompt: string, size?: string }): Promise<{ src: string }> {
    return axios.post(`${SERVER_URL}/tools/ai_image`, body)
  },
}