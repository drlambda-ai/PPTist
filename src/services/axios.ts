import { t } from '@/i18n'
import axios from 'axios'
import message from '@/utils/message'
import { host, route } from '@/embed'

const instance = axios.create({ timeout: 1000 * 300 })

const chatslide = host
if (chatslide) {
  instance.interceptors.request.use(async config => {
    const to = config.url && route(config.url)
    if (to) {
      config.url = to
      config.headers.set('Authorization', `Bearer ${await chatslide.token()}`)
    }
    return config
  })
}

instance.interceptors.response.use(
  response => {
    if (response.status >= 200 && response.status < 400) {
      return Promise.resolve(response.data)
    }

    message.error(t('未知的请求错误！'))
    return Promise.reject(response)
  },
  error => {
    if (error && error.response) {
      if (error.response.status >= 400 && error.response.status < 500) {
        return Promise.reject(error.message)
      }
      else if (error.response.status >= 500) {
        return Promise.reject(error.message)
      }
      
      message.error(t('服务器遇到未知错误！'))
      return Promise.reject(error.message)
    }

    message.error(t('连接到服务器失败 或 服务器响应超时！'))
    return Promise.reject(error)
  }
)

export default instance
