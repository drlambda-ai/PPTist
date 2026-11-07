import { readFileSync, readdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import path from 'node:path'
import process from 'node:process'
import ts from 'typescript'

const root = fileURLToPath(new URL('../', import.meta.url))
const read = file => JSON.parse(readFileSync(path.join(root, file), 'utf8'))
const locales = read('src/locales.json')
const english = read('src/locales/en.json')
const parameters = text => [...text.matchAll(/\{p\d+\}/g)].map(m => m[0]).sort().join(',')
const errors = []
for (const locale of locales) {
  let messages
  try {
    messages = read(`src/locales/${locale}.json`)
  }
  catch {
    errors.push(`${locale}: missing or unreadable catalog`); continue
  }
  for (const [key, source] of Object.entries(english)) {
    const value = messages[key]
    if (typeof value !== 'string' || !value.trim()) errors.push(`${locale}: missing ${key}`)
    else if (parameters(value) !== parameters(source)) errors.push(`${locale}: parameters differ for ${key}`)
  }
  for (const key of Object.keys(messages)) if (!(key in english)) errors.push(`${locale}: extra ${key}`)
}
// Check authored t() calls, including Vue template expressions, against the same source catalog.
function inspect(directory) {
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name)
    if (entry.isDirectory()) {
      inspect(file); continue
    }
    if (!/\.(ts|vue)$/.test(entry.name)) continue
    const text = readFileSync(file, 'utf8')
    const chunks = entry.name.endsWith('.vue')
      ? [...text.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1])
      : [text]
    // Template t calls have literal keys; decode JS escapes using the TS parser.
    if (entry.name.endsWith('.vue')) chunks.push(...[...text.slice(0, text.indexOf('<script')).matchAll(/\bt\(("(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*')/g)].map(m => `t(${m[1]})`))
    for (const chunk of chunks) {
      const ast = ts.createSourceFile(file, chunk, ts.ScriptTarget.Latest, true)
      const visit = node => {
        if (ts.isCallExpression(node) && node.expression.getText(ast) === 't' && node.arguments.length) {
          const key = node.arguments[0]
          if (ts.isStringLiteral(key) && !(key.text in english)) errors.push(`${path.relative(root, file)}: unknown key ${key.text}`)
        }
        ts.forEachChild(node, visit)
      }
      visit(ast)
    }
  }
}
inspect(path.join(root, 'src'))
if (errors.length) {
  process.stderr.write(errors.join('\n') + '\n')
  process.exit(1)
}
process.stdout.write(`i18n: ${locales.length} locales, ${Object.keys(english).length} keys; calls and parameters match\n`)
