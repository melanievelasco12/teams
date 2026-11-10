const fs = require('node:fs')
const path = require('node:path')

const publicDirectory = path.resolve(__dirname, '..', 'server', 'public')
const html = fs.readFileSync(path.join(publicDirectory, 'index.html'), 'utf8')
const referencedAssets = new Set()
const assetReferencePattern = /(?:src|href)=["'][^"']*\/?assets\/(index-[\w-]+\.(?:js|css))["']/g

for (const match of html.matchAll(assetReferencePattern)) {
  referencedAssets.add(match[1])
}

if (referencedAssets.size === 0) {
  throw new Error('No generated JavaScript or CSS assets were found in server/public/index.html.')
}

const assetsDirectory = path.join(publicDirectory, 'assets')
const staleAssetPattern = /^index-[\w-]+\.(?:js|css)$/
let removedCount = 0

for (const entry of fs.readdirSync(assetsDirectory, { withFileTypes: true })) {
  if (entry.isFile() && staleAssetPattern.test(entry.name) && !referencedAssets.has(entry.name)) {
    fs.unlinkSync(path.join(assetsDirectory, entry.name))
    removedCount += 1
  }
}

console.log(`Removed ${removedCount} stale build asset(s).`)
