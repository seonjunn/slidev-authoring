// Minimal ZIP writer (deflate, UTF-8 names) — enough for an Office package.
import { deflateRawSync, crc32 } from 'node:zlib'

function dosTime(d) {
  return {
    time: (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2),
    date: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
  }
}

// entries: [[name, Buffer | string], ...] in the order they should appear
export function zip(entries, now = new Date()) {
  const { time, date } = dosTime(now)
  const locals = [], central = []
  let offset = 0
  for (const [name, content] of entries) {
    const data = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8')
    const nameBuf = Buffer.from(name, 'utf8')
    const packed = deflateRawSync(data, { level: 9 })
    const stored = packed.length >= data.length
    const body = stored ? data : packed
    const crc = crc32(data)
    const head = Buffer.alloc(30)
    head.writeUInt32LE(0x04034b50, 0); head.writeUInt16LE(20, 4); head.writeUInt16LE(0x0800, 6)
    head.writeUInt16LE(stored ? 0 : 8, 8); head.writeUInt16LE(time, 10); head.writeUInt16LE(date, 12)
    head.writeUInt32LE(crc, 14); head.writeUInt32LE(body.length, 18); head.writeUInt32LE(data.length, 22)
    head.writeUInt16LE(nameBuf.length, 26); head.writeUInt16LE(0, 28)
    locals.push(head, nameBuf, body)
    const dir = Buffer.alloc(46)
    dir.writeUInt32LE(0x02014b50, 0); dir.writeUInt16LE(20, 4); dir.writeUInt16LE(20, 6); dir.writeUInt16LE(0x0800, 8)
    dir.writeUInt16LE(stored ? 0 : 8, 10); dir.writeUInt16LE(time, 12); dir.writeUInt16LE(date, 14)
    dir.writeUInt32LE(crc, 16); dir.writeUInt32LE(body.length, 20); dir.writeUInt32LE(data.length, 24)
    dir.writeUInt16LE(nameBuf.length, 28); dir.writeUInt32LE(offset, 42)
    central.push(dir, nameBuf)
    offset += 30 + nameBuf.length + body.length
  }
  const dirSize = central.reduce((s, b) => s + b.length, 0)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10)
  end.writeUInt32LE(dirSize, 12); end.writeUInt32LE(offset, 16)
  return Buffer.concat([...locals, ...central, end])
}
