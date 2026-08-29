import { getSettings } from './settingsManager'
import { finishRequest } from './requestManager'
import { buildMessageId, extractRequestIdFromReferences } from './emailManager'

const QUOTE_MARKERS: RegExp[] = [
  /^-{2,}\s*原始邮件\s*-{2,}$/,
  /^-{2,}\s*原始消息\s*-{2,}$/,
  /^-{2,}\s*Forwarded message\s*-{2,}$/i,
  /^_{2,}\s*Original Message\s*_{2,}$/i,
  /^-{2,}\s*Original Message\s*-{2,}$/i,
  /^\s*On\b.*\bwrote\s*:\s*$/i,
  /^在\s*\S.*\s写道\s*[:：]?\s*$/,
  /^发件人\s*[:：]/,
  /^发送自\s*(我的\s*)?(iPhone|iPad|Android).*$/i,
  /^Sent from my (iPhone|iPad|Android).*$/i,
  /^_{5,}\s*$/
]

// Strip the quoted conversation below a reply so only the admin's own words are delivered
export const stripQuotedText = (text: string): string => {
  const lines = text.replace(/\r\n/g, '\n').split('\n')
  let end = lines.length
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? ''
    if (QUOTE_MARKERS.some(re => re.test(line)) || /^\s*>/.test(line)) {
      // Ignore quote markers that appear in the first line of an empty-ish body (rare), keep at least one line
      if (i === 0 && lines.length > 1) continue
      end = i
      break
    }
  }
  const stripped = lines.slice(0, end).join('\n').trim()
  // Never deliver an empty reply: if stripping removed everything, fall back to the original text
  return stripped || text.trim()
}

export const matchRequestId = (subject: string | undefined, inReplyTo: string | string[] | undefined): string | null => {
  const fromRefs = extractRequestIdFromReferences(inReplyTo)
  if (fromRefs) return fromRefs
  if (subject) {
    const match = subject.match(/\[call-me-as-agent #([a-f0-9]+)\]/)
    if (match) return match[1] ?? null
  }
  return null
}

const isSenderAllowed = (fromAddress: string): boolean => {
  const s = getSettings()
  if (!s.emailReplyAllowFrom.trim()) return true
  const allowed = s.emailReplyAllowFrom.split(',').map(a => a.trim().toLowerCase()).filter(Boolean)
  return allowed.includes(fromAddress.toLowerCase())
}

interface ParsedReply {
  requestId: string
  fromAddress: string
  text: string
}

const parseReplyFromSource = async (source: Buffer, subject: string | undefined, inReplyTo: string | string[] | undefined): Promise<ParsedReply | null> => {
  const requestId = matchRequestId(subject, inReplyTo)
  if (!requestId) return null

  const { simpleParser } = await import('mailparser')
  const parsed = await simpleParser(source)
  const fromAddress = (parsed.from?.value?.[0]?.address) || ''

  let text = parsed.text || ''
  if (!text && parsed.html) {
    text = parsed.html
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|tr)>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
  }

  return {
    requestId,
    fromAddress,
    text: stripQuotedText(text)
  }
}

const pollOnce = async (): Promise<void> => {
  const s = getSettings()
  if (!s.enableEmailReply || !s.imapHost || !s.imapUser || !s.imapPass) return

  const { ImapFlow } = await import('imapflow')
  const client = new ImapFlow({
    host: s.imapHost,
    port: s.imapPort,
    secure: s.imapSecure,
    auth: { user: s.imapUser, pass: s.imapPass },
    logger: false,
    emitLogs: false
  })

  try {
    await client.connect()
    const lock = await client.getMailboxLock(s.imapMailbox || 'INBOX')
    try {
      const candidates: { uid: number, subject?: string, inReplyTo?: string | string[] }[] = []
      for await (const msg of client.fetch({ seen: false }, { uid: true, envelope: true })) {
        candidates.push({
          uid: msg.uid,
          subject: msg.envelope?.subject,
          inReplyTo: (msg.envelope as unknown as { inReplyTo?: string })?.inReplyTo
        })
      }

      for (const candidate of candidates) {
        const requestId = matchRequestId(candidate.subject, candidate.inReplyTo)
        if (!requestId) continue

        // Fetch the full source only for messages that look like replies to our notifications
        const fetched = await client.fetchOne(candidate.uid, { source: true }, { uid: true })
        if (!fetched || !fetched.source) continue
        const reply = await parseReplyFromSource(fetched.source, candidate.subject, candidate.inReplyTo)
        if (!reply) continue

        if (reply.fromAddress && !isSenderAllowed(reply.fromAddress)) {
          console.log(`[EmailReply] Rejected reply from ${reply.fromAddress} (not in allowlist) for request ${requestId}`)
        } else {
          try {
            await finishRequest(requestId, { content: reply.text, simulateStream: true })
            console.log(`[EmailReply] Delivered email reply to request ${requestId}`)
            if (reply.fromAddress && reply.text) {
              import('./emailManager').then(({ sendMail }) => sendMail({
                to: reply.fromAddress,
                subject: `[call-me-as-agent #${requestId}] Reply delivered`,
                inReplyTo: buildMessageId(requestId),
                text: `Your reply to request #${requestId} has been delivered to the client.\n\n你对请求 #${requestId} 的回复已成功发送给客户端。`
              })).catch(() => {})
            }
          } catch {
            console.log(`[EmailReply] Request ${requestId} no longer exists, skipping email reply`)
          }
        }

        // Mark processed so it is not picked up again
        try {
          await client.messageFlagsAdd(String(candidate.uid), ['\\Seen'], { uid: true })
        } catch {
          // best effort
        }
      }
    } finally {
      lock.release()
    }
  } finally {
    try {
      await client.logout()
    } catch { /* ignore */ }
  }
}

let pollTimer: ReturnType<typeof setInterval> | null = null
let polling = false
let lastRun = 0

const tick = async () => {
  const s = getSettings()
  if (!s.enableEmailReply || polling) return
  const interval = Math.max(10, s.emailPollInterval || 60) * 1000
  if (Date.now() - lastRun < interval) return
  polling = true
  lastRun = Date.now()
  try {
    await pollOnce()
  } catch (e) {
    console.error('[EmailReply] Poll failed:', e instanceof Error ? e.message : e)
  } finally {
    polling = false
  }
}

export const startEmailReplyPolling = () => {
  if (pollTimer) return
  pollTimer = setInterval(() => {
    void tick()
  }, 10000)
}
