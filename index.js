const express = require('express');
const cors = require('cors');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;

// IMAP server configurations for each provider
const IMAP_CONFIGS = {
  gmail: { host: 'imap.gmail.com', port: 993, secure: true },
  outlook: { host: 'outlook.office365.com', port: 993, secure: true },
  icloud: { host: 'imap.mail.me.com', port: 993, secure: true },
};

// Detect provider from email
function detectProvider(email) {
  const domain = email.split('@')[1]?.toLowerCase();
  if (!domain) return null;

  if (domain === 'gmail.com' || domain === 'googlemail.com') return 'gmail';
  if (['outlook.com', 'hotmail.com', 'live.com', 'msn.com'].includes(domain)) return 'outlook';
  if (['icloud.com', 'me.com', 'mac.com'].includes(domain)) return 'icloud';

  return null;
}

// Create IMAP client (supports both password and OAuth)
async function createClient(email, password, provider, options = {}) {
  const config = IMAP_CONFIGS[provider];
  if (!config) throw new Error(`Unsupported provider: ${provider}`);

  let authConfig;

  // Check if using OAuth authentication
  if (options.authType === 'oauth' && options.accessToken) {
    // Use XOAUTH2 for Gmail OAuth
    authConfig = {
      user: email,
      accessToken: options.accessToken,
    };
    console.log(`Using OAuth authentication for ${email}`);
  } else {
    // Use regular password authentication
    authConfig = {
      user: email,
      pass: password,
    };
  }

  const client = new ImapFlow({
    host: config.host,
    port: config.port,
    secure: config.secure,
    auth: authConfig,
    logger: false,
  });

  await client.connect();
  return client;
}

// Health check
app.get('/', (req, res) => {
  res.json({ status: 'ok', message: 'Mail Sentinel IMAP Proxy' });
});

// Test connection
app.post('/test', async (req, res) => {
  const { email, password, authType, accessToken } = req.body;

  if (!email || (!password && !accessToken)) {
    return res.status(400).json({ error: 'Email and password/accessToken required' });
  }

  const provider = detectProvider(email);
  if (!provider) {
    return res.status(400).json({ error: 'Unsupported email provider' });
  }

  let client;
  try {
    client = await createClient(email, password, provider, { authType, accessToken });
    await client.logout();
    res.json({ success: true, provider, authType: authType || 'password' });
  } catch (error) {
    console.error('Connection test failed:', error.message);
    res.status(401).json({
      success: false,
      error: error.message || 'Authentication failed'
    });
  }
});

// Get email count
app.post('/count', async (req, res) => {
  const { email, password, authType, accessToken } = req.body;

  if (!email || (!password && !accessToken)) {
    return res.status(400).json({ error: 'Email and password/accessToken required' });
  }

  const provider = detectProvider(email);
  if (!provider) {
    return res.status(400).json({ error: 'Unsupported email provider' });
  }

  let client;
  try {
    client = await createClient(email, password, provider, { authType, accessToken });
    await client.mailboxOpen('INBOX');

    const status = await client.status('INBOX', { messages: true });
    await client.logout();

    res.json({ count: status.messages || 0 });
  } catch (error) {
    console.error('Count error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// Fetch emails
app.post('/fetch', async (req, res) => {
  const { email, password, limit = 50, offset = 0, folder = 'INBOX', authType, accessToken } = req.body;

  if (!email || (!password && !accessToken)) {
    return res.status(400).json({ error: 'Email and password/accessToken required' });
  }

  const provider = detectProvider(email);
  if (!provider) {
    return res.status(400).json({ error: 'Unsupported email provider' });
  }

  let client;
  try {
    client = await createClient(email, password, provider, { authType, accessToken });
    await client.mailboxOpen(folder);

    const status = await client.status(folder, { messages: true });
    const totalCount = status.messages || 0;

    if (totalCount === 0) {
      await client.logout();
      return res.json({ emails: [], totalCount: 0, hasMore: false });
    }

    // Calculate range (IMAP uses 1-based indexing, newest first)
    const start = Math.max(1, totalCount - offset - limit + 1);
    const end = Math.max(1, totalCount - offset);

    if (start > end) {
      await client.logout();
      return res.json({ emails: [], totalCount, hasMore: false });
    }

    const emails = [];

    // Fetch messages in reverse order (newest first)
    for await (const message of client.fetch(`${start}:${end}`, {
      envelope: true,
      flags: true,
      bodyStructure: true,
      source: true,
    })) {
      try {
        const parsed = await simpleParser(message.source);

        emails.push({
          id: message.uid.toString(),
          seq: message.seq,
          subject: parsed.subject || '(No Subject)',
          from: {
            name: parsed.from?.value?.[0]?.name || '',
            address: parsed.from?.value?.[0]?.address || '',
          },
          to: parsed.to?.value?.map(t => ({ name: t.name, address: t.address })) || [],
          date: parsed.date?.toISOString() || new Date().toISOString(),
          text: parsed.text?.substring(0, 10000) || '',
          html: parsed.html?.substring(0, 50000) || '',
          flags: message.flags ? Array.from(message.flags) : [],
          isRead: message.flags?.has('\\Seen') || false,
        });
      } catch (parseError) {
        console.error('Parse error for message:', message.uid, parseError.message);
      }
    }

    await client.logout();

    // Reverse to get newest first
    emails.reverse();

    res.json({
      emails,
      totalCount,
      hasMore: offset + emails.length < totalCount,
    });
  } catch (error) {
    console.error('Fetch error:', error.message);
    if (client) {
      try { await client.logout(); } catch (e) {}
    }
    res.status(500).json({ error: error.message });
  }
});

// Get single email by UID
app.post('/message/:uid', async (req, res) => {
  const { uid } = req.params;
  const { email, password, folder = 'INBOX', authType, accessToken } = req.body;

  if (!email || (!password && !accessToken)) {
    return res.status(400).json({ error: 'Email and password/accessToken required' });
  }

  const provider = detectProvider(email);
  if (!provider) {
    return res.status(400).json({ error: 'Unsupported email provider' });
  }

  let client;
  try {
    client = await createClient(email, password, provider, { authType, accessToken });
    await client.mailboxOpen(folder);

    const message = await client.fetchOne(uid, {
      envelope: true,
      flags: true,
      source: true,
    }, { uid: true });

    if (!message) {
      await client.logout();
      return res.status(404).json({ error: 'Message not found' });
    }

    const parsed = await simpleParser(message.source);

    await client.logout();

    res.json({
      id: uid,
      subject: parsed.subject || '(No Subject)',
      from: {
        name: parsed.from?.value?.[0]?.name || '',
        address: parsed.from?.value?.[0]?.address || '',
      },
      to: parsed.to?.value?.map(t => ({ name: t.name, address: t.address })) || [],
      date: parsed.date?.toISOString() || new Date().toISOString(),
      text: parsed.text || '',
      html: parsed.html || '',
      flags: message.flags ? Array.from(message.flags) : [],
      isRead: message.flags?.has('\\Seen') || false,
      attachments: parsed.attachments?.map(a => ({
        filename: a.filename,
        contentType: a.contentType,
        size: a.size,
      })) || [],
    });
  } catch (error) {
    console.error('Message fetch error:', error.message);
    if (client) {
      try { await client.logout(); } catch (e) {}
    }
    res.status(500).json({ error: error.message });
  }
});

// Search emails
app.post('/search', async (req, res) => {
  const { email, password, query, folder = 'INBOX', limit = 50, authType, accessToken } = req.body;

  if (!email || (!password && !accessToken) || !query) {
    return res.status(400).json({ error: 'Email, password/accessToken, and query required' });
  }

  const provider = detectProvider(email);
  if (!provider) {
    return res.status(400).json({ error: 'Unsupported email provider' });
  }

  let client;
  try {
    client = await createClient(email, password, provider, { authType, accessToken });
    await client.mailboxOpen(folder);

    // Search for messages containing the query in subject or body
    const uids = await client.search({
      or: [
        { subject: query },
        { body: query },
        { from: query },
      ],
    }, { uid: true });

    // Take only the most recent matches
    const recentUids = uids.slice(-limit).reverse();

    if (recentUids.length === 0) {
      await client.logout();
      return res.json({ emails: [], totalCount: 0 });
    }

    const emails = [];

    for await (const message of client.fetch(recentUids, {
      envelope: true,
      flags: true,
      source: true,
    }, { uid: true })) {
      try {
        const parsed = await simpleParser(message.source);

        emails.push({
          id: message.uid.toString(),
          subject: parsed.subject || '(No Subject)',
          from: {
            name: parsed.from?.value?.[0]?.name || '',
            address: parsed.from?.value?.[0]?.address || '',
          },
          date: parsed.date?.toISOString() || new Date().toISOString(),
          text: parsed.text?.substring(0, 5000) || '',
          isRead: message.flags?.has('\\Seen') || false,
        });
      } catch (parseError) {
        console.error('Parse error:', parseError.message);
      }
    }

    await client.logout();

    res.json({
      emails,
      totalCount: uids.length,
    });
  } catch (error) {
    console.error('Search error:', error.message);
    if (client) {
      try { await client.logout(); } catch (e) {}
    }
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Mail Sentinel IMAP Proxy running on port ${PORT}`);
  console.log(`Endpoints:`);
  console.log(`  POST /test   - Test connection`);
  console.log(`  POST /count  - Get email count`);
  console.log(`  POST /fetch  - Fetch emails`);
  console.log(`  POST /message/:uid - Get single email`);
  console.log(`  POST /search - Search emails`);
});
