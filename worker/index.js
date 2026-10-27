/**
 * Cloudflare Worker for Microsoft Teams Bot Connector & Proactive Messaging
 * Exact same behaviour as server.js with support for Cloudflare KV / D1 / In-Memory storage
 */

// In-memory fallback if KV is not bound
const memoryStore = {
  conversations: new Map(),
  logs: []
};

function addLog(env, type, message, details = null) {
  const logEntry = {
    id: Date.now() + '-' + Math.random().toString(36).substr(2, 4),
    timestamp: new Date().toISOString(),
    type,
    message,
    details
  };
  memoryStore.logs.unshift(logEntry);
  if (memoryStore.logs.length > 100) memoryStore.logs.pop();
  console.log(`[${type}] ${message}`);
}

// -----------------------------------------------------------------
// Bot Authentication Service (Using standard Web API fetch)
// -----------------------------------------------------------------
let cachedToken = null;
let tokenExpiresAt = 0;

async function getAccessToken(env, customAppId, customAppPassword, customTenantId) {
  const clientId = customAppId || env.MICROSOFT_APP_ID;
  const clientSecret = customAppPassword || env.MICROSOFT_APP_PASSWORD;
  const tenant = customTenantId || env.MICROSOFT_TENANT_ID || 'botframework.com';

  if (!clientId || !clientSecret) {
    throw new Error("Missing Microsoft App ID or App Password. Check your environment variables.");
  }

  const now = Math.floor(Date.now() / 1000);
  if (cachedToken && tokenExpiresAt > now + 120) {
    return cachedToken;
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
  const body = new URLSearchParams();
  body.append('grant_type', 'client_credentials');
  body.append('client_id', clientId);
  body.append('client_secret', clientSecret);
  body.append('scope', 'https://api.botframework.com/.default');

  const res = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });

  if (!res.ok) {
    const errText = await res.text();
    console.error('[BotAuth] Token failed:', errText);
    throw new Error(`Token request failed: ${errText}`);
  }

  const data = await res.json();
  cachedToken = data.access_token;
  tokenExpiresAt = now + (data.expires_in || 3600);
  return cachedToken;
}

// -----------------------------------------------------------------
// Bot Connector Service (Replies & Proactive Messages)
// -----------------------------------------------------------------
async function replyToActivity(env, serviceUrl, conversationId, activityId, replyActivity, credentials = {}) {
  const token = await getAccessToken(env, credentials.appId, credentials.appPassword, credentials.tenantId);
  const cleanServiceUrl = serviceUrl.replace(/\/+$/, '');
  const url = `${cleanServiceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities/${encodeURIComponent(activityId)}`;

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(replyActivity)
  });

  return await res.json();
}

async function sendProactiveMessage(env, serviceUrl, conversationId, activity, credentials = {}) {
  const token = await getAccessToken(env, credentials.appId, credentials.appPassword, credentials.tenantId);
  const cleanServiceUrl = serviceUrl.replace(/\/+$/, '');
  const url = `${cleanServiceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities`;

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(activity)
  });

  if (!res.ok) {
    const errorBody = await res.text();
    throw new Error(errorBody);
  }

  return await res.json();
}

// Helper to save conversation to KV or memory
async function saveConversation(env, conversationId, record) {
  memoryStore.conversations.set(conversationId, record);
  if (env.TEAMS_BOT_KV) {
    try {
      await env.TEAMS_BOT_KV.put(`conv:${conversationId}`, JSON.stringify(record));
    } catch (e) {
      console.error('KV put error:', e);
    }
  }
}

// Helper to get all conversations from KV or memory
async function getAllConversations(env) {
  if (env.TEAMS_BOT_KV) {
    try {
      const list = await env.TEAMS_BOT_KV.list({ prefix: 'conv:' });
      const records = [];
      for (const key of list.keys) {
        const val = await env.TEAMS_BOT_KV.get(key.name);
        if (val) records.push(JSON.parse(val));
      }
      if (records.length > 0) return records;
    } catch (e) {
      console.error('KV list error:', e);
    }
  }
  return Array.from(memoryStore.conversations.values());
}

// -----------------------------------------------------------------
// Main Worker Router
// -----------------------------------------------------------------
export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const pathname = url.pathname;

    // CORS Headers
    const corsHeaders = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization'
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }

    const jsonResponse = (data, status = 200) => {
      return new Response(JSON.stringify(data), {
        status,
        headers: {
          'Content-Type': 'application/json',
          ...corsHeaders
        }
      });
    };

    try {
      // 1. Health Status
      if (pathname === '/api/status' && request.method === 'GET') {
        const convs = await getAllConversations(env);
        return jsonResponse({
          status: 'online',
          runtime: 'cloudflare-worker',
          appIdConfigured: Boolean(env.MICROSOFT_APP_ID),
          conversationsCount: convs.length,
          timestamp: new Date().toISOString()
        });
      }

      // 2. Auth Test / Token Generation
      if (pathname === '/api/auth/token' && request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        const token = await getAccessToken(env, body.appId, body.appPassword, body.tenantId);
        addLog(env, 'AUTH', 'Successfully generated Bot Connector token');
        return jsonResponse({ success: true, token: token.substring(0, 15) + '...' });
      }

      // 3. Teams Webhook: Bot Framework sends messages here
      if (pathname === '/api/messages' && request.method === 'POST') {
        const activity = await request.json().catch(() => ({}));
        addLog(env, 'INCOMING', `Received activity type: ${activity.type}`, activity);

        if (activity.type === 'message') {
          const { serviceUrl, conversation, from, recipient, text, id: activityId } = activity;

          if (conversation && conversation.id && serviceUrl) {
            const record = {
              conversationId: conversation.id,
              serviceUrl,
              user: from,
              bot: recipient,
              lastMessage: text,
              updatedAt: new Date().toISOString()
            };
            await saveConversation(env, conversation.id, record);
          }

          // Auto-reply
          try {
            const replyActivity = {
              type: 'message',
              from: recipient,
              recipient: from,
              conversation: conversation,
              replyToId: activityId,
              text: `Echo from Bot Connector: Received "${text}"`
            };
            ctx.waitUntil(
              replyToActivity(env, serviceUrl, conversation.id, activityId, replyActivity)
                .then(() => addLog(env, 'OUTGOING', `Auto-replied to conversation: ${conversation.id}`))
                .catch(err => addLog(env, 'ERROR', `Failed to reply: ${err.message}`))
            );
          } catch (err) {
            addLog(env, 'ERROR', `Reply exception: ${err.message}`);
          }
        } else if (activity.type === 'conversationUpdate') {
          const { serviceUrl, conversation, recipient } = activity;
          if (conversation && conversation.id && serviceUrl) {
            const record = {
              conversationId: conversation.id,
              serviceUrl,
              bot: recipient,
              updatedAt: new Date().toISOString()
            };
            await saveConversation(env, conversation.id, record);
            addLog(env, 'INFO', `Bot added to conversation: ${conversation.id}`);
          }
        }

        return new Response('OK', { status: 200, headers: corsHeaders });
      }

      // 4. Send Proactive Message
      if (pathname === '/api/send-proactive' && request.method === 'POST') {
        const body = await request.json();
        const { serviceUrl, conversationId, messageType, title, message, buttons, credentials } = body;

        if (!serviceUrl || !conversationId) {
          return jsonResponse({ error: 'serviceUrl and conversationId are required.' }, 400);
        }

        const convs = await getAllConversations(env);
        const convData = convs.find(c => c.conversationId === conversationId) || {};
        const botInfo = convData.bot || { id: credentials?.appId || env.MICROSOFT_APP_ID || 'bot-id', name: 'System Bot' };
        const userInfo = convData.user || { id: 'recipient-id', name: 'User' };
        const conversationInfo = { id: conversationId };

        let activity;
        if (messageType === 'card') {
          activity = {
            type: 'message',
            from: botInfo,
            recipient: userInfo,
            conversation: conversationInfo,
            attachmentLayout: 'list',
            attachments: [
              {
                contentType: 'application/vnd.microsoft.card.thumbnail',
                content: {
                  title: title || 'System Notification',
                  text: message || 'Automated background alert',
                  buttons: (buttons || []).map(b => ({
                    type: b.type || 'openUrl',
                    title: b.title,
                    value: b.value
                  }))
                }
              }
            ]
          };
        } else {
          activity = {
            type: 'message',
            from: botInfo,
            recipient: userInfo,
            conversation: conversationInfo,
            text: message || 'Automated system alert from background process.'
          };
        }

        const result = await sendProactiveMessage(env, serviceUrl, conversationId, activity, credentials);
        addLog(env, 'PROACTIVE', `Proactive message sent to ${conversationId}`, result);
        return jsonResponse({ success: true, result });
      }

      // 5. Get Conversations & Logs for UI
      if (pathname === '/api/conversations' && request.method === 'GET') {
        const convs = await getAllConversations(env);
        return jsonResponse(convs);
      }

      if (pathname === '/api/logs' && request.method === 'GET') {
        return jsonResponse(memoryStore.logs);
      }

      // 6. Mock Incoming
      if (pathname === '/api/mock-incoming' && request.method === 'POST') {
        const mockConversationId = '19:mock-teams-chat-id-' + Math.floor(Math.random() * 1000);
        const mockRecord = {
          conversationId: mockConversationId,
          serviceUrl: 'https://smba.trafficmanager.net/teams',
          user: { id: 'user-aad-object-id', name: 'Alex Johnson' },
          bot: { id: env.MICROSOFT_APP_ID || 'bot-id', name: 'NotificationBot' },
          lastMessage: 'Hello Bot! (Simulated)',
          updatedAt: new Date().toISOString()
        };
        await saveConversation(env, mockConversationId, mockRecord);
        addLog(env, 'INFO', `Simulated user joined conversation ${mockConversationId}`);
        return jsonResponse({ success: true, conversation: mockRecord });
      }

      return new Response('Not Found', { status: 404, headers: corsHeaders });
    } catch (error) {
      addLog(env, 'ERROR', `Worker error: ${error.message}`);
      return jsonResponse({ success: false, error: error.message }, 500);
    }
  }
};
