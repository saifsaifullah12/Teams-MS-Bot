import React, { useState, useEffect } from 'react';
import { 
  Bot, 
  Send, 
  Key, 
  MessageSquare, 
  Activity, 
  CheckCircle, 
  AlertTriangle, 
  RefreshCw, 
  ShieldCheck, 
  Layers, 
  Users,
  Copy,
  Terminal
} from 'lucide-react';

const API_BASE = import.meta.env.VITE_API_BASE || '/api';

export default function App() {
  const [appId, setAppId] = useState('');
  const [appPassword, setAppPassword] = useState('');
  const [tokenStatus, setTokenStatus] = useState(null);
  const [loadingToken, setLoadingToken] = useState(false);

  // Message Form State
  const [conversations, setConversations] = useState([]);
  const [selectedConv, setSelectedConv] = useState('');
  const [customServiceUrl, setCustomServiceUrl] = useState('https://smba.trafficmanager.net/teams');
  const [customConvId, setCustomConvId] = useState('');
  const [messageType, setMessageType] = useState('text');
  const [cardTitle, setCardTitle] = useState('System Build Alert');
  const [messageText, setMessageText] = useState('Build #4192 deployed successfully on Production Server!');
  const [sendingMsg, setSendingMsg] = useState(false);
  const [sendResult, setSendResult] = useState(null);

  // System Logs
  const [logs, setLogs] = useState([]);

  // Fetch initial data
  const fetchData = async () => {
    try {
      const convRes = await fetch(`${API_BASE}/conversations`);
      if (convRes.ok) {
        const convList = await convRes.json();
        setConversations(convList);
        setSelectedConv(prev => {
          if (!prev && convList.length > 0) return convList[0].conversationId;
          return prev || (convList.length > 0 ? convList[0].conversationId : '');
        });
      }

      const logRes = await fetch(`${API_BASE}/logs`);
      if (logRes.ok) {
        const logData = await logRes.json();
        setLogs(logData);
      }
    } catch (err) {
      console.error('Fetch error:', err);
    }
  };

  useEffect(() => {
    fetchData();
    const interval = setInterval(fetchData, 4000);
    return () => clearInterval(interval);
  }, []);

  const handleTestToken = async (e) => {
    e.preventDefault();
    setLoadingToken(true);
    setTokenStatus(null);
    try {
      const res = await fetch(`${API_BASE}/auth/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ appId, appPassword })
      });
      const data = await res.json();
      if (data.success) {
        setTokenStatus({ success: true, message: 'Valid! Token: ' + data.token });
      } else {
        setTokenStatus({ success: false, message: data.error });
      }
    } catch (err) {
      setTokenStatus({ success: false, message: 'Could not connect to Node.js backend on port 5000' });
    } finally {
      setLoadingToken(false);
      fetchData();
    }
  };

  const handleSendMessage = async (e) => {
    e.preventDefault();
    setSendingMsg(true);
    setSendResult(null);

    const activeConvId = selectedConv === 'custom' ? customConvId : selectedConv;
    const targetServiceUrl = selectedConv === 'custom' ? customServiceUrl : 
      conversations.find(c => c.conversationId === selectedConv)?.serviceUrl || customServiceUrl;

    if (!activeConvId) {
      setSendResult({ success: false, error: 'Please select or enter a Conversation ID' });
      setSendingMsg(false);
      return;
    }

    try {
      const payload = {
        serviceUrl: targetServiceUrl,
        conversationId: activeConvId,
        messageType,
        title: cardTitle,
        message: messageText,
        buttons: messageType === 'card' ? [
          { type: 'openUrl', title: 'View Logs', value: 'https://portal.azure.com' },
          { type: 'openUrl', title: 'Open Dashboard', value: 'http://localhost:5173' }
        ] : [],
        credentials: (appId && appPassword) ? { appId, appPassword } : null
      };

      const res = await fetch(`${API_BASE}/send-proactive`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (data.success) {
        setSendResult({ success: true, message: 'Message successfully sent to Teams Bot Connector!' });
      } else {
        setSendResult({ success: false, error: JSON.stringify(data.error) });
      }
    } catch (err) {
      setSendResult({ success: false, error: err.message });
    } finally {
      setSendingMsg(false);
      fetchData();
    }
  };

  const handleMockIncoming = async () => {
    try {
      await fetch(`${API_BASE}/mock-incoming`, { method: 'POST' });
      fetchData();
    } catch (err) {
      console.error(err);
    }
  };

  return (
    <div style={{ maxWidth: '1280px', margin: '0 auto', padding: '32px 20px' }}>
      {/* Top Header */}
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '32px', flexWrap: 'wrap', gap: '16px' }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '6px' }}>
            <div style={{ background: 'linear-gradient(135deg, #6366f1, #06b6d4)', padding: '10px', borderRadius: '12px', display: 'flex' }}>
              <Bot size={28} color="#fff" />
            </div>
            <h1 style={{ fontSize: '1.75rem', fontWeight: '700', letterSpacing: '-0.02em' }}>
              Microsoft Teams Bot Connector Hub
            </h1>
          </div>
          <p style={{ color: 'var(--text-muted)', fontSize: '0.95rem' }}>
            System-driven Proactive Messaging & REST Activity Controller (Node.js + React.js)
          </p>
        </div>
        <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
          <button onClick={handleMockIncoming} className="btn-secondary" title="Simulate a user message event">
            <Users size={16} /> Simulate Teams User Chat
          </button>
          <button onClick={fetchData} className="btn-secondary">
            <RefreshCw size={16} /> Refresh
          </button>
        </div>
      </header>

      {/* Grid Layout */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: '24px' }}>
        
        {/* Left Column: Credentials & Setup */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          {/* Card 1: Auth & Token Generation */}
          <div className="glass-panel" style={{ padding: '24px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '16px' }}>
              <Key size={20} color="var(--accent-primary)" />
              <h2 style={{ fontSize: '1.15rem', fontWeight: '600' }}>Bot Framework OAuth2 Authentication</h2>
            </div>
            <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginBottom: '16px' }}>
              Bot Connector API endpoint: <code>https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token</code>
            </p>

            <form onSubmit={handleTestToken} style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '6px' }}>
                  MICROSOFT APP ID (Client ID)
                </label>
                <input 
                  type="text" 
                  className="input-field"
                  placeholder="e.g. 5e18b82e-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                  value={appId}
                  onChange={(e) => setAppId(e.target.value)}
                />
              </div>

              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '6px' }}>
                  MICROSOFT APP PASSWORD (Client Secret)
                </label>
                <input 
                  type="password" 
                  className="input-field"
                  placeholder="Bot Framework client secret password"
                  value={appPassword}
                  onChange={(e) => setAppPassword(e.target.value)}
                />
              </div>

              <button type="submit" className="btn-primary" disabled={loadingToken} style={{ alignSelf: 'flex-start' }}>
                <ShieldCheck size={16} />
                {loadingToken ? 'Verifying Token...' : 'Test & Generate Token'}
              </button>

              {tokenStatus && (
                <div style={{ 
                  marginTop: '10px', 
                  padding: '12px', 
                  borderRadius: '8px', 
                  fontSize: '0.85rem',
                  background: tokenStatus.success ? 'rgba(16, 185, 129, 0.1)' : 'rgba(244, 63, 94, 0.1)',
                  border: `1px solid ${tokenStatus.success ? 'rgba(16, 185, 129, 0.3)' : 'rgba(244, 63, 94, 0.3)'}`,
                  color: tokenStatus.success ? '#34d399' : '#fb7185'
                }}>
                  {tokenStatus.message}
                </div>
              )}
            </form>
          </div>



        </div>

        {/* Right Column: Send System Alert / Proactive Message */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          <div className="glass-panel" style={{ padding: '24px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '16px' }}>
              <Send size={20} color="var(--accent-emerald)" />
              <h2 style={{ fontSize: '1.15rem', fontWeight: '600' }}>Send Proactive System Message</h2>
            </div>
            <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginBottom: '18px' }}>
              Sends an automated system message or rich card directly to Teams without waiting for the user to chat first.
            </p>

            <form onSubmit={handleSendMessage} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
              
              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '6px' }}>
                  Target Teams Conversation
                </label>
                <select 
                  className="input-field" 
                  value={selectedConv}
                  onChange={(e) => setSelectedConv(e.target.value)}
                >
                  {conversations.map(c => (
                    <option key={c.conversationId} value={c.conversationId}>
                      {c.user?.name || 'User'} ({c.conversationId.substring(0, 24)}...)
                    </option>
                  ))}
                  <option value="custom">+ Enter Custom Conversation ID</option>
                </select>
              </div>

              {selectedConv === 'custom' && (
                <>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '6px' }}>
                      Service URL
                    </label>
                    <input 
                      type="text" 
                      className="input-field"
                      value={customServiceUrl}
                      onChange={(e) => setCustomServiceUrl(e.target.value)}
                      placeholder="https://smba.trafficmanager.net/teams"
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '6px' }}>
                      Conversation ID
                    </label>
                    <input 
                      type="text" 
                      className="input-field"
                      value={customConvId}
                      onChange={(e) => setCustomConvId(e.target.value)}
                      placeholder="19:xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx_xxxxxxxx@unq.gbl.spaces"
                    />
                  </div>
                </>
              )}

              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '6px' }}>
                  Payload Format
                </label>
                <div style={{ display: 'flex', gap: '16px' }}>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '0.88rem' }}>
                    <input 
                      type="radio" 
                      name="msgType" 
                      value="text" 
                      checked={messageType === 'text'} 
                      onChange={() => setMessageType('text')} 
                    />
                    Plain Text Message
                  </label>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '0.88rem' }}>
                    <input 
                      type="radio" 
                      name="msgType" 
                      value="card" 
                      checked={messageType === 'card'} 
                      onChange={() => setMessageType('card')} 
                    />
                    Interactive Card with Action Buttons
                  </label>
                </div>
              </div>

              {messageType === 'card' && (
                <div>
                  <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '6px' }}>
                    Card Title
                  </label>
                  <input 
                    type="text" 
                    className="input-field"
                    value={cardTitle}
                    onChange={(e) => setCardTitle(e.target.value)}
                  />
                </div>
              )}

              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '6px' }}>
                  Message Content / System Alert
                </label>
                <textarea 
                  rows={3} 
                  className="input-field"
                  value={messageText}
                  onChange={(e) => setMessageText(e.target.value)}
                  style={{ resize: 'vertical' }}
                />
              </div>

              <button type="submit" className="btn-primary" disabled={sendingMsg} style={{ alignSelf: 'flex-start' }}>
                <Send size={16} />
                {sendingMsg ? 'Posting to Bot Connector API...' : 'Send Message Now'}
              </button>

              {sendResult && (
                <div style={{ 
                  padding: '12px', 
                  borderRadius: '8px', 
                  fontSize: '0.85rem',
                  background: sendResult.success ? 'rgba(16, 185, 129, 0.1)' : 'rgba(244, 63, 94, 0.1)',
                  border: `1px solid ${sendResult.success ? 'rgba(16, 185, 129, 0.3)' : 'rgba(244, 63, 94, 0.3)'}`,
                  color: sendResult.success ? '#34d399' : '#fb7185'
                }}>
                  {sendResult.success ? sendResult.message : `Error: ${sendResult.error}`}
                </div>
              )}
            </form>
          </div>

        </div>

      </div>

      {/* Bottom Section: Real-time Activity Logs */}
      <div className="glass-panel" style={{ marginTop: '32px', padding: '24px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <Terminal size={20} color="var(--accent-amber)" />
            <h2 style={{ fontSize: '1.15rem', fontWeight: '600' }}>Live Bot Activity & Request Logs</h2>
          </div>
          <span className="badge badge-info">{logs.length} Total Logs</span>
        </div>

        <div style={{ 
          background: '#070a12', 
          borderRadius: '10px', 
          border: '1px solid var(--border-color)', 
          padding: '14px', 
          maxHeight: '300px', 
          overflowY: 'auto' 
        }}>
          {logs.length === 0 ? (
            <div style={{ color: 'var(--text-dim)', textAlign: 'center', padding: '20px' }}>
              No activities yet. Connect to Teams or click "Simulate Teams User Chat" to test.
            </div>
          ) : (
            logs.map(log => (
              <div key={log.id} style={{ 
                padding: '8px 0', 
                borderBottom: '1px solid rgba(255, 255, 255, 0.05)',
                display: 'flex',
                gap: '12px',
                fontSize: '0.82rem',
                fontFamily: 'JetBrains Mono, monospace'
              }}>
                <span style={{ color: 'var(--text-dim)' }}>
                  {new Date(log.timestamp).toLocaleTimeString()}
                </span>
                <span className={`badge ${
                  log.type === 'ERROR' ? 'badge-danger' : 
                  log.type === 'AUTH' ? 'badge-warning' : 
                  log.type === 'PROACTIVE' ? 'badge-success' : 'badge-info'
                }`}>
                  {log.type}
                </span>
                <span style={{ color: 'var(--text-main)', flex: 1 }}>
                  {log.message}
                </span>
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
}
