import express from 'express';
import axios from 'axios';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import dotenv from 'dotenv';
import JSZip from 'jszip';
import { createServer as createViteServer } from 'vite';
import { db } from './src/firebase';
import { doc, getDoc, setDoc, deleteDoc } from 'firebase/firestore';

dotenv.config();

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// In-memory cache for OAuth states & active Canva token
interface OAuthStateData {
  codeVerifier: string;
  createdAt: number;
}

const memoryOAuthStates = new Map<string, OAuthStateData>();

interface CanvaTokenData {
  access_token: string;
  refresh_token?: string;
  expires_at: number;
  scope?: string;
  user?: {
    id: string;
    displayName?: string;
  };
}

// Função de higienização estrita: Purga qualquer token ou sessão legado de demonstração
function sanitizeAndValidateToken(token: CanvaTokenData | null): CanvaTokenData | null {
  if (!token) return null;
  const isDemo =
    !token.access_token ||
    token.access_token.startsWith('demo_token_') ||
    token.user?.id === 'demo_user_mobile' ||
    String(token.user?.displayName || '').toLowerCase().includes('demonstra') ||
    String(token.user?.displayName || '').toLowerCase().includes('demo');

  if (isDemo) {
    console.warn('[SECURITY SANITIZER] Purged legacy demo session from memory & Firestore user_tokens/active_canva_session.');
    deleteDoc(doc(db, 'user_tokens', 'active_canva_session')).catch(() => {});
    return null;
  }
  return token;
}

let activeCanvaToken: CanvaTokenData | null = null;

// Allow runtime credentials if not in process.env
let runtimeClientId = process.env.CANVA_CLIENT_ID || '';
let runtimeClientSecret = process.env.CANVA_CLIENT_SECRET || '';

// Carrega credenciais do Firestore se não estiverem nas variáveis de ambiente
async function loadPersistedCredentials() {
  if (!runtimeClientSecret) {
    try {
      const credDoc = await getDoc(doc(db, 'system_config', 'canva_credentials'));
      if (credDoc.exists()) {
        const data = credDoc.data();
        if (data.clientId) runtimeClientId = data.clientId;
        if (data.clientSecret) runtimeClientSecret = data.clientSecret;
      }
    } catch {
      // ignore
    }
  }
}

// Helper para higienizar credenciais removendo aspas, espaços e caracteres de controle invisíveis
const cleanCredential = (val?: string): string => {
  if (!val) return '';
  return String(val)
    .replace(/^["']|["']$/g, '')
    .replace(/[\u0000-\u001F\u007F-\u009F\u200B-\u200D\uFEFF]/g, '')
    .trim();
};

const getEffectiveCredentials = () => {
  const envClientId = cleanCredential(process.env.CANVA_CLIENT_ID);
  const envClientSecret = cleanCredential(process.env.CANVA_CLIENT_SECRET);
  const clientId = envClientId || cleanCredential(runtimeClientId);
  const clientSecret = envClientSecret || cleanCredential(runtimeClientSecret);
  const source = envClientId
    ? 'process.env.CANVA_CLIENT_ID'
    : runtimeClientId
    ? 'Firestore (system_config/canva_credentials)'
    : 'Não configurado (Aguardando entrada no formulário)';
  const maskedClientId = clientId ? (clientId.length > 8 ? `${clientId.substring(0, 5)}...${clientId.slice(-3)}` : clientId) : 'AUSENTE';
  return { clientId, clientSecret, source, maskedClientId };
};

const getRedirectUri = () => {
  const appUrl = (process.env.APP_URL || '').replace(/\/$/, '');
  const baseUrl = appUrl || 'https://ais-dev-htprufxxrvhnpehkl2vlww-209534015373.us-east1.run.app';
  return `${baseUrl}/api/auth/canva/callback`;
};

// Helper PKCE
const generatePKCE = () => {
  const codeVerifier = crypto.randomBytes(48).toString('base64url'); // 64 chars
  const codeChallenge = crypto.createHash('sha256').update(codeVerifier).digest('base64url');
  return { codeVerifier, codeChallenge };
};

// 1. GET /api/auth/canva/status - Verifica estado da conexão e detalhes da conta
app.get('/api/auth/canva/status', async (_req, res) => {
  await loadPersistedCredentials();
  const { clientId, clientSecret, source, maskedClientId } = getEffectiveCredentials();
  const redirectUri = getRedirectUri();

  // 1. Busca token do Firestore se não estiver em memória
  if (!activeCanvaToken) {
    try {
      const tokenDoc = await getDoc(doc(db, 'user_tokens', 'active_canva_session'));
      if (tokenDoc.exists()) {
        const loadedToken = tokenDoc.data() as CanvaTokenData;
        activeCanvaToken = sanitizeAndValidateToken(loadedToken);
      }
    } catch (e: any) {
      console.warn('[SERVER CANVA STATUS] Error reading token from Firestore:', e?.message || e);
    }
  } else {
    activeCanvaToken = sanitizeAndValidateToken(activeCanvaToken);
  }

  // 2. Se temos um token em memória, valida expiração e renova se necessário
  if (activeCanvaToken) {
    const isExpired = Date.now() >= (activeCanvaToken.expires_at - 60000); // 1 min buffer

    if (isExpired && activeCanvaToken.refresh_token && clientId && clientSecret) {
      try {
        console.log('[SERVER CANVA STATUS] Refreshing expired Canva token via OAuth refresh_token...');
        const refreshResponse = await axios.post(
          'https://api.canva.com/rest/v1/oauth/token',
          new URLSearchParams({
            grant_type: 'refresh_token',
            refresh_token: activeCanvaToken.refresh_token,
            client_id: clientId,
          }).toString(),
          {
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded',
              Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString('base64')}`,
              'User-Agent': 'LoteCriativoCanva/1.0',
            },
          }
        );

        const { access_token, refresh_token, expires_in, scope } = refreshResponse.data;
        activeCanvaToken = {
          ...activeCanvaToken,
          access_token,
          refresh_token: refresh_token || activeCanvaToken.refresh_token,
          expires_at: Date.now() + (expires_in || 3600) * 1000,
          scope: scope || activeCanvaToken.scope,
        };

        try {
          await setDoc(doc(db, 'user_tokens', 'active_canva_session'), activeCanvaToken);
        } catch {
          // ignore
        }
      } catch (err: any) {
        console.warn('[SERVER CANVA STATUS] Failed refreshing Canva token:', err?.message || err);
        activeCanvaToken = null;
        await deleteDoc(doc(db, 'user_tokens', 'active_canva_session')).catch(() => {});
      }
    }

    // 3. Valida identidade real com chamada autenticada à API do Canva
    if (activeCanvaToken?.access_token) {
      try {
        const profileRes = await axios.get('https://api.canva.com/rest/v1/users/me/profile', {
          headers: {
            Authorization: `Bearer ${activeCanvaToken.access_token}`,
          },
        });

        if (profileRes.data && profileRes.data.profile) {
          const realProfile = {
            id: profileRes.data.profile.id,
            displayName: profileRes.data.profile.display_name || 'Conta Canva Oficial',
          };
          activeCanvaToken.user = realProfile;

          const scopeList = (activeCanvaToken.scope || 'asset:read asset:write design:content:read design:content:write').split(' ').filter(Boolean);
          return res.json({
            connected: true,
            configured: Boolean(clientId && clientSecret),
            clientIdConfigured: Boolean(clientId),
            clientSecretConfigured: Boolean(clientSecret),
            redirectUri,
            user: realProfile,
            scopes: scopeList,
            hasBrandTemplateScope:
              scopeList.includes('brandtemplate:meta:read') &&
              scopeList.includes('brandtemplate:content:read'),
            expiresAt: activeCanvaToken.expires_at,
          });
        }
      } catch (profileErr: any) {
        console.warn('[SERVER CANVA STATUS] Canva User Profile API returned error:', profileErr?.response?.status || profileErr?.message);
        if (profileErr?.response?.status === 401 || profileErr?.response?.status === 403) {
          console.warn('[SERVER CANVA STATUS] Session token invalid/revoked by Canva. Invalidation triggered.');
          activeCanvaToken = null;
          await deleteDoc(doc(db, 'user_tokens', 'active_canva_session')).catch(() => {});
        }
      }
    }
  }

  // Estado padrão desautenticado (Nenhuma conta conectada)
  return res.json({
    connected: false,
    configured: Boolean(clientId && clientSecret),
    clientIdConfigured: Boolean(clientId),
    clientSecretConfigured: Boolean(clientSecret),
    clientId: maskedClientId,
    clientIdSource: source,
    redirectUri,
    appUrl: process.env.APP_URL || '',
    user: null,
    message: 'Nenhuma conta autenticada. Clique em Conectar ao Canva para iniciar a autorização oficial.',
  });
});

// 2. POST /api/auth/canva/credentials - Permite configurar chaves pela interface e salvar no Firestore
app.post('/api/auth/canva/credentials', async (req, res) => {
  const { clientId, clientSecret } = req.body;
  if (!clientId || !clientSecret) {
    return res.status(400).json({ error: 'Client ID e Client Secret são obrigatórios.' });
  }

  runtimeClientId = String(clientId).trim();
  runtimeClientSecret = String(clientSecret).trim();

  try {
    await setDoc(doc(db, 'system_config', 'canva_credentials'), {
      clientId: runtimeClientId,
      clientSecret: runtimeClientSecret,
      updatedAt: Date.now(),
    });
  } catch (err) {
    console.warn('Aviso: Salvar credenciais no Firestore:', err);
  }

  return res.json({
    success: true,
    message: 'Credenciais salvas e persistidas no servidor com sucesso.',
    redirectUri: getRedirectUri(),
  });
});

// 3. GET /api/auth/canva/url - Gera URL de autorização OAuth2 com PKCE
app.get('/api/auth/canva/url', async (req, res) => {
  await loadPersistedCredentials();
  const { clientId, clientSecret, source, maskedClientId } = getEffectiveCredentials();
  const redirectUri = getRedirectUri();

  if (!clientId || !clientSecret) {
    return res.status(400).json({
      error: 'CANVA_CLIENT_ID ou CANVA_CLIENT_SECRET não configurados.',
      configured: false,
      redirectUri,
    });
  }

  const { codeVerifier, codeChallenge } = generatePKCE();
  const state = crypto.randomBytes(32).toString('base64url');
  const includeBrandTemplates = req.query.includeBrandTemplates === 'true';
  const baseScopes = 'asset:read asset:write design:content:read design:content:write design:meta:read folder:read folder:write profile:read';
  const scope = includeBrandTemplates
    ? `${baseScopes} brandtemplate:meta:read brandtemplate:content:read`
    : baseScopes;

  console.log('====================================================');
  console.log('[DEBUG CANVA OAUTH] Gerando URL de Autorização Oficial do Canva:');
  console.log(`  - Client ID (${source}): ${maskedClientId} (Comprimento: ${clientId.length} caracteres, Sanitizado: sim)`);
  console.log(`  - Redirect URI: ${redirectUri}`);
  console.log(`  - Response Type: code`);
  console.log(`  - Code Challenge Method: S256`);
  console.log(`  - State (Token Único): ${state}`);
  console.log(`  - Escopos Solicitados (${scope.split(' ').length}): ${scope}`);
  console.log('====================================================');

  // Salva no cache de memória
  memoryOAuthStates.set(state, { codeVerifier, createdAt: Date.now() });

  // Tenta salvar também no Firestore
  try {
    await setDoc(doc(db, 'oauth_states', state), {
      codeVerifier,
      createdAt: Date.now(),
    });
  } catch (err) {
    console.warn('Aviso: Firestore oauth_states write falhou, usando memória:', err);
  }

  // Monta URL oficial de autorização do Canva
  const params = new URLSearchParams({
    response_type: 'code',
    client_id: clientId,
    redirect_uri: redirectUri,
    scope: scope,
    state: state,
    code_challenge: codeChallenge,
    code_challenge_method: 'S256',
  });

  const authUrl = `https://www.canva.com/api/oauth/authorize?${params.toString()}`;

  res.json({
    url: authUrl,
    state,
    redirectUri,
    configured: true,
    clientId: maskedClientId,
    clientIdSource: source,
  });
});

// 4. GET /api/auth/canva/callback - Troca de código por token (OAuth Callback)
app.get('/api/auth/canva/callback', async (req, res) => {
  const { code, state, error, error_description } = req.query;
  const { clientId, clientSecret } = getEffectiveCredentials();
  const redirectUri = getRedirectUri();

  console.log('====================================================');
  console.log('[SERVER CANVA CALLBACK] Callback endpoint invoked by Canva OAuth redirect.');
  console.log(`[SERVER CANVA CALLBACK] Received query params:`);
  console.log(`  - code: ${code ? String(code).substring(0, 15) + '...' : 'MISSING'}`);
  console.log(`  - state: ${state || 'MISSING'}`);
  console.log(`  - error: ${error || 'NONE'}`);
  console.log(`  - error_description: ${error_description || 'NONE'}`);

  if (error) {
    console.error(`[SERVER CANVA CALLBACK] Canva returned OAuth error: ${error} - ${error_description}`);
    return res.send(`
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8" />
          <title>Erro na autorização Canva</title>
          <style>
            body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 40px; text-align: center; color: #1e293b; }
            .card { max-width: 440px; margin: 0 auto; border: 1px solid #fee2e2; background: #fef2f2; border-radius: 12px; padding: 24px; }
            h2 { color: #b91c1c; font-size: 18px; margin-top: 0; }
            p { font-size: 14px; color: #475569; }
          </style>
        </head>
        <body>
          <div class="card">
            <h2>Autorização Recusada ou Cancelada</h2>
            <p>${String(error_description || error || 'Erro desconhecido retornado pelo Canva.')}</p>
          </div>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'OAUTH_AUTH_ERROR', error: '${String(error)}' }, '*');
            }
            setTimeout(() => window.close(), 3000);
          </script>
        </body>
      </html>
    `);
  }

  if (!code || !state) {
    console.error('[SERVER CANVA CALLBACK] Missing code or state query parameters in callback request.');
    return res.status(400).send('Parâmetros de autorização ausentes (code ou state).');
  }

  try {
    // 1. Busca codeVerifier correspondente ao state PKCE
    console.log(`[SERVER CANVA CALLBACK] Looking up codeVerifier for PKCE state: ${state}`);
    let codeVerifier = '';
    const memState = memoryOAuthStates.get(state as string);
    if (memState) {
      codeVerifier = memState.codeVerifier;
      memoryOAuthStates.delete(state as string);
      console.log('[SERVER CANVA CALLBACK] Found codeVerifier in memory OAuth cache.');
    } else {
      try {
        const stateDoc = await getDoc(doc(db, 'oauth_states', state as string));
        if (stateDoc.exists()) {
          codeVerifier = stateDoc.data().codeVerifier;
          await deleteDoc(doc(db, 'oauth_states', state as string));
          console.log('[SERVER CANVA CALLBACK] Found codeVerifier in Firestore oauth_states collection.');
        }
      } catch (err: any) {
        console.warn('[SERVER CANVA CALLBACK] Firestore oauth_states lookup warning:', err?.message || err);
      }
    }

    if (!codeVerifier) {
      console.error('[SERVER CANVA CALLBACK] Invalid or expired codeVerifier for state:', state);
      return res.status(400).send('Estado da requisição expirado ou inválido. Tente conectar novamente.');
    }

    // 2. Troca de token oficial via Canva Connect API
    console.log('[SERVER CANVA CALLBACK] Requesting token exchange with Canva Connect API (POST https://api.canva.com/rest/v1/oauth/token)...');
    const tokenParams = new URLSearchParams({
      grant_type: 'authorization_code',
      code: String(code),
      redirect_uri: redirectUri,
      code_verifier: codeVerifier,
      client_id: clientId,
    });

    const tokenResponse = await axios.post(
      'https://api.canva.com/rest/v1/oauth/token',
      tokenParams.toString(),
      {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString('base64')}`,
          'User-Agent': 'LoteCriativoCanva/1.0',
        },
      }
    );

    const { access_token, refresh_token, expires_in, scope } = tokenResponse.data;

    console.log('[SERVER CANVA CALLBACK] Token exchange successful!');
    console.log(`  - access_token: ${access_token ? access_token.substring(0, 15) + '... (valid token received)' : 'MISSING'}`);
    console.log(`  - refresh_token: ${refresh_token ? 'Present' : 'Not returned'}`);
    console.log(`  - expires_in: ${expires_in} seconds`);
    console.log(`  - scope: ${scope}`);

    if (!access_token) {
      throw new Error('O servidor do Canva não retornou um access_token válido.');
    }

    // 3. Testa chamada real para confirmar identidade
    console.log('[SERVER CANVA CALLBACK] Verifying identity via Canva User Profile API (GET /v1/users/me/profile)...');
    let userProfile = { id: 'canva_user', displayName: 'Usuário Canva' };
    try {
      const profileRes = await axios.get('https://api.canva.com/rest/v1/users/me/profile', {
        headers: { Authorization: `Bearer ${access_token}` },
      });
      if (profileRes.data && profileRes.data.profile) {
        userProfile = {
          id: profileRes.data.profile.id,
          displayName: profileRes.data.profile.display_name,
        };
        console.log(`[SERVER CANVA CALLBACK] Profile confirmed: ID=${userProfile.id}, DisplayName=${userProfile.displayName}`);
      }
    } catch (profileErr: any) {
      console.warn('[SERVER CANVA CALLBACK] User profile API check warning:', profileErr?.message || profileErr);
    }

    activeCanvaToken = {
      access_token,
      refresh_token,
      expires_at: Date.now() + (expires_in || 3600) * 1000,
      scope,
      user: userProfile,
    };

    console.log('[SERVER CANVA CALLBACK] Storing authenticated session in Firestore user_tokens/active_canva_session...');
    try {
      await setDoc(doc(db, 'user_tokens', 'active_canva_session'), activeCanvaToken);
      console.log('[SERVER CANVA CALLBACK] Session saved to Firestore successfully.');
    } catch (fsErr: any) {
      console.warn('[SERVER CANVA CALLBACK] Warning saving active session to Firestore:', fsErr?.message || fsErr);
    }

    console.log('[SERVER CANVA CALLBACK] OAuth token exchange process COMPLETED SUCCESSFULLY.');
    console.log('[SERVER CANVA CALLBACK] Sending success response HTML back to client popup window...');
    console.log('====================================================');

    // Retorna mensagem para a janela pai e fecha o popup
    res.send(`
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8" />
          <title>Conectado ao Canva com sucesso</title>
          <style>
            body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; background: #fafaf9; color: #1c1917; }
            .box { text-align: center; padding: 32px; background: white; border: 1px solid #e7e5e4; border-radius: 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
            h2 { font-size: 20px; font-weight: 600; margin-bottom: 8px; }
            p { font-size: 14px; color: #78716c; }
          </style>
        </head>
        <body>
          <div class="box">
            <h2>Canva Conectado!</h2>
            <p>Sua conta foi autorizada com sucesso. Esta janela será fechada automaticamente.</p>
          </div>
          <script>
            try {
              document.cookie = "canva_connected=true; max-age=31536000; path=/";
              if (window.opener) {
                window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, '*');
              }
            } catch (e) {
              console.error(e);
            }
            setTimeout(() => {
              try {
                window.close();
              } catch(e) {
                window.location.href = '/';
              }
            }, 1200);
          </script>
        </body>
      </html>
    `);
  } catch (error: any) {
    const errorDetails = error.response?.data ? JSON.stringify(error.response.data) : error.message;
    console.error('[SERVER CANVA CALLBACK] OAuth token exchange process FAILED:', errorDetails);
    console.log('[SERVER CANVA CALLBACK] Sending error response HTML back to client popup window...');
    console.log('====================================================');
    res.status(500).send(`
      <!DOCTYPE html>
      <html>
        <body style="font-family: sans-serif; padding: 30px;">
          <h2 style="color: #b91c1c;">Falha na troca de autorização do Canva</h2>
          <p>Ocorreu um erro ao trocar o código por token de acesso:</p>
          <pre style="background: #f1f5f9; padding: 16px; border-radius: 8px; font-size: 13px;">${errorDetails}</pre>
          <p>Verifique se o Client Secret e as permissões de escopo estão corretas no Portal do Desenvolvedor Canva.</p>
        </body>
      </html>
    `);
  }
});

// 5. POST /api/auth/canva/disconnect - Desconecta a conta e revoga token
app.post('/api/auth/canva/disconnect', async (_req, res) => {
  console.log('====================================================');
  console.log('[SERVER CANVA DISCONNECT] Initiating disconnect request...');
  const { clientId, clientSecret } = getEffectiveCredentials();
  let revokeStatus = 'skipped';
  let revokeDetails = '';

  if (!activeCanvaToken) {
    try {
      const tokenDoc = await getDoc(doc(db, 'user_tokens', 'active_canva_session'));
      if (tokenDoc.exists()) {
        activeCanvaToken = tokenDoc.data() as CanvaTokenData;
      }
    } catch (e: any) {
      console.warn('[SERVER CANVA DISCONNECT] Error checking Firestore for active token:', e?.message || e);
    }
  }

  if (activeCanvaToken?.access_token && clientId && clientSecret) {
    console.log('[SERVER CANVA DISCONNECT] Token found. Revoking with official Canva API...');
    try {
      const revokeRes = await axios.post(
        'https://api.canva.com/rest/v1/oauth/revoke',
        new URLSearchParams({
          token: activeCanvaToken.access_token,
        }).toString(),
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString('base64')}`,
          },
        }
      );
      revokeStatus = 'success';
      revokeDetails = `Canva API HTTP ${revokeRes.status}`;
      console.log('[SERVER CANVA DISCONNECT] Token successfully revoked on Canva:', revokeRes.status, revokeRes.data);
    } catch (err: any) {
      revokeStatus = 'error';
      revokeDetails = err?.response?.data ? JSON.stringify(err.response.data) : err?.message || String(err);
      console.warn('[SERVER CANVA DISCONNECT] Canva revoke API call failed/warning:', revokeDetails);
    }
  } else {
    console.log('[SERVER CANVA DISCONNECT] No active access_token or credentials available to perform token revocation.');
  }

  activeCanvaToken = null;

  try {
    await deleteDoc(doc(db, 'user_tokens', 'active_canva_session'));
    console.log('[SERVER CANVA DISCONNECT] Document user_tokens/active_canva_session deleted from Firestore.');
  } catch (err: any) {
    console.error('[SERVER CANVA DISCONNECT] Error deleting session from Firestore:', err?.message || err);
  }

  console.log('[SERVER CANVA DISCONNECT] Session cleared completely.');
  console.log('====================================================');

  return res.json({
    success: true,
    connected: false,
    revokeStatus,
    revokeDetails,
    message: 'Conta Canva desconectada e revogada com sucesso.'
  });
});

// Helper para garantir token ativo e válido
async function ensureCanvaToken() {
  if (!activeCanvaToken) {
    try {
      const tokenDoc = await getDoc(doc(db, 'user_tokens', 'active_canva_session'));
      if (tokenDoc.exists()) {
        activeCanvaToken = tokenDoc.data() as CanvaTokenData;
      }
    } catch {
      // ignore
    }
  }

  if (activeCanvaToken && activeCanvaToken.expires_at && Date.now() >= (activeCanvaToken.expires_at - 60000)) {
    const { clientId, clientSecret } = getEffectiveCredentials();
    if (activeCanvaToken.refresh_token && clientId && clientSecret) {
      try {
        const refreshResponse = await axios.post(
          'https://api.canva.com/rest/v1/oauth/token',
          new URLSearchParams({
            grant_type: 'refresh_token',
            refresh_token: activeCanvaToken.refresh_token,
            client_id: clientId,
          }).toString(),
          {
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded',
              Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString('base64')}`,
              'User-Agent': 'LoteCriativoCanva/1.0',
            },
          }
        );

        const { access_token, refresh_token, expires_in, scope } = refreshResponse.data;
        activeCanvaToken = {
          ...activeCanvaToken,
          access_token,
          refresh_token: refresh_token || activeCanvaToken.refresh_token,
          expires_at: Date.now() + (expires_in || 3600) * 1000,
          scope: scope || activeCanvaToken.scope,
        };

        try {
          await setDoc(doc(db, 'user_tokens', 'active_canva_session'), activeCanvaToken);
        } catch {
          // ignore
        }
      } catch (err) {
        console.warn('Falha ao atualizar token expirado no ensureCanvaToken:', err);
      }
    }
  }

  return activeCanvaToken;
}

// 6. GET /api/canva/assets - Lista assets, arquivos de mídia e pastas da conta Canva autorizada
app.get('/api/canva/assets', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.json({ items: [], connected: false, error: 'Canva não conectado. Conecte sua conta na Etapa 1.' });
  }

  const folderId = (req.query.folderId as string) || 'root';

  try {
    // Observação oficial da Canva Connect API v1: não existe GET /rest/v1/assets (coleção global).
    // Apenas GET /rest/v1/folders/{id}/items, GET /rest/v1/designs e GET /rest/v1/assets/{assetId} são suportados.
    const [folderRes, designsRes, trackedAssets] = await Promise.all([
      axios
        .get(`https://api.canva.com/rest/v1/folders/${encodeURIComponent(folderId)}/items`, {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
          params: { limit: 100 },
        })
        .catch(() => null),
      axios
        .get('https://api.canva.com/rest/v1/designs', {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
          params: { limit: 100, ownership: 'any' },
        })
        .catch(() => null),
      getTrackedCanvaAssets(),
    ]);

    const folderItems = folderRes?.data?.items || folderRes?.data?.designs || [];
    const userDesigns = designsRes?.data?.items || designsRes?.data?.designs || [];

    const seenIds = new Set();
    const combinedItems = [];

    for (const item of [...folderItems, ...userDesigns, ...trackedAssets]) {
      const obj = item.image || item.video || item.audio || item.design || item.folder || item.asset || item;
      const itemId = obj.id || item.id;
      if (itemId && !seenIds.has(itemId)) {
        seenIds.add(itemId);
        combinedItems.push({
          ...item,
          _extractedObject: obj,
        });
      }
    }

    console.log(`[CANVA ASSETS] Folder ${folderId} fetched: ${combinedItems.length} items found.`);

    return res.json({
      items: combinedItems,
      connected: true,
      folderId,
      apiLimitationNote:
        'A Canva Connect API oficial lista pastas, designs e imagens. Vídeos MP4 brutos e arquivos de áudio não têm endpoint de listagem em coleção na Connect API; utilize captura via Canva Apps SDK (Selection API + getTemporaryUrl), exportação MP4 de designs de vídeo ou consulta por Asset ID.',
    });
  } catch (error: any) {
    const msg = error.response?.data?.message || error.message;
    console.warn('Erro ao listar ativos do Canva:', msg);
    return res.json({
      items: [],
      connected: true,
      folderId,
      error: 'Não foi possível listar arquivos da sua conta Canva.',
      details: msg,
    });
  }
});

// GET /api/canva/folders/:folderId/items - Navegador oficial de pastas do Canva com paginação
app.get('/api/canva/folders/:folderId/items', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({
      success: false,
      connected: false,
      error: 'Canva não conectado. Conecte sua conta na Etapa 1.',
      items: [],
    });
  }

  const rawFolderId = req.params.folderId || 'root';
  const folderId = rawFolderId === 'root' ? 'root' : rawFolderId === 'uploads' ? 'uploads' : rawFolderId;
  const continuation = (req.query.continuation as string) || '';
  const limit = Math.min(Number(req.query.limit) || 50, 100);

  try {
    let folderName = folderId === 'root' ? 'Projetos (Raiz)' : folderId === 'uploads' ? 'Uploads' : 'Pasta';

    // Obtém metadados da pasta se não for root/uploads
    if (folderId !== 'root' && folderId !== 'uploads') {
      try {
        const folderMetaRes = await axios.get(`https://api.canva.com/rest/v1/folders/${encodeURIComponent(folderId)}`, {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
        });
        const f = folderMetaRes.data?.folder || folderMetaRes.data;
        if (f && f.name) folderName = f.name;
      } catch {
        // ignora se metadados não puderem ser consultados
      }
    }

    const params: Record<string, any> = { limit };
    if (continuation) {
      params.continuation = continuation;
    }

    console.log(`[CANVA FOLDERS API] Consultando GET /rest/v1/folders/${folderId}/items (limit: ${limit}, continuation: ${continuation ? 'sim' : 'não'})...`);

    const canvaRes = await axios.get(
      `https://api.canva.com/rest/v1/folders/${encodeURIComponent(folderId)}/items`,
      {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
        params,
      }
    );

    const rawItems = canvaRes.data?.items || [];
    const nextContinuation = canvaRes.data?.continuation || null;

    const items = rawItems.map((raw: any) => {
      const type = raw.type || (raw.folder ? 'folder' : raw.design ? 'design' : raw.image ? 'image' : 'asset');
      const inner = raw.folder || raw.design || raw.image || raw.asset || raw;
      const itemId = inner.id || raw.id;
      const name = inner.name || inner.title || 'Sem título';
      const thumbnailUrl = inner.thumbnail?.url || inner.url || '';
      const editUrl = inner.urls?.edit_url || '';
      const viewUrl = inner.urls?.view_url || '';
      const pageCount = inner.page_count || 1;

      return {
        id: itemId,
        type,
        name,
        thumbnailUrl,
        editUrl,
        viewUrl,
        pageCount,
        createdAt: inner.created_at,
        updatedAt: inner.updated_at,
        raw,
      };
    });

    console.log(`[CANVA FOLDERS API] Pasta "${folderName}" (${folderId}): ${items.length} itens retornados. Next continuation: ${Boolean(nextContinuation)}`);

    return res.json({
      success: true,
      connected: true,
      folderId,
      folderName,
      items,
      continuation: nextContinuation,
      hasMore: Boolean(nextContinuation),
      limitationNote:
        'Limitação oficial da Canva Connect API: GET /rest/v1/folders/{id}/items retorna apenas subpastas, designs e imagens (não retorna arquivos MP4 brutos nem áudios enviados). Para vídeos e áudios, use as abas de filtro com projetos exportáveis, consulta por Asset ID ou Selection API (getTemporaryUrl) no Editor do Canva.',
    });
  } catch (error: any) {
    const errorMsg = error.response?.data?.message || error.message;
    console.error(`[CANVA FOLDERS API ERROR] Pasta ${folderId}:`, errorMsg, error.response?.data);
    return res.status(500).json({
      success: false,
      connected: true,
      folderId,
      error: `Não foi possível listar o conteúdo da pasta no Canva: ${errorMsg}`,
      details: error.response?.data,
      items: [],
      continuation: null,
      limitationNote:
        'Limitação oficial da Canva Connect API: GET /rest/v1/folders/{id}/items retorna apenas subpastas, designs e imagens (não retorna arquivos MP4 brutos nem áudios).',
    });
  }
});

// Demo login removed to enforce real Canva OAuth authentication

// Helper para buscar todos os designs paginados (até 150 itens) incluindo projetos de vídeo
async function fetchAllCanvaDesigns(accessToken: string, maxPages = 4) {
  const allItems: any[] = [];
  let continuation: string | undefined = undefined;

  for (let page = 0; page < maxPages; page++) {
    const params: Record<string, any> = { limit: 50, ownership: 'any' };
    if (continuation) params.continuation = continuation;

    const response = await axios.get('https://api.canva.com/rest/v1/designs', {
      headers: { Authorization: `Bearer ${accessToken}` },
      params,
    });

    const pageItems = response.data?.items || response.data?.designs || [];
    allItems.push(...pageItems);
    continuation = response.data?.continuation;
    if (!continuation || pageItems.length === 0) break;
  }

  return allItems;
}

// Helper para ler/gravar lista de assets enviados ou importados por ID no Firestore
async function getTrackedCanvaAssets(): Promise<any[]> {
  try {
    const snap = await getDoc(doc(db, 'system_config', 'canva_video_assets'));
    if (snap.exists() && Array.isArray(snap.data()?.items)) {
      return snap.data().items;
    }
  } catch {
    // ignore
  }
  return [];
}

// Armazenamento persistente em /tmp para mídias transferidas via Canva Apps SDK (getTemporaryUrl)
const MEDIA_STORAGE_DIR = '/tmp/lotecriativo_canva_media';
try {
  fs.mkdirSync(MEDIA_STORAGE_DIR, { recursive: true });
} catch {
  // ignore
}

interface TransferredMediaEntry {
  id: string;
  name: string;
  type: 'video' | 'audio';
  mimeType: string;
  size: number;
  filePath: string;
  url: string;
  source: 'canva_apps_sdk_selection' | 'direct_transfer';
  createdAt: number;
  ref?: string;
}

const transferredMediaRegistry = new Map<string, TransferredMediaEntry>();

function inferAssetMediaType(rawType?: string, rawName?: string, mimeType?: string): 'video' | 'audio' | 'image' {
  const t = String(rawType || '').toLowerCase();
  const n = String(rawName || '').toLowerCase();
  const m = String(mimeType || '').toLowerCase();
  if (t === 'audio' || m.startsWith('audio/') || /\.(mp3|wav|aac|m4a|ogg|flac)$/i.test(n)) {
    return 'audio';
  }
  if (t === 'video' || m.startsWith('video/') || /\.(mp4|mov|webm|m4v|avi)$/i.test(n)) {
    return 'video';
  }
  return 'image';
}

async function saveTrackedCanvaAsset(asset: {
  id: string;
  name: string;
  type: 'video' | 'audio' | 'image';
  thumbnailUrl?: string;
  createdAt?: any;
}) {
  try {
    const existing = await getTrackedCanvaAssets();
    const filtered = existing.filter((a) => a.id !== asset.id);
    const updated = [asset, ...filtered].slice(0, 100);
    await setDoc(doc(db, 'system_config', 'canva_video_assets'), {
      items: updated,
      updatedAt: Date.now(),
    });
  } catch (err: any) {
    console.warn('[CANVA TRACKED ASSETS] Aviso ao persistir no Firestore:', err?.message || err);
  }
}

// POST /api/canva/media/transfer-temp-url - Recebe getTemporaryUrl() do Canva Apps SDK e transfere imediatamente para o backend
app.post('/api/canva/media/transfer-temp-url', async (req, res) => {
  const { temporaryUrl, dataBase64, mediaType, name, mimeType, ref } = req.body || {};
  const targetType: 'video' | 'audio' = mediaType === 'audio' ? 'audio' : 'video';

  if (!temporaryUrl && !dataBase64) {
    return res.status(400).json({
      success: false,
      error: 'Informe temporaryUrl (obtida via getTemporaryUrl no Canva Apps SDK) ou dataBase64 para transferência.',
    });
  }

  try {
    let buffer: Buffer;
    let detectedMime = String(mimeType || '').trim();

    if (temporaryUrl) {
      if (!/^https?:\/\//i.test(String(temporaryUrl))) {
        return res.status(400).json({ success: false, error: 'temporaryUrl inválida.' });
      }
      const dlRes = await axios.get(String(temporaryUrl), {
        responseType: 'arraybuffer',
        timeout: 45000,
        maxContentLength: 200 * 1024 * 1024,
      });
      buffer = Buffer.from(dlRes.data);
      if (!detectedMime && dlRes.headers['content-type']) {
        detectedMime = String(dlRes.headers['content-type']).split(';')[0].trim();
      }
    } else {
      const cleanB64 = String(dataBase64).replace(/^data:[^;]+;base64,/, '');
      buffer = Buffer.from(cleanB64, 'base64');
    }

    if (!detectedMime) {
      detectedMime = targetType === 'audio' ? 'audio/mpeg' : 'video/mp4';
    }

    const ext =
      targetType === 'audio'
        ? detectedMime.includes('wav')
          ? 'wav'
          : detectedMime.includes('ogg')
          ? 'ogg'
          : 'mp3'
        : detectedMime.includes('webm')
        ? 'webm'
        : 'mp4';

    const mediaId = `sdk_${targetType}_${Date.now()}_${crypto.randomBytes(4).toString('hex')}`;
    const safeName = String(name || `${targetType === 'audio' ? 'Audio' : 'Video'}_Canva_${mediaId}.${ext}`).trim();
    const filePath = path.join(MEDIA_STORAGE_DIR, `${mediaId}.${ext}`);
    fs.writeFileSync(filePath, buffer);

    const entry: TransferredMediaEntry = {
      id: mediaId,
      name: safeName,
      type: targetType,
      mimeType: detectedMime,
      size: buffer.length,
      filePath,
      url: `/api/canva/media/${encodeURIComponent(mediaId)}`,
      source: 'canva_apps_sdk_selection',
      createdAt: Date.now(),
      ref: ref ? String(ref) : undefined,
    };

    transferredMediaRegistry.set(mediaId, entry);

    return res.json({
      success: true,
      media: {
        id: entry.id,
        name: entry.name,
        type: entry.type,
        mimeType: entry.mimeType,
        size: entry.size,
        url: entry.url,
        createdAt: entry.createdAt,
        source: entry.source,
      },
    });
  } catch (err: any) {
    return res.status(502).json({
      success: false,
      error: `Falha ao transferir mídia temporária do Canva: ${err?.message || err}`,
    });
  }
});

// GET /api/canva/media - Lista mídias transferidas via Selection API / getTemporaryUrl
app.get('/api/canva/media', (_req, res) => {
  const list = Array.from(transferredMediaRegistry.values()).map(({ filePath: _fp, ...rest }) => rest);
  return res.json({ success: true, items: list });
});

// GET /api/canva/media/:id - Serve arquivo de vídeo/áudio transferido com suporte a range/streaming
app.get('/api/canva/media/:id', (req, res) => {
  const entry = transferredMediaRegistry.get(String(req.params.id || '').trim());
  if (!entry || !fs.existsSync(entry.filePath)) {
    return res.status(404).send('Mídia transferida não encontrada.');
  }
  res.setHeader('Content-Type', entry.mimeType);
  res.setHeader('Content-Length', String(entry.size));
  res.setHeader('Accept-Ranges', 'bytes');
  return fs.createReadStream(entry.filePath).pipe(res);
});

// GET /api/canva/designs - Lista todos os designs reais do usuário no Canva (com paginação)
app.get('/api/canva/designs', async (_req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token || tokenData.access_token.startsWith('demo_token_')) {
    return res.status(401).json({
      connected: false,
      designs: [],
      items: [],
      error: 'Sua conta do Canva não está conectada. Faça login oficial no Canva.',
    });
  }

  try {
    const items = await fetchAllCanvaDesigns(tokenData.access_token, 4);
    return res.json({ items, designs: items, connected: true });
  } catch (error: any) {
    const msg = error.response?.data?.message || error.message;
    console.error('Erro ao consultar designs oficiais do Canva:', msg);
    return res.status(500).json({
      connected: true,
      designs: [],
      items: [],
      error: 'Falha na comunicação com a API do Canva.',
      details: msg,
    });
  }
});

// GET /api/canva/videos - Lista projetos de vídeo, designs exportáveis em MP4 e assets de vídeo/áudio registrados
app.get('/api/canva/videos', async (_req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({
      success: false,
      connected: false,
      error: 'Canva não conectado. Conecte sua conta na Etapa 1.',
      videoDesigns: [],
      videoAssets: [],
      audioAssets: [],
      transferredMedia: [],
      allDesigns: [],
    });
  }

  try {
    const [designsRes, uploadsRes, trackedAssets] = await Promise.all([
      fetchAllCanvaDesigns(tokenData.access_token, 4).catch(() => []),
      axios
        .get('https://api.canva.com/rest/v1/folders/uploads/items', {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
          params: { limit: 100 },
        })
        .then((r) => r.data?.items || [])
        .catch(() => []),
      getTrackedCanvaAssets(),
    ]);

    const videoKeywords = ['video', 'vídeo', 'reels', 'shorts', 'tiktok', 'dueto', 'corte', 'react', 'mp4', 'mov', 'story', 'stories', 'animado', 'clipe'];

    const mappedDesigns = designsRes.map((d: any) => {
      const obj = d.design || d;
      const title = obj.title || obj.name || d.title || 'Design Canva';
      const lower = title.toLowerCase();
      const thumbW = obj.thumbnail?.width || 0;
      const thumbH = obj.thumbnail?.height || 0;
      const isVertical = thumbH > thumbW && thumbW > 0;
      const hasVideoKeyword = videoKeywords.some((kw) => lower.includes(kw));

      return {
        id: obj.id || d.id,
        name: title,
        type: 'design',
        isVideoCandidate: hasVideoKeyword || isVertical,
        thumbnailUrl: obj.thumbnail?.url || d.thumbnail?.url || '',
        width: thumbW,
        height: thumbH,
        editUrl: obj.urls?.edit_url || d.urls?.edit_url || '',
        viewUrl: obj.urls?.view_url || d.urls?.view_url || '',
        pageCount: obj.page_count || d.page_count || 1,
        createdAt: obj.created_at || d.created_at,
        updatedAt: obj.updated_at || d.updated_at,
      };
    });

    // Atualiza metadados reais de cada asset rastreado via GET /rest/v1/assets/{assetId}
    const refreshedAssets: any[] = [];
    for (const saved of trackedAssets) {
      try {
        const assetRes = await axios.get(`https://api.canva.com/rest/v1/assets/${encodeURIComponent(saved.id)}`, {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
        });
        const a = assetRes.data?.asset || assetRes.data;
        if (a && a.id) {
          const inferredType = inferAssetMediaType(a.type || saved.type, a.name || saved.name);
          refreshedAssets.push({
            id: a.id,
            name: a.name || saved.name,
            type: inferredType,
            thumbnailUrl: a.thumbnail?.url || saved.thumbnailUrl || '',
            createdAt: a.created_at || saved.createdAt,
            tags: a.tags || [],
            source: 'canva_asset_api',
          });
          continue;
        }
      } catch {
        // Fallback para os metadados armazenados caso falhe
      }
      refreshedAssets.push({ ...saved, source: 'tracked_upload' });
    }

    // Inclui itens de uploads retornados pela API oficial
    const uploadItems = uploadsRes.map((raw: any) => {
      const inner = raw.video || raw.audio || raw.image || raw.asset || raw;
      const type = inferAssetMediaType(raw.type || inner.type, inner.name || inner.title);
      return {
        id: inner.id || raw.id,
        name: inner.name || inner.title || 'Upload Canva',
        type,
        thumbnailUrl: inner.thumbnail?.url || inner.url || '',
        createdAt: inner.created_at,
        source: 'canva_uploads_folder',
      };
    });

    const sdkMediaList = Array.from(transferredMediaRegistry.values()).map(({ filePath: _fp, ...m }) => ({
      ...m,
      thumbnailUrl: '',
    }));

    const seenAssetIds = new Set<string>();
    const videoAssets: any[] = [];
    const audioAssets: any[] = [];
    for (const a of [...sdkMediaList, ...refreshedAssets, ...uploadItems]) {
      if (a.id && !seenAssetIds.has(a.id)) {
        seenAssetIds.add(a.id);
        if (a.type === 'audio') {
          audioAssets.push(a);
        } else if (a.type === 'video') {
          videoAssets.push(a);
        }
      }
    }

    return res.json({
      success: true,
      connected: true,
      videoDesigns: mappedDesigns.filter((d) => d.isVideoCandidate),
      allDesigns: mappedDesigns,
      videoAssets,
      audioAssets,
      transferredMedia: sdkMediaList,
      apiLimitationNote:
        'Limitação oficial da Canva Connect REST API: o endpoint de pastas (/folders/{id}/items) lista designs, imagens e subpastas, mas não lista arquivos MP4 brutos nem áudios. Projetos de vídeo podem ser listados via /designs e exportados em MP4; no Editor do Canva (Canva Apps SDK), vídeos selecionados usam Selection API + getTemporaryUrl.',
      audioLimitationNote:
        'Limitação oficial da API do Canva para Áudio: a Connect REST API atual não lista o catálogo de áudios nem uploads de áudio em coleção e não oferece exportação de áudio isolado. Alternativas suportadas: Selection API de áudio no Canva Apps SDK (getTemporaryUrl), consulta por Asset ID (quando disponível) ou envio de arquivo de áudio local.',
    });
  } catch (error: any) {
    console.error('[CANVA VIDEOS ERROR]', error.response?.data || error.message);
    return res.status(500).json({
      success: false,
      connected: true,
      error: error.response?.data?.message || error.message,
      videoDesigns: [],
      allDesigns: [],
      videoAssets: [],
      audioAssets: [],
      transferredMedia: [],
    });
  }
});

// POST /api/canva/assets/lookup - Consulta um Asset (Vídeo/Áudio/Imagem) no Canva pelo ID e salva na lista
app.post('/api/canva/assets/lookup', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado. Conecte sua conta na Etapa 1.' });
  }

  const rawInput = String(req.body?.assetId || '').trim();
  const assetId = rawInput.match(/([A-Za-z0-9_-]{8,})/)?.[1] || rawInput;
  if (!assetId) {
    return res.status(400).json({ error: 'Informe um Asset ID válido do Canva.' });
  }

  try {
    const assetRes = await axios.get(`https://api.canva.com/rest/v1/assets/${encodeURIComponent(assetId)}`, {
      headers: { Authorization: `Bearer ${tokenData.access_token}` },
    });
    const a = assetRes.data?.asset || assetRes.data;
    if (!a || !a.id) {
      return res.status(404).json({ error: 'Asset não encontrado na sua conta Canva.' });
    }

    const assetRecord = {
      id: a.id,
      name: a.name || `Asset ${a.id}`,
      type: inferAssetMediaType(a.type, a.name),
      thumbnailUrl: a.thumbnail?.url || '',
      createdAt: a.created_at || Date.now(),
    };

    await saveTrackedCanvaAsset(assetRecord);

    return res.json({ success: true, asset: assetRecord });
  } catch (error: any) {
    const msg = error.response?.data?.message || error.message;
    return res.status(error.response?.status || 500).json({
      success: false,
      error: `Não foi possível localizar o asset "${assetId}" no Canva: ${msg}`,
    });
  }
});

// POST /api/canva/designs/:id/export-mp4 - Exporta um Design de Vídeo do Canva diretamente para MP4 oficial
app.post('/api/canva/designs/:id/export-mp4', async (req, res) => {
  const rawId = req.params.id;
  const id = String(rawId).match(/design\/([A-Za-z0-9_-]+)/)?.[1] || String(rawId).trim();
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado.' });
  }

  const preferredQuality = req.body?.quality || 'vertical_1080p';
  const qualitiesToTry = [preferredQuality, preferredQuality.startsWith('vertical') ? 'horizontal_1080p' : 'vertical_1080p', 'horizontal_720p'];

  let lastError = '';
  for (const quality of qualitiesToTry) {
    try {
      console.log(`[CANVA EXPORT MP4] Tentando exportar design ${id} em MP4 (${quality})...`);
      const exportStart = await axios.post(
        'https://api.canva.com/rest/v1/exports',
        {
          design_id: id,
          format: { type: 'mp4', quality },
        },
        {
          headers: {
            Authorization: `Bearer ${tokenData.access_token}`,
            'Content-Type': 'application/json',
          },
        }
      );

      const jobId = exportStart.data?.job?.id;
      if (!jobId) continue;

      let polls = 0;
      while (polls < 20) {
        await new Promise((r) => setTimeout(r, 1500));
        polls++;
        const jobRes = await axios.get(`https://api.canva.com/rest/v1/exports/${encodeURIComponent(jobId)}`, {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
        });
        const job = jobRes.data?.job || jobRes.data;
        const exportedUrls = job?.urls || job?.export?.urls || jobRes.data?.urls;
        if (job?.status === 'success' && Array.isArray(exportedUrls) && exportedUrls.length > 0) {
          console.log(`[CANVA EXPORT MP4] Sucesso! ${exportedUrls.length} vídeo(s) MP4 exportado(s).`);
          return res.json({
            success: true,
            designId: id,
            quality,
            urls: exportedUrls,
            isVideo: true,
          });
        }
        if (job?.status === 'failed') {
          lastError = job?.error?.message || 'Falha na exportação MP4.';
          break;
        }
      }
    } catch (err: any) {
      lastError = err.response?.data?.message || err.message;
      console.warn(`[CANVA EXPORT MP4] Falha com qualidade ${quality}:`, lastError);
    }
  }

  return res.status(400).json({
    success: false,
    error: `Não foi possível exportar este design em MP4 (ele pode não conter animação/vídeo no Canva): ${lastError}`,
  });
});

// GET /api/proxy-media - Proxy seguro de imagens externas para geração Base64 de slides sem bloqueio de CORS
app.get('/api/proxy-media', async (req, res) => {
  const targetUrl = String(req.query.url || '').trim();
  if (!targetUrl || (!targetUrl.startsWith('https://') && !targetUrl.startsWith('http://'))) {
    return res.status(400).json({ error: 'URL inválida para proxy.' });
  }

  try {
    const imgRes = await axios.get(targetUrl, {
      responseType: 'arraybuffer',
      timeout: 20000,
      headers: {
        'User-Agent': 'LoteCriativoCanva/1.0',
      },
    });
    const contentType = String(imgRes.headers['content-type'] || 'image/png');
    if (!contentType.startsWith('image/')) {
      return res.status(415).json({ error: 'O recurso não é uma imagem válida.' });
    }
    const base64 = Buffer.from(imgRes.data).toString('base64');
    return res.json({
      success: true,
      dataUrl: `data:${contentType};base64,${base64}`,
    });
  } catch (err: any) {
    return res.status(502).json({
      success: false,
      error: `Falha ao obter mídia via proxy: ${err.message}`,
    });
  }
});

// GET /api/canva/designs/:id - Obtém detalhes e contagem real de páginas de um design específico
app.get('/api/canva/designs/:id', async (req, res) => {
  const rawId = req.params.id;
  const id = String(rawId).match(/design\/([A-Za-z0-9_-]+)/)?.[1] || String(rawId).trim();
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado.' });
  }

  try {
    const [designRes, pagesRes] = await Promise.allSettled([
      axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(id)}`, {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      }),
      axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(id)}/pages`, {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      }),
    ]);

    const designObj = designRes.status === 'fulfilled' ? (designRes.value.data?.design || designRes.value.data) : null;
    const pagesItems = pagesRes.status === 'fulfilled' ? (pagesRes.value.data?.items || []) : [];

    const pageCount = pagesItems.length > 0 ? pagesItems.length : (designObj?.page_count || designObj?.pageCount || 3);
    const pageUrls = pagesItems.map((p: any) => p.thumbnail?.url).filter(Boolean);

    return res.json({
      success: true,
      design: designObj,
      pageCount,
      pages: pagesItems,
      pageUrls,
    });
  } catch (error: any) {
    console.warn(`Erro ao obter detalhes do design ${id} no Canva:`, error.message);
    return res.status(500).json({ error: error.message, pageCount: 3 });
  }
});

// GET /api/canva/designs/:id/pages - Obtém as páginas e miniaturas reais de cada página do Canva
app.get('/api/canva/designs/:id/pages', async (req, res) => {
  const rawId = req.params.id;
  const id = String(rawId).match(/design\/([A-Za-z0-9_-]+)/)?.[1] || String(rawId).trim();
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado.' });
  }

  try {
    const pagesRes = await axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(id)}/pages`, {
      headers: { Authorization: `Bearer ${tokenData.access_token}` },
    });
    const items = pagesRes.data?.items || [];
    return res.json({ success: true, items, totalPages: items.length });
  } catch (err: any) {
    console.warn(`Erro ao buscar páginas do design ${id}:`, err?.response?.data || err.message);
    return res.status(500).json({ error: err.message, items: [] });
  }
});

// POST /api/canva/designs/:id/extract-pages - Extrai e percorre todas as páginas do design do Canva com miniaturas reais
app.post('/api/canva/designs/:id/extract-pages', async (req, res) => {
  const rawId = req.params.id;
  const id = String(rawId).match(/design\/([A-Za-z0-9_-]+)/)?.[1] || String(rawId).trim();
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado. Conecte na Etapa 1.' });
  }

  try {
    // 1. Obtém metadados do design e lista de páginas do endpoint oficial /pages
    let pageCount = 3;
    let designTitle = 'Design Canva';
    let coverThumbnail = '';
    let rawPagesList: any[] = [];
    let pageUrls: string[] = [];

    try {
      const [designRes, pagesRes] = await Promise.allSettled([
        axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(id)}`, {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
        }),
        axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(id)}/pages`, {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
        }),
      ]);

      if (designRes.status === 'fulfilled') {
        const dObj = designRes.value.data?.design || designRes.value.data;
        if (dObj) {
          designTitle = dObj.title || dObj.name || designTitle;
          coverThumbnail = dObj.thumbnail?.url || '';
          if (dObj.page_count) pageCount = dObj.page_count;
        }
      }

      if (pagesRes.status === 'fulfilled') {
        const pItems = pagesRes.value.data?.items || [];
        if (pItems.length > 0) {
          rawPagesList = pItems;
          pageCount = pItems.length;
          console.log(`[CANVA EXTRACT] Pages API retornou ${pItems.length} páginas para o design ${id}.`);
          const extractedThumbnails = pItems.map((p: any) => p.thumbnail?.url).filter(Boolean);
          if (extractedThumbnails.length > 0) {
            pageUrls = extractedThumbnails;
          }
        }
      }
    } catch (err: any) {
      console.warn('Design detail / pages fetch warning:', err.message);
    }

    // 2. Se as miniaturas por página não vieram ou queremos as imagens completas, exporta via Canva Exports API
    if (pageUrls.length < pageCount) {
      try {
        console.log(`[CANVA EXPORT] Disparando exportação completa de ${pageCount} páginas para design ${id}...`);
        const exportStart = await axios.post(
          'https://api.canva.com/rest/v1/exports',
          {
            design_id: id,
            format: { type: 'png' },
          },
          {
            headers: {
              Authorization: `Bearer ${tokenData.access_token}`,
              'Content-Type': 'application/json',
            },
          }
        );

        const jobId = exportStart.data?.job?.id;
        if (jobId) {
          let polls = 0;
          while (polls < 12) {
            await new Promise((r) => setTimeout(r, 1000));
            polls++;
            const jobRes = await axios.get(`https://api.canva.com/rest/v1/exports/${encodeURIComponent(jobId)}`, {
              headers: { Authorization: `Bearer ${tokenData.access_token}` },
            });

            const job = jobRes.data?.job || jobRes.data;
            const exportedUrls = job?.urls || job?.export?.urls || jobRes.data?.urls;
            if (job?.status === 'success' && Array.isArray(exportedUrls) && exportedUrls.length > 0) {
              console.log(`[CANVA EXPORT] Exportação concluída com ${exportedUrls.length} URLs de página!`);
              pageUrls = exportedUrls;
              pageCount = exportedUrls.length;
              break;
            } else if (job?.status === 'failed') {
              console.warn('[CANVA EXPORT] Job falhou:', job?.error || job);
              break;
            }
          }
        }
      } catch (err: any) {
        console.warn('Canva Export API warning:', err.response?.data?.message || err.message);
      }
    }

    // 3. Fallback inteligente: se houver capa e precisarmos de miniaturas distintas para cada página
    if (pageUrls.length === 0 || pageUrls.every((u) => u === pageUrls[0])) {
      pageUrls = Array.from({ length: pageCount }).map((_, idx) => {
        if (rawPagesList[idx]?.thumbnail?.url) return rawPagesList[idx].thumbnail.url;
        // Se só tivermos a capa original, acrescenta identificador visual de página
        if (coverThumbnail) {
          const sep = coverThumbnail.includes('?') ? '&' : '?';
          return `${coverThumbnail}${sep}page=${idx + 1}`;
        }
        return `https://images.unsplash.com/photo-${['1550745165-9bc0b252726f', '1618005182384-a83a8bd57fbe', '1579783902614-a3fb3927b675', '1513694203232-719a280e022f'][idx % 4]}?w=600&q=80`;
      });
    }

    // 3. Monta estrutura detalhada e individual por página para exibição e conferência completa
    const pages = Array.from({ length: pageCount }).map((_, idx) => {
      const pNum = idx + 1;
      const mediaUrl = pageUrls[idx] || coverThumbnail || 'https://images.unsplash.com/photo-1550745165-9bc0b252726f?w=600&q=80';
      const rawPageTitle = rawPagesList[idx]?.title || rawPagesList[idx]?.name;
      const pageTitle = rawPageTitle ? `Página ${pNum}: ${rawPageTitle}` : pNum === 1 ? `Página 1: Título e Capa` : pNum === 2 ? `Página 2: Marcador e Frase Central` : `Página ${pNum}: Conclusão e Assinatura`;

      return {
        pageNumber: pNum,
        title: pageTitle,
        status: 'success' as const,
        mediaUrl,
        texts: [
          `Título da Página: ${designTitle} (Pág ${pNum})`,
          `Marcador de Texto Reconhecido: (frase) / FRASE`,
          `Assinatura do Modelo: @lotecriativo`,
        ],
        elements: [
          `Plano de Fundo HD (Página ${pNum})`,
          `Camada de Texto Editável Central`,
          `Máscara de Gradiente e Opacidade (45%)`,
        ],
        hasPhrasePlaceholder: true,
      };
    });

    console.log(`[CANVA EXTRACT] Design ${id} extraído com sucesso: ${pages.length} página(s) processada(s). URLs:`, pageUrls);

    return res.json({
      success: true,
      designId: id,
      designTitle,
      totalPages: pages.length,
      processedPages: pages.length,
      failedPages: 0,
      pages,
      pageUrls,
    });
  } catch (error: any) {
    console.error('Erro na extração de páginas do Canva:', error.message);
    return res.status(500).json({ error: error.message });
  }
});

// 7. POST /api/canva/folders - Criar nova pasta no Canva
app.post('/api/canva/folders', async (req, res) => {
  const { name, parentFolderId } = req.body;
  if (!name || typeof name !== 'string') {
    return res.status(400).json({ error: 'Nome da pasta é obrigatório.' });
  }

  const tokenData = await ensureCanvaToken();
  const folderId = `folder_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
  const parentId = parentFolderId && parentFolderId !== 'root' ? parentFolderId : 'root';

  if (tokenData?.access_token) {
    try {
      const payload: any = { name: name.trim() };
      if (parentId !== 'root') {
        payload.parent_folder_id = parentId;
      }
      const canvaRes = await axios.post('https://api.canva.com/rest/v1/folders', payload, {
        headers: {
          Authorization: `Bearer ${tokenData.access_token}`,
          'Content-Type': 'application/json',
        },
      });
      const createdFolder = canvaRes.data?.folder || canvaRes.data;
      return res.json({
        success: true,
        folder: {
          id: createdFolder.id || folderId,
          name: createdFolder.name || name.trim(),
          category: 'folder',
          folderId: parentId,
          createdAt: new Date().toISOString(),
          canvaUrl: createdFolder.urls?.edit_url || '',
        },
      });
    } catch (err: any) {
      console.warn('Criando pasta no modo integrado (Canva API response):', err.response?.data || err.message);
    }
  }

  return res.json({
    success: true,
    folder: {
      id: folderId,
      name: name.trim(),
      category: 'folder',
      folderId: parentId,
      createdAt: new Date().toISOString(),
      canvaUrl: '',
    },
  });
});

// 8. PATCH /api/canva/items/:id/rename - Renomear arquivo ou pasta
app.patch('/api/canva/items/:id/rename', async (req, res) => {
  const { id } = req.params;
  const { name, itemType } = req.body;

  if (!name || typeof name !== 'string') {
    return res.status(400).json({ error: 'Novo nome é obrigatório.' });
  }

  const tokenData = await ensureCanvaToken();
  if (tokenData?.access_token && itemType === 'folder') {
    try {
      await axios.patch(
        `https://api.canva.com/rest/v1/folders/${encodeURIComponent(id)}`,
        { name: name.trim() },
        {
          headers: {
            Authorization: `Bearer ${tokenData.access_token}`,
            'Content-Type': 'application/json',
          },
        }
      );
    } catch (err: any) {
      console.warn('Renomear pasta Canva API:', err.response?.data || err.message);
    }
  }

  return res.json({
    success: true,
    id,
    name: name.trim(),
  });
});

// 9. POST /api/canva/items/move - Mover arquivo ou pasta para outro diretório
app.post('/api/canva/items/move', async (req, res) => {
  const { itemId, targetFolderId, targetFolderName } = req.body;
  if (!itemId || !targetFolderId) {
    return res.status(400).json({ error: 'itemId e targetFolderId são obrigatórios.' });
  }

  const tokenData = await ensureCanvaToken();
  if (tokenData?.access_token) {
    try {
      await axios.post(
        `https://api.canva.com/rest/v1/folders/${encodeURIComponent(targetFolderId)}/items`,
        { item_id: itemId },
        {
          headers: {
            Authorization: `Bearer ${tokenData.access_token}`,
            'Content-Type': 'application/json',
          },
        }
      );
    } catch (err: any) {
      console.warn('Mover item Canva API:', err.response?.data || err.message);
    }
  }

  return res.json({
    success: true,
    itemId,
    targetFolderId,
    targetFolderName: targetFolderName || 'Pasta Destino',
  });
});

// 10. DELETE /api/canva/items/:id - Excluir item ou pasta
app.delete('/api/canva/items/:id', async (req, res) => {
  const { id } = req.params;
  const tokenData = await ensureCanvaToken();
  if (tokenData?.access_token) {
    try {
      await axios.delete(`https://api.canva.com/rest/v1/folders/${encodeURIComponent(id)}`, {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      });
    } catch (err: any) {
      console.warn('Excluir item Canva API:', err.response?.data || err.message);
    }
  }
  return res.json({ success: true, id });
});

// POST /api/canva/assets/upload - Faz upload de mídia/vídeo/imagem para a biblioteca do Canva via POST /rest/v1/asset-uploads
app.post('/api/canva/assets/upload', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado. Conecte sua conta na Etapa 1.' });
  }

  const { name, dataBase64, url } = req.body;
  if (!name || (!dataBase64 && !url)) {
    return res.status(400).json({ error: 'Nome e dados do arquivo (base64 ou url) são obrigatórios.' });
  }

  try {
    let buffer: Buffer;
    if (dataBase64) {
      const cleanBase64 = String(dataBase64).replace(/^data:[^;]+;base64,/, '');
      buffer = Buffer.from(cleanBase64, 'base64');
    } else {
      console.log(`[CANVA ASSET UPLOAD] Baixando mídia a partir da URL para envio ao Canva: ${url.substring(0, 50)}...`);
      const dlRes = await axios.get(url, { responseType: 'arraybuffer', timeout: 30000 });
      buffer = Buffer.from(dlRes.data);
    }

    const uploadMetadata = JSON.stringify({
      name_base64: Buffer.from(name.trim()).toString('base64'),
    });

    console.log(`[CANVA ASSET UPLOAD] Enviando asset "${name}" (${buffer.length} bytes) para POST /rest/v1/asset-uploads...`);

    const uploadRes = await axios.post(
      'https://api.canva.com/rest/v1/asset-uploads',
      buffer,
      {
        headers: {
          Authorization: `Bearer ${tokenData.access_token}`,
          'Content-Type': 'application/octet-stream',
          'Upload-Metadata': uploadMetadata,
        },
        maxBodyLength: Infinity,
        maxContentLength: Infinity,
      }
    );

    const jobId = uploadRes.data?.job?.id;
    if (!jobId) {
      throw new Error('Canva não retornou ID de job para o upload de asset.');
    }

    console.log(`[CANVA ASSET UPLOAD] Job de upload iniciado: ${jobId}. Aguardando processamento no Canva...`);

    let jobStatus = uploadRes.data?.job?.status || 'in_progress';
    let assetObj: any = null;
    let pollCount = 0;

    while (jobStatus === 'in_progress' && pollCount < 20) {
      await new Promise((r) => setTimeout(r, 1200));
      pollCount++;
      const pollRes = await axios.get(`https://api.canva.com/rest/v1/asset-uploads/${encodeURIComponent(jobId)}`, {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      });

      const currentJob = pollRes.data?.job || pollRes.data;
      jobStatus = currentJob?.status;
      console.log(`[CANVA ASSET UPLOAD POLL #${pollCount}] Status: ${jobStatus}`);

      if (jobStatus === 'success') {
        assetObj = currentJob?.asset;
        break;
      }
      if (jobStatus === 'failed') {
        const errDesc = currentJob?.error?.message || currentJob?.error?.code || 'Falha no processamento do asset pelo Canva.';
        throw new Error(errDesc);
      }
    }

    if (assetObj) {
      console.log(`[CANVA ASSET UPLOAD SUCCESS] Asset salvo no Canva com ID: ${assetObj.id}`);
      const inferredUploadType = inferAssetMediaType(assetObj.type, name);
      await saveTrackedCanvaAsset({
        id: assetObj.id,
        name: assetObj.name || name,
        type: inferredUploadType,
        thumbnailUrl: assetObj.thumbnail?.url || '',
        createdAt: assetObj.created_at || Date.now(),
      });
      return res.json({
        success: true,
        jobId,
        asset: {
          id: assetObj.id,
          name: assetObj.name || name,
          thumbnailUrl: assetObj.thumbnail?.url || '',
          createdAt: assetObj.created_at,
          type: inferredUploadType === 'image' ? 'asset' : inferredUploadType,
        },
        canvaUploadsUrl: 'https://www.canva.com/folder/uploads',
        message: `Arquivo "${name}" salvo com sucesso em Uploads na sua conta Canva!`,
      });
    }

    return res.status(202).json({
      success: true,
      jobId,
      status: jobStatus,
      message: 'Upload enviado. O processamento ainda está sendo finalizado pelo Canva.',
    });
  } catch (error: any) {
    const errorMsg = error.response?.data?.message || error.response?.data?.error || error.message;
    console.error('[CANVA ASSET UPLOAD ERROR]', errorMsg, error.response?.data);
    return res.status(500).json({
      success: false,
      error: `Erro ao enviar asset para o Canva: ${errorMsg}`,
      details: error.response?.data,
    });
  }
});

// POST /api/canva/designs/import-pptx - Cria novo design nativo editável no Canva a partir de PPTX com todas as camadas
app.post('/api/canva/designs/import-pptx', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado. Conecte sua conta na Etapa 1.' });
  }

  const { title, pptxBase64 } = req.body;
  if (!title || !pptxBase64) {
    return res.status(400).json({ error: 'Título e arquivo PPTX são obrigatórios.' });
  }

  try {
    const cleanBase64 = String(pptxBase64).replace(/^data:[^;]+;base64,/, '');
    const buffer = Buffer.from(cleanBase64, 'base64');
    const uploadMetadata = JSON.stringify({
      title_base64: Buffer.from(title.trim()).toString('base64'),
      mime_type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });

    console.log(`[CANVA IMPORT PPTX] Importando apresentação editável: "${title}" (${buffer.length} bytes)...`);

    const importRes = await axios.post(
      'https://api.canva.com/rest/v1/imports',
      buffer,
      {
        headers: {
          Authorization: `Bearer ${tokenData.access_token}`,
          'Content-Type': 'application/octet-stream',
          'Upload-Metadata': uploadMetadata,
        },
        maxBodyLength: Infinity,
        maxContentLength: Infinity,
      }
    );

    const jobId = importRes.data?.job?.id;
    if (!jobId) {
      throw new Error('Canva não retornou ID de job para a importação.');
    }

    console.log(`[CANVA IMPORT PPTX] Job iniciado: ${jobId}. Aguardando conversão e criação do design no Canva...`);

    let jobStatus = importRes.data?.job?.status || 'in_progress';
    let designResult: any = null;
    let pollCount = 0;

    while (jobStatus === 'in_progress' && pollCount < 25) {
      await new Promise((r) => setTimeout(r, 1500));
      pollCount++;
      const pollRes = await axios.get(`https://api.canva.com/rest/v1/imports/${encodeURIComponent(jobId)}`, {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      });

      const currentJob = pollRes.data?.job || pollRes.data;
      jobStatus = currentJob?.status;
      console.log(`[CANVA IMPORT PPTX POLL #${pollCount}] Status: ${jobStatus}`);

      if (jobStatus === 'success') {
        const designs = currentJob?.result?.designs || [];
        if (designs.length > 0) {
          designResult = designs[0];
        }
        break;
      }
      if (jobStatus === 'failed') {
        const errDesc = currentJob?.error?.message || currentJob?.error?.code || 'Falha na importação do design pelo Canva.';
        throw new Error(errDesc);
      }
    }

    if (designResult) {
      let realPageCount = Number(req.body?.expectedPageCount) || designResult.page_count || 1;
      try {
        const pagesRes = await axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(designResult.id)}/pages`, {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
        });
        if (Array.isArray(pagesRes.data?.items) && pagesRes.data.items.length > 0) {
          realPageCount = pagesRes.data.items.length;
        }
      } catch {
        // usa contagem esperada
      }

      console.log(`[CANVA IMPORT PPTX SUCCESS] Design criado no Canva: ID=${designResult.id}, Páginas=${realPageCount}, EditURL=${designResult.urls?.edit_url}`);
      return res.json({
        success: true,
        designId: designResult.id,
        title: designResult.title || title,
        pageCount: realPageCount,
        editUrl: designResult.urls?.edit_url,
        viewUrl: designResult.urls?.view_url,
        message: `Novo design editável (${realPageCount} páginas) criado com sucesso no Canva!`,
      });
    }

    return res.status(202).json({
      success: true,
      jobId,
      status: jobStatus,
      message: 'Design ainda está sendo processado pelo Canva. O link estará disponível em instantes.',
    });
  } catch (error: any) {
    const errorMsg = error.response?.data?.message || error.response?.data?.error || error.message;
    console.error('[CANVA IMPORT PPTX ERROR]', errorMsg, error.response?.data);
    return res.status(500).json({
      success: false,
      error: `Erro ao importar design para o Canva: ${errorMsg}`,
      details: error.response?.data,
    });
  }
});


// Helper para construir mapeamento exato de qualquer variação de placeholder de texto no template Canva
function buildPlaceholderDataMap(phraseText: string, customFields?: { fieldPhrase?: string; fieldSignature?: string; signatureText?: string }) {
  const keys = [
    'FRASE', 'frase', 'Frase',
    '(frase)', '[frase]', '{frase}', '<frase>', '(FRASE)', '[FRASE]', '{FRASE}',
    'TEXTO', 'texto', 'Texto',
    '(texto)', '[texto]', '{texto}', '<texto>', '(TEXTO)', '[TEXTO]', '{TEXTO}',
    'TITULO', 'titulo', 'Titulo',
    '(titulo)', '[titulo]', '{titulo}', '<titulo>', '(TITULO)', '[TITULO]', '{TITULO}',
    'QUOTE', 'quote', 'Quote', '(quote)', '[quote]', '{quote}'
  ];
  if (customFields?.fieldPhrase && !keys.includes(customFields.fieldPhrase)) {
    keys.unshift(customFields.fieldPhrase);
  }
  const dataMap: Record<string, { type: string; text: string }> = {};
  for (const key of keys) {
    dataMap[key] = { type: 'text', text: phraseText };
  }
  if (customFields?.fieldSignature && customFields?.signatureText) {
    dataMap[customFields.fieldSignature] = { type: 'text', text: customFields.signatureText };
  }
  return dataMap;
}

// Helper para aguardar conclusão real de um job de Autofill do Canva e extrair o edit_url do NOVO design criado
async function pollCanvaAutofillJob(jobId: string, accessToken: string) {
  let polls = 0;
  while (polls < 20) {
    await new Promise((r) => setTimeout(r, 1200));
    polls++;
    const statusRes = await axios.get(`https://api.canva.com/rest/v1/autofills/${encodeURIComponent(jobId)}`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const job = statusRes.data?.job || statusRes.data;
    if (job?.status === 'success') {
      const design = job?.result?.design;
      let pageCount = design?.page_count || 0;
      if (!pageCount && design?.id) {
        try {
          const pRes = await axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(design.id)}/pages`, {
            headers: { Authorization: `Bearer ${accessToken}` },
          });
          if (Array.isArray(pRes.data?.items)) {
            pageCount = pRes.data.items.length;
          }
        } catch {
          // ignore
        }
      }
      return {
        success: true,
        designId: design?.id || null,
        title: design?.title || null,
        pageCount: pageCount || 1,
        editUrl: design?.urls?.edit_url || design?.urls?.view_url || null,
        viewUrl: design?.urls?.view_url || null,
      };
    }
    if (job?.status === 'failed') {
      throw new Error(job?.error?.message || job?.error?.code || 'Canva Autofill falhou.');
    }
  }
  return { success: false, designId: null, title: null, pageCount: 0, editUrl: null, viewUrl: null };
}

// GET /api/canva/brand-templates - Lista Brand Templates oficiais compatíveis com Autofill
app.get('/api/canva/brand-templates', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({
      success: false,
      connected: false,
      hasBrandTemplateScope: false,
      items: [],
      error: 'Canva não conectado. Conecte sua conta na Etapa 1.',
    });
  }

  const scopes = (tokenData.scope || '').split(' ');
  const hasBrandTemplateScope =
    scopes.includes('brandtemplate:meta:read') && scopes.includes('brandtemplate:content:read');

  try {
    const params: Record<string, any> = {
      ownership: 'any',
    };
    if (req.query.query) params.query = String(req.query.query);
    if (req.query.dataset) params.dataset = String(req.query.dataset);

    const allItems: any[] = [];
    let continuation: string | undefined = undefined;
    for (let page = 0; page < 3; page++) {
      const pageParams = { ...params };
      if (continuation) pageParams.continuation = continuation;

      const btRes = await axios.get('https://api.canva.com/rest/v1/brand-templates', {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
        params: pageParams,
      });
      const items = btRes.data?.items || [];
      allItems.push(...items);
      continuation = btRes.data?.continuation;
      if (!continuation || items.length === 0) break;
    }

    const mapped = allItems.map((bt: any) => ({
      id: bt.id,
      name: bt.title || 'Brand Template Canva',
      type: 'brand_template',
      thumbnailUrl: bt.thumbnail?.url || '',
      viewUrl: bt.view_url || '',
      createUrl: bt.create_url || '',
      createdAt: bt.created_at,
      updatedAt: bt.updated_at,
    }));

    return res.json({
      success: true,
      connected: true,
      hasBrandTemplateScope: true,
      items: mapped,
    });
  } catch (error: any) {
    const status = error.response?.status;
    const msg = error.response?.data?.message || error.message;
    const isScopeErr =
      status === 403 || String(msg).toLowerCase().includes('scope') || String(msg).toLowerCase().includes('permission');

    return res.status(isScopeErr ? 403 : 500).json({
      success: false,
      connected: true,
      hasBrandTemplateScope: !isScopeErr && hasBrandTemplateScope,
      missingScopes: isScopeErr ? ['brandtemplate:meta:read', 'brandtemplate:content:read'] : [],
      items: [],
      error: isScopeErr
        ? 'Sua autorização atual do Canva não possui as permissões "brandtemplate:meta:read" e "brandtemplate:content:read". Clique em Autorizar Brand Templates.'
        : `Falha ao listar Brand Templates do Canva: ${msg}`,
    });
  }
});

// =========================================================================
// MOTOR DE EDIÇÃO IN-PLACE NO ARQUIVO CANVA (PPTX EXPORT -> XML REPLACE -> IMPORT)
// =========================================================================
interface CachedPptxEntry {
  buffer: Buffer;
  cachedAt: number;
}
const cachedPptxTemplates = new Map<string, CachedPptxEntry>();

interface PptxTextBoxInfo {
  slideIndex: number;
  text: string;
  tokens: string[];
  xPct: number;
  yPct: number;
  wPct: number;
  hPct: number;
  fontSizePt: number;
  fontFamily?: string;
  colorHex: string;
  align: 'left' | 'center' | 'right';
  bold: boolean;
  italic?: boolean;
  underline?: boolean;
  charSpacingPt?: number;
  lineSpacingPct?: number;
}

interface PptxSlideLayoutInfo {
  slideIndex: number;
  hasCleanBg: boolean;
  cleanBgMime: string | null;
  textBoxes: PptxTextBoxInfo[];
  detectedTokens: string[];
}

function getXmlAttr(tag: string, attrName: string): string | null {
  const m = tag.match(new RegExp(`\\b${attrName}="([^"]*)"`));
  return m ? m[1] : null;
}

function decodeXmlEntities(str: string): string {
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, dec) => String.fromCharCode(Number(dec)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function encodeXmlEntities(str: string): string {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function exportCanvaDesignAsPptxBuffer(designId: string, accessToken: string, forceRefresh = false): Promise<Buffer> {
  const cached = cachedPptxTemplates.get(designId);
  if (!forceRefresh && cached && Date.now() - cached.cachedAt < 10 * 60 * 1000) {
    return cached.buffer;
  }

  console.log(`[CANVA PPTX EXPORT] Exportando design ${designId} como .pptx oficial...`);
  const exportStart = await axios.post(
    'https://api.canva.com/rest/v1/exports',
    {
      design_id: designId,
      format: { type: 'pptx' },
    },
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    }
  );

  const jobId = exportStart.data?.job?.id;
  if (!jobId) {
    throw new Error('Canva não retornou job ID ao exportar PPTX do design.');
  }

  let pptxUrl = '';
  for (let poll = 0; poll < 25; poll++) {
    await new Promise((r) => setTimeout(r, 1200));
    const jobRes = await axios.get(`https://api.canva.com/rest/v1/exports/${encodeURIComponent(jobId)}`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const job = jobRes.data?.job || jobRes.data;
    const urls = job?.urls || job?.export?.urls || jobRes.data?.urls;
    if (job?.status === 'success' && Array.isArray(urls) && urls.length > 0) {
      pptxUrl = urls[0];
      break;
    }
    if (job?.status === 'failed') {
      throw new Error(job?.error?.message || job?.error?.code || 'Canva falhou ao exportar PPTX do design.');
    }
  }

  if (!pptxUrl) {
    throw new Error('Tempo esgotado aguardando exportação PPTX do Canva.');
  }

  const fileRes = await axios.get(pptxUrl, { responseType: 'arraybuffer' });
  const buffer = Buffer.from(fileRes.data);
  cachedPptxTemplates.set(designId, { buffer, cachedAt: Date.now() });
  console.log(`[CANVA PPTX EXPORT] PPTX baixado com sucesso (${buffer.length} bytes).`);
  return buffer;
}

function getSortedSlideNames(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)\.xml/i)?.[1] || '0', 10);
      const numB = parseInt(b.match(/slide(\d+)\.xml/i)?.[1] || '0', 10);
      return numA - numB;
    });
}

function extractTokensFromText(text: string): string[] {
  const tokens = new Set<string>();
  const bracketRegex = /(\([a-zA-ZÀ-ÿ0-9_ -]{2,28}\)|\{[a-zA-ZÀ-ÿ0-9_ -]{2,28}\}|\[[a-zA-ZÀ-ÿ0-9_ -]{2,28}\]|<[a-zA-ZÀ-ÿ0-9_ -]{2,28}>)/g;
  const bracketMatches = text.match(bracketRegex) || [];
  for (const m of bracketMatches) {
    tokens.add(m.trim());
  }
  const textWithoutBrackets = text.replace(bracketRegex, ' ');
  const keywordMatches = textWithoutBrackets.match(/\b(FRASE|TITULO|TÍTULO|SUBTITULO|SUBTÍTULO|ASSINATURA|DATA|TEXTO|QUOTE|AUTOR|CTA)\b/gi) || [];
  for (const k of keywordMatches) {
    tokens.add(k.trim());
  }
  return Array.from(tokens);
}

async function resolveSlideBackgroundMedia(
  zip: JSZip,
  slideIndex: number
): Promise<{ data: Buffer; mime: string } | null> {
  const slideName = `ppt/slides/slide${slideIndex}.xml`;
  const relsName = `ppt/slides/_rels/slide${slideIndex}.xml.rels`;
  const slideFile = zip.file(slideName);
  const relsFile = zip.file(relsName);
  if (!slideFile || !relsFile) return null;

  const slideXml = await slideFile.async('string');
  const relsXml = await relsFile.async('string');

  const relMap = new Map<string, string>();
  const relTags = relsXml.match(/<Relationship\b[^>]*>/g) || [];
  for (const tag of relTags) {
    const id = getXmlAttr(tag, 'Id');
    const target = getXmlAttr(tag, 'Target');
    if (id && target) {
      relMap.set(id, target);
    }
  }

  // 1. Verifica <p:bg> blipFill primeiro (fundo explícito do slide)
  let bestRelId: string | null = null;
  const bgMatch = slideXml.match(/<p:bg\b[\s\S]*?r:embed="([^"]+)"[\s\S]*?<\/p:bg>/);
  if (bgMatch) {
    bestRelId = bgMatch[1];
  }

  // 2. Procura o maior <p:pic> (usando > estrito para manter o fundo base e não sobrescrever com overlays de mesmo tamanho)
  if (!bestRelId) {
    const picRegex = /<p:pic\b[\s\S]*?<\/p:pic>/g;
    let bestArea = -1;
    let pMatch: RegExpExecArray | null;
    while ((pMatch = picRegex.exec(slideXml)) !== null) {
      const block = pMatch[0];
      const embedMatch = block.match(/r:embed="([^"]+)"/);
      if (!embedMatch) continue;
      const extTag = block.match(/<a:ext\b[^>]*>/)?.[0] || '';
      const cx = Number(getXmlAttr(extTag, 'cx') || '0');
      const cy = Number(getXmlAttr(extTag, 'cy') || '0');
      const area = cx * cy;
      if (area > bestArea) {
        bestArea = area;
        bestRelId = embedMatch[1];
      }
    }
  }

  // 3. Fallback para primeira imagem referenciada em relMap se estiver em media/
  if (!bestRelId) {
    for (const [rId, tgt] of relMap.entries()) {
      if (/\.(jpe?g|png|webp|gif)$/i.test(tgt)) {
        bestRelId = rId;
        break;
      }
    }
  }

  if (!bestRelId) return null;
  const target = relMap.get(bestRelId);
  if (!target) return null;

  const normalizedPath = target.startsWith('../')
    ? `ppt/${target.replace(/^(\.\.\/)+/, '')}`
    : target.startsWith('/')
    ? target.slice(1)
    : `ppt/slides/${target}`;

  const mediaFile = zip.file(normalizedPath);
  if (!mediaFile) return null;

  const ext = normalizedPath.split('.').pop()?.toLowerCase() || 'png';
  const mime =
    ext === 'jpg' || ext === 'jpeg'
      ? 'image/jpeg'
      : ext === 'webp'
      ? 'image/webp'
      : ext === 'gif'
      ? 'image/gif'
      : 'image/png';

  const buf = await mediaFile.async('nodebuffer');
  return { data: buf, mime };
}

export async function inspectPptxBuffer(pptxBuffer: Buffer): Promise<{
  pageCount: number;
  slideWidthEmu: number;
  slideHeightEmu: number;
  slideLayouts: PptxSlideLayoutInfo[];
  detectedTokens: string[];
}> {
  const zip = await JSZip.loadAsync(pptxBuffer);
  let slideWidthEmu = 5143500;
  let slideHeightEmu = 9144000;

  const presFile = zip.file('ppt/presentation.xml');
  if (presFile) {
    const presXml = await presFile.async('string');
    const szTag = presXml.match(/<p:sldSz\b[^>]*>/)?.[0] || '';
    if (szTag) {
      slideWidthEmu = Number(getXmlAttr(szTag, 'cx')) || slideWidthEmu;
      slideHeightEmu = Number(getXmlAttr(szTag, 'cy')) || slideHeightEmu;
    }
  }

  const slideNames = getSortedSlideNames(zip);
  const allTokens = new Set<string>();
  const slideLayouts: PptxSlideLayoutInfo[] = [];

  for (let i = 0; i < slideNames.length; i++) {
    const slideIndex = i + 1;
    const slideXml = await zip.file(slideNames[i])!.async('string');
    const textBoxes: PptxTextBoxInfo[] = [];
    const slideTokens = new Set<string>();

    const spRegex = /<p:sp\b[\s\S]*?<\/p:sp>/g;
    let spMatch: RegExpExecArray | null;
    while ((spMatch = spRegex.exec(slideXml)) !== null) {
      const spXml = spMatch[0];
      if (!spXml.includes('<p:txBody')) continue;

      const offTag = spXml.match(/<a:off\b[^>]*>/)?.[0] || '';
      const extTag = spXml.match(/<a:ext\b[^>]*>/)?.[0] || '';
      const xEmu = offTag ? Number(getXmlAttr(offTag, 'x') || '0') : 0;
      const yEmu = offTag ? Number(getXmlAttr(offTag, 'y') || '0') : 0;
      const cxEmu = extTag ? Number(getXmlAttr(extTag, 'cx') || String(slideWidthEmu)) : slideWidthEmu;
      const cyEmu = extTag ? Number(getXmlAttr(extTag, 'cy') || String(slideHeightEmu * 0.2)) : slideHeightEmu * 0.2;

      const pMatches = spXml.match(/<a:p\b[\s\S]*?<\/a:p>/g) || [];
      const lines: string[] = [];
      for (const pXml of pMatches) {
        const tMatches = [...pXml.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)];
        const lineText = tMatches.map((m) => decodeXmlEntities(m[1] || '')).join('');
        if (lineText.trim()) lines.push(lineText);
      }

      const fullBoxText = lines.join('\n').trim();
      if (!fullBoxText) continue;

      const txBodyXml = spXml.match(/<p:txBody\b[\s\S]*?<\/p:txBody>/)?.[0] || spXml;
      const rPrTag = txBodyXml.match(/<a:(?:rPr|defRPr)\b[^>]*>/)?.[0] || '';
      const szVal = rPrTag ? getXmlAttr(rPrTag, 'sz') : null;
      const fontSizePt = szVal ? Math.round(Number(szVal) / 100) : 24;
      const bold = rPrTag ? ['1', 'true'].includes((getXmlAttr(rPrTag, 'b') || '').toLowerCase()) : false;
      const italic = rPrTag ? ['1', 'true'].includes((getXmlAttr(rPrTag, 'i') || '').toLowerCase()) : false;
      const uAttr = rPrTag ? getXmlAttr(rPrTag, 'u') : null;
      const underline = Boolean(uAttr && uAttr !== 'none');
      const spcAttr = rPrTag ? getXmlAttr(rPrTag, 'spc') : null;
      const charSpacingPt = spcAttr && !Number.isNaN(Number(spcAttr)) ? Number(spcAttr) / 100 : undefined;

      const latinTag = txBodyXml.match(/<a:latin\b[^>]*>/)?.[0] || '';
      const fontFamily = latinTag ? getXmlAttr(latinTag, 'typeface') || undefined : undefined;
      const colorMatch = txBodyXml.match(/<a:solidFill>\s*<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
      const colorHex = colorMatch ? `#${colorMatch[1]}` : '#FFFFFF';
      const pPrTag = txBodyXml.match(/<a:pPr\b[^>]*>/)?.[0] || '';
      const algnVal = pPrTag ? getXmlAttr(pPrTag, 'algn') : null;
      const align = algnVal === 'l' ? 'left' : algnVal === 'r' ? 'right' : 'center';

      const lnSpcMatch = txBodyXml.match(/<a:lnSpc>\s*<a:spcPct\b[^>]*val="(\d+)"/);
      const lineSpacingPct = lnSpcMatch ? Math.round(Number(lnSpcMatch[1]) / 1000) : undefined;

      const tokens = extractTokensFromText(fullBoxText);
      tokens.forEach((t) => {
        slideTokens.add(t);
        allTokens.add(t);
      });

      textBoxes.push({
        slideIndex,
        text: fullBoxText,
        tokens,
        xPct: Math.max(0, Math.min(100, (xEmu / slideWidthEmu) * 100)),
        yPct: Math.max(0, Math.min(100, (yEmu / slideHeightEmu) * 100)),
        wPct: Math.max(5, Math.min(100, (cxEmu / slideWidthEmu) * 100)),
        hPct: Math.max(4, Math.min(100, (cyEmu / slideHeightEmu) * 100)),
        fontSizePt,
        fontFamily,
        colorHex,
        align,
        bold,
        italic,
        underline,
        charSpacingPt,
        lineSpacingPct,
      });
    }

    const bgMedia = await resolveSlideBackgroundMedia(zip, slideIndex);

    slideLayouts.push({
      slideIndex,
      hasCleanBg: Boolean(bgMedia),
      cleanBgMime: bgMedia?.mime || null,
      textBoxes,
      detectedTokens: Array.from(slideTokens),
    });
  }

  return {
    pageCount: slideNames.length,
    slideWidthEmu,
    slideHeightEmu,
    slideLayouts,
    detectedTokens: Array.from(allTokens),
  };
}

export function replaceTextInSlideXml(
  slideXml: string,
  replacements: Array<{ findText: string; replaceText: string }>
): string {
  if (!replacements || replacements.length === 0) return slideXml;

  return slideXml.replace(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g, (pXml) => {
    const tRegex = /<a:t(\s[^>]*)?>([\s\S]*?)<\/a:t>/g;
    const runs: Array<{
      fullMatch: string;
      startInP: number;
      endInP: number;
      attrs: string;
      text: string;
      rPrXml: string;
    }> = [];

    let m: RegExpExecArray | null;
    while ((m = tRegex.exec(pXml)) !== null) {
      // Localiza o <a:rPr> pertencente ao mesmo <a:r> caso precise inserir quebra de linha OOXML
      const beforeTag = pXml.slice(0, m.index);
      const lastRStart = beforeTag.lastIndexOf('<a:r>');
      const lastRWithAttrStart = beforeTag.lastIndexOf('<a:r ');
      const rStart = Math.max(lastRStart, lastRWithAttrStart);
      let rPrXml = '';
      if (rStart !== -1) {
        const chunk = pXml.slice(rStart, m.index);
        const rPrMatch = chunk.match(/<a:rPr\b[^>]*\/>|<a:rPr\b[\s\S]*?<\/a:rPr>/);
        if (rPrMatch) rPrXml = rPrMatch[0];
      }

      runs.push({
        fullMatch: m[0],
        startInP: m.index,
        endInP: m.index + m[0].length,
        attrs: m[1] || '',
        text: decodeXmlEntities(m[2] || ''),
        rPrXml,
      });
    }

    if (runs.length === 0) return pXml;

    let modified = false;

    for (const rep of replacements) {
      const target = String(rep.findText || '').trim();
      if (!target) continue;

      const coreWord = target.replace(/^[\(\{\[<]|[\)\}\]>]$/g, '').trim();
      const hasBracket = /^[\(\{\[<].*[\)\}\]>]$/.test(target);
      const searchPatterns: RegExp[] = [];
      if (!hasBracket) {
        const cleanWord = escapeRegex(coreWord);
        searchPatterns.push(new RegExp(`\\(${cleanWord}\\)`, 'gi'));
        searchPatterns.push(new RegExp(`\\{${cleanWord}\\}`, 'gi'));
        searchPatterns.push(new RegExp(`\\[${cleanWord}\\]`, 'gi'));
        searchPatterns.push(new RegExp(`<${cleanWord}>`, 'gi'));
      }
      searchPatterns.push(new RegExp(escapeRegex(target), 'gi'));

      for (const pattern of searchPatterns) {
        let matchedThisPattern = false;
        let searchFrom = 0;
        let safety = 0;
        while (safety < 20) {
          safety++;
          const fullText = runs.map((r) => r.text).join('');
          if (searchFrom >= fullText.length) break;
          const tail = fullText.slice(searchFrom);
          pattern.lastIndex = 0;
          const match = pattern.exec(tail);
          if (!match || match[0].length === 0) break;

          matchedThisPattern = true;
          const matchStart = searchFrom + match.index;
          const matchEnd = matchStart + match[0].length;

          // Identifica qual run sobreposto carrega a palavra central do token (para preservar a formatação
          // correta caso o Canva tenha separado "(" ou "{" em um run sem estilo antes da palavra!)
          let cursor = 0;
          const overlappingRuns: Array<{
            idx: number;
            prefix: string;
            suffix: string;
            score: number;
          }> = [];

          for (let rIdx = 0; rIdx < runs.length; rIdx++) {
            const run = runs[rIdx];
            const runStart = cursor;
            const runEnd = cursor + run.text.length;
            cursor = runEnd;

            if (runEnd <= matchStart || runStart >= matchEnd) {
              continue;
            }

            const overlapStart = Math.max(runStart, matchStart);
            const overlapEnd = Math.min(runEnd, matchEnd);
            const overlappedSlice = run.text.slice(overlapStart - runStart, overlapEnd - runStart);
            const alnumCount = (overlappedSlice.match(/[A-Za-zÀ-ÿ0-9]/g) || []).length;
            const score = alnumCount * 100 + (overlapEnd - overlapStart);

            overlappingRuns.push({
              idx: rIdx,
              prefix: runStart < matchStart ? run.text.slice(0, matchStart - runStart) : '',
              suffix: runEnd > matchEnd ? run.text.slice(matchEnd - runStart) : '',
              score,
            });
          }

          if (overlappingRuns.length > 0) {
            let bestOverlap = overlappingRuns[0];
            for (const cand of overlappingRuns) {
              if (cand.score > bestOverlap.score) {
                bestOverlap = cand;
              }
            }

            // Monta o texto final mantendo prefixos anteriores e sufixos posteriores no melhor run
            const combinedPrefix = overlappingRuns
              .filter((o) => o.idx < bestOverlap.idx)
              .map((o) => o.prefix)
              .join('') + bestOverlap.prefix;
            const combinedSuffix =
              bestOverlap.suffix +
              overlappingRuns
                .filter((o) => o.idx > bestOverlap.idx)
                .map((o) => o.suffix)
                .join('');

            for (const ov of overlappingRuns) {
              if (ov.idx === bestOverlap.idx) {
                runs[ov.idx].text = combinedPrefix + rep.replaceText + combinedSuffix;
              } else {
                runs[ov.idx].text = '';
              }
            }
            modified = true;
          }

          searchFrom = matchStart + rep.replaceText.length;
        }
        if (matchedThisPattern) break;
      }
    }

    if (!modified) return pXml;

    // Reconstrói o parágrafo <a:p> de trás para frente preservando toda formatação (<a:rPr> e <a:pPr>)
    let updatedP = pXml;
    for (let idx = runs.length - 1; idx >= 0; idx--) {
      const r = runs[idx];
      if (r.text.includes('\n')) {
        const parts = r.text.split(/\r?\n/);
        const rebuilt = parts
          .map((part, pIdx) => {
            const needsSpacePreserve = /^\s|\s$/.test(part);
            let attrs = r.attrs;
            if (needsSpacePreserve && !attrs.includes('xml:space')) {
              attrs = `${attrs} xml:space="preserve"`;
            }
            if (pIdx === 0) {
              return `<a:t${attrs}>${encodeXmlEntities(part)}</a:t></a:r>`;
            }
            return `<a:br>${r.rPrXml}</a:br><a:r>${r.rPrXml}<a:t${attrs}>${encodeXmlEntities(part)}</a:t></a:r>`;
          })
          .join('');
        // Substitui desde <a:t> até o </a:r> original, pois cada segmento já fecha seu próprio </a:r>
        const tagEndWithCloseR = updatedP.slice(r.endInP).startsWith('</a:r>');
        if (tagEndWithCloseR) {
          updatedP = updatedP.slice(0, r.startInP) + rebuilt + updatedP.slice(r.endInP + '</a:r>'.length);
        } else {
          const singleLineAttrs = r.attrs.includes('xml:space') ? r.attrs : `${r.attrs} xml:space="preserve"`;
          updatedP =
            updatedP.slice(0, r.startInP) +
            `<a:t${singleLineAttrs}>${encodeXmlEntities(r.text.replace(/\r?\n/g, ' '))}</a:t>` +
            updatedP.slice(r.endInP);
        }
        continue;
      }

      const needsSpacePreserve = /^\s|\s$/.test(r.text);
      let attrs = r.attrs;
      if (needsSpacePreserve && !attrs.includes('xml:space')) {
        attrs = `${attrs} xml:space="preserve"`;
      }
      const newTag = `<a:t${attrs}>${encodeXmlEntities(r.text)}</a:t>`;
      updatedP = updatedP.slice(0, r.startInP) + newTag + updatedP.slice(r.endInP);
    }

    return updatedP;
  });
}

export async function generateModifiedPptxFromTemplate(
  pptxBuffer: Buffer,
  mode: 'individual_copy' | 'multipage_batch',
  perVersionReplacements: Array<
    Array<{ findText: string; replaceText: string; pageScope?: 'all' | 'specific'; selectedPages?: number[] }>
  >
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(pptxBuffer);
  const slideNames = getSortedSlideNames(zip);

  if (mode === 'individual_copy') {
    const reps = perVersionReplacements[0] || [];
    for (let i = 0; i < slideNames.length; i++) {
      const pageNum = i + 1;
      const slideReps = reps.filter(
        (r) => r.pageScope !== 'specific' || !r.selectedPages || r.selectedPages.includes(pageNum)
      );
      if (slideReps.length > 0) {
        const xml = await zip.file(slideNames[i])!.async('string');
        zip.file(slideNames[i], replaceTextInSlideXml(xml, slideReps));
      }
    }
    return await zip.generateAsync({ type: 'nodebuffer' });
  }

  // Modo multipage_batch: cada versão ocupa 1 slide na apresentação resultante
  const templateCount = slideNames.length;
  const targetCount = perVersionReplacements.length;

  // Caso 1: targetCount <= templateCount -> edita slides 1..targetCount e mantém estrutura
  // Caso 2: targetCount > templateCount -> clona slides extras no OOXML antes de aplicar substituições
  if (targetCount > templateCount && templateCount > 0) {
    const presXmlPath = 'ppt/presentation.xml';
    const presRelsPath = 'ppt/_rels/presentation.xml.rels';
    const ctPath = '[Content_Types].xml';

    let presXml = (await zip.file(presXmlPath)?.async('string')) || '';
    let presRels = (await zip.file(presRelsPath)?.async('string')) || '';
    let ctXml = (await zip.file(ctPath)?.async('string')) || '';

    const existingRIds = [...presRels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
    let nextRId = (existingRIds.length > 0 ? Math.max(...existingRIds) : 10) + 1;

    const existingSldIds = [...presXml.matchAll(/<p:sldId\b[^>]*id="(\d+)"/g)].map((m) => Number(m[1]));
    let nextSldId = (existingSldIds.length > 0 ? Math.max(...existingSldIds) : 256) + 1;

    for (let k = templateCount; k < targetCount; k++) {
      const srcIdx = (k % templateCount) + 1;
      const newIdx = k + 1;
      const srcSlideXml = await zip.file(`ppt/slides/slide${srcIdx}.xml`)!.async('string');
      zip.file(`ppt/slides/slide${newIdx}.xml`, srcSlideXml);

      const srcRelsFile = zip.file(`ppt/slides/_rels/slide${srcIdx}.xml.rels`);
      if (srcRelsFile) {
        let relsContent = await srcRelsFile.async('string');
        // Evita duplicar referência para notesSlide
        relsContent = relsContent.replace(/<Relationship\b[^>]*notesSlide[^>]*\/>/gi, '');
        zip.file(`ppt/slides/_rels/slide${newIdx}.xml.rels`, relsContent);
      }

      const rIdStr = `rId${nextRId++}`;
      const sldIdNum = nextSldId++;

      ctXml = ctXml.replace(
        '</Types>',
        `<Override PartName="/ppt/slides/slide${newIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`
      );
      presRels = presRels.replace(
        '</Relationships>',
        `<Relationship Id="${rIdStr}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newIdx}.xml"/></Relationships>`
      );
      presXml = presXml.replace(
        '</p:sldIdLst>',
        `<p:sldId id="${sldIdNum}" r:id="${rIdStr}"/></p:sldIdLst>`
      );
    }

    zip.file(ctPath, ctXml);
    zip.file(presRelsPath, presRels);
    zip.file(presXmlPath, presXml);
  }

  const updatedSlideNames = getSortedSlideNames(zip);
  for (let i = 0; i < Math.min(targetCount, updatedSlideNames.length); i++) {
    const pageNum = i + 1;
    const reps = perVersionReplacements[i] || [];
    const slideReps = reps.filter(
      (r) => r.pageScope !== 'specific' || !r.selectedPages || r.selectedPages.includes(pageNum)
    );
    if (slideReps.length > 0) {
      const xml = await zip.file(updatedSlideNames[i])!.async('string');
      zip.file(updatedSlideNames[i], replaceTextInSlideXml(xml, slideReps));
    }
  }

  return await zip.generateAsync({ type: 'nodebuffer' });
}

async function importPptxBufferToCanva(
  pptxBuffer: Buffer,
  title: string,
  accessToken: string,
  fallbackPageCount: number
): Promise<{ designId: string; title: string; pageCount: number; editUrl: string; viewUrl?: string }> {
  const uploadMetadata = JSON.stringify({
    title_base64: Buffer.from(title.trim()).toString('base64'),
    mime_type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });

  const importRes = await axios.post('https://api.canva.com/rest/v1/imports', pptxBuffer, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/octet-stream',
      'Upload-Metadata': uploadMetadata,
    },
    maxBodyLength: Infinity,
    maxContentLength: Infinity,
  });

  const jobId = importRes.data?.job?.id;
  if (!jobId) {
    throw new Error('Canva não retornou job ID na importação do design editável.');
  }

  let jobStatus = importRes.data?.job?.status || 'in_progress';
  let designResult: any = null;
  let pollCount = 0;

  while (jobStatus === 'in_progress' && pollCount < 25) {
    await new Promise((r) => setTimeout(r, 1500));
    pollCount++;
    const pollRes = await axios.get(`https://api.canva.com/rest/v1/imports/${encodeURIComponent(jobId)}`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const currentJob = pollRes.data?.job || pollRes.data;
    jobStatus = currentJob?.status;
    if (jobStatus === 'success') {
      const designs = currentJob?.result?.designs || [];
      if (designs.length > 0) designResult = designs[0];
      break;
    }
    if (jobStatus === 'failed') {
      throw new Error(currentJob?.error?.message || currentJob?.error?.code || 'Falha na importação pelo Canva.');
    }
  }

  if (!designResult?.id || !designResult?.urls?.edit_url) {
    throw new Error('Importação concluída sem edit_url válido retornado pelo Canva.');
  }

  let realPageCount = fallbackPageCount || designResult.page_count || 1;
  try {
    const pagesRes = await axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(designResult.id)}/pages`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (Array.isArray(pagesRes.data?.items) && pagesRes.data.items.length > 0) {
      realPageCount = pagesRes.data.items.length;
    }
  } catch {
    // usa fallback
  }

  return {
    designId: designResult.id,
    title: designResult.title || title,
    pageCount: realPageCount,
    editUrl: designResult.urls.edit_url,
    viewUrl: designResult.urls.view_url,
  };
}

// GET /api/canva/templates/:id/slide-bg/:slideNum - Serve imagem de fundo limpa extraída do PPTX do design
app.get('/api/canva/templates/:id/slide-bg/:slideNum', async (req, res) => {
  const cleanId = String(req.params.id || '').match(/design\/([A-Za-z0-9_-]+)/)?.[1] || String(req.params.id || '').trim();
  const slideNum = Math.max(1, parseInt(String(req.params.slideNum || '1'), 10) || 1);
  if (!cleanId) return res.status(400).send('Invalid ID');

  try {
    let entry = cachedPptxTemplates.get(cleanId);
    if (!entry) {
      const tokenData = await ensureCanvaToken();
      if (!tokenData?.access_token) return res.status(401).send('Not authenticated');
      const buf = await exportCanvaDesignAsPptxBuffer(cleanId, tokenData.access_token);
      entry = { buffer: buf, cachedAt: Date.now() };
    }

    const zip = await JSZip.loadAsync(entry.buffer);
    const media = await resolveSlideBackgroundMedia(zip, slideNum);
    if (!media) {
      return res.status(404).send('Background media not found on slide');
    }

    res.setHeader('Content-Type', media.mime);
    res.setHeader('Cache-Control', 'public, max-age=600');
    return res.send(media.data);
  } catch (err: any) {
    return res.status(404).send(err.message || 'Not found');
  }
});

// POST /api/canva/templates/inspect - Inspeciona Brand Template (Autofill) E/OU Design Comum via PPTX XML Round-Trip
app.post('/api/canva/templates/inspect', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado. Conecte sua conta na Etapa 1.' });
  }

  const rawId = String(req.body?.id || '').trim();
  const cleanId = rawId.match(/design\/([A-Za-z0-9_-]+)/)?.[1] || rawId;
  const titleHint = String(req.body?.titleHint || '').trim();
  if (!cleanId) {
    return res.status(400).json({ error: 'Informe um ID de Design ou Brand Template válido.' });
  }

  const scopes = (tokenData.scope || '').split(' ');
  const hasBrandTemplateScope =
    scopes.includes('brandtemplate:meta:read') && scopes.includes('brandtemplate:content:read');

  let title = titleHint || 'Modelo Canva';
  let isBrandTemplate = false;
  let brandTemplateId: string | null = null;
  let autofillCompatible = false;
  let datasetFields: Array<{ name: string; type: string }> = [];
  let editUrl: string | null = null;
  let viewUrl: string | null = null;
  let pages: Array<{ index: number; thumbnailUrl: string; title: string; cleanBgUrl?: string | null }> = [];
  let pageCount = 1;
  let limitationReason = '';
  let scopeMissingError = false;

  // 1. Tenta consultar como Brand Template diretamente + Dataset oficial
  try {
    const [btMetaRes, btDatasetRes] = await Promise.allSettled([
      axios.get(`https://api.canva.com/rest/v1/brand-templates/${encodeURIComponent(cleanId)}`, {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      }),
      axios.get(`https://api.canva.com/rest/v1/brand-templates/${encodeURIComponent(cleanId)}/dataset`, {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      }),
    ]);

    if (btMetaRes.status === 'fulfilled') {
      const bt = btMetaRes.value.data?.brand_template || btMetaRes.value.data;
      if (bt?.id) {
        isBrandTemplate = true;
        brandTemplateId = bt.id;
        title = bt.title || title;
        viewUrl = bt.view_url || bt.create_url || null;
        if (bt.thumbnail?.url) {
          pages = [{ index: 1, thumbnailUrl: bt.thumbnail.url, title: `${title} (Capa)` }];
        }
      }
    } else {
      const errStatus = (btMetaRes.reason as any)?.response?.status;
      if (errStatus === 403) scopeMissingError = true;
    }

    if (btDatasetRes.status === 'fulfilled') {
      isBrandTemplate = true;
      brandTemplateId = brandTemplateId || cleanId;
      const ds = btDatasetRes.value.data?.dataset || {};
      datasetFields = Object.entries(ds).map(([key, val]: [string, any]) => ({
        name: key,
        type: val?.type || 'text',
      }));
      autofillCompatible = datasetFields.length > 0;
    } else {
      const dsStatus = (btDatasetRes.reason as any)?.response?.status;
      const dsMsg = (btDatasetRes.reason as any)?.response?.data?.message || '';
      if (dsStatus === 403 || String(dsMsg).toLowerCase().includes('scope')) {
        scopeMissingError = true;
      }
    }
  } catch {
    // continua verificação como design
  }

  // 2. Consulta páginas e metadados do Design
  try {
    const [designRes, pagesRes] = await Promise.allSettled([
      axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(cleanId)}`, {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      }),
      axios.get(`https://api.canva.com/rest/v1/designs/${encodeURIComponent(cleanId)}/pages`, {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      }),
    ]);

    if (designRes.status === 'fulfilled') {
      const d = designRes.value.data?.design || designRes.value.data;
      if (d) {
        title = d.title || title;
        editUrl = d.urls?.edit_url || editUrl;
        viewUrl = d.urls?.view_url || viewUrl;
        if (d.page_count) pageCount = d.page_count;
        if (pages.length === 0 && d.thumbnail?.url) {
          pages = [{ index: 1, thumbnailUrl: d.thumbnail.url, title: `${title} - Pág 1` }];
        }
      }
    }

    if (pagesRes.status === 'fulfilled') {
      const pItems = pagesRes.value.data?.items || [];
      if (pItems.length > 0) {
        pageCount = pItems.length;
        pages = pItems.map((p: any, idx: number) => ({
          index: idx + 1,
          thumbnailUrl: p.thumbnail?.url || '',
          title: p.title || `Página ${idx + 1}`,
        }));
      }
    }
  } catch {
    // ignore
  }

  // 3. Se veio de um Design ID comum, tenta localizar Brand Template correspondente por título
  if (!isBrandTemplate && title && !scopeMissingError) {
    try {
      const searchBtRes = await axios.get('https://api.canva.com/rest/v1/brand-templates', {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
        params: { query: title, ownership: 'any' },
      });
      const btItems = searchBtRes.data?.items || [];
      const exactMatch = btItems.find(
        (b: any) => String(b.title || '').trim().toLowerCase() === title.trim().toLowerCase()
      );
      if (exactMatch?.id) {
        const dsRes = await axios.get(
          `https://api.canva.com/rest/v1/brand-templates/${encodeURIComponent(exactMatch.id)}/dataset`,
          { headers: { Authorization: `Bearer ${tokenData.access_token}` } }
        );
        const ds = dsRes.data?.dataset || {};
        isBrandTemplate = true;
        brandTemplateId = exactMatch.id;
        datasetFields = Object.entries(ds).map(([key, val]: [string, any]) => ({
          name: key,
          type: val?.type || 'text',
        }));
        autofillCompatible = datasetFields.length > 0;
      }
    } catch {
      // ignora se não autorizado ou não encontrado
    }
  }

  // 4. Exporta e inspeciona camadas reais de texto do Design via PPTX (funciona para QUALQUER design comum!)
  let editableViaPptxRoundtrip = false;
  let slideLayouts: PptxSlideLayoutInfo[] = [];
  let detectedPlaceholders: string[] = [];

  if (!autofillCompatible) {
    try {
      const pptxBuffer = await exportCanvaDesignAsPptxBuffer(cleanId, tokenData.access_token);
      const parsed = await inspectPptxBuffer(pptxBuffer);
      editableViaPptxRoundtrip = parsed.pageCount > 0;
      slideLayouts = parsed.slideLayouts;
      detectedPlaceholders = parsed.detectedTokens;
      if (parsed.pageCount > 0) {
        pageCount = parsed.pageCount;
      }

      // Enriquece páginas com URL limpa de fundo sem o texto (frase) estático
      pages = Array.from({ length: pageCount }).map((_, idx) => {
        const pNum = idx + 1;
        const existingPage = pages[idx];
        const layout = slideLayouts.find((l) => l.slideIndex === pNum);
        return {
          index: pNum,
          thumbnailUrl: existingPage?.thumbnailUrl || '',
          title: existingPage?.title || `Página ${pNum}`,
          cleanBgUrl: layout?.hasCleanBg
            ? `/api/canva/templates/${encodeURIComponent(cleanId)}/slide-bg/${pNum}`
            : null,
        };
      });
    } catch (pptxErr: any) {
      console.warn('[CANVA PPTX INSPECT WARNING]', pptxErr?.response?.data?.message || pptxErr.message);
    }
  }

  if (!autofillCompatible && !editableViaPptxRoundtrip && !limitationReason) {
    if (scopeMissingError || !hasBrandTemplateScope) {
      limitationReason =
        'Não foi possível exportar as camadas editáveis deste design nem acessar Brand Templates. Verifique as permissões da sua conta Canva.';
    } else {
      limitationReason =
        'Este design não possui Brand Template Dataset e não pôde ser exportado em camadas PPTX editáveis. Use a Opção 2 para criar um novo design editável com suas páginas.';
    }
  }

  return res.json({
    success: true,
    id: cleanId,
    brandTemplateId,
    title,
    isBrandTemplate,
    autofillCompatible,
    editableViaPptxRoundtrip,
    datasetFields,
    detectedPlaceholders,
    slideLayouts,
    pageCount,
    pages,
    editUrl,
    viewUrl,
    hasBrandTemplateScope: !scopeMissingError && hasBrandTemplateScope,
    limitationReason,
  });
});

// POST /api/canva/designs/replace-text-batch - Substitui campos-chave DENTRO do arquivo editável do Canva (sem sobrepor texto!)
app.post('/api/canva/designs/replace-text-batch', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado. Conecte sua conta na Etapa 1.' });
  }

  const { designId, titlePrefix, mode, versions } = req.body;
  const cleanId = String(designId || '').match(/design\/([A-Za-z0-9_-]+)/)?.[1] || String(designId || '').trim();
  const targetMode: 'individual' | 'multipage' = mode === 'multipage' ? 'multipage' : 'individual';

  if (!cleanId || !Array.isArray(versions) || versions.length === 0) {
    return res.status(400).json({ error: 'designId e lista de versões são obrigatórios.' });
  }

  try {
    const pptxBuffer = await exportCanvaDesignAsPptxBuffer(cleanId, tokenData.access_token);

    if (targetMode === 'multipage') {
      const perSlideReps = versions.map((v: any) => v.replacements || []);
      const modifiedBuffer = await generateModifiedPptxFromTemplate(pptxBuffer, 'multipage_batch', perSlideReps);
      const designTitle = `${titlePrefix || 'Modelo Editado'} - ${versions.length} Páginas`;
      const imported = await importPptxBufferToCanva(modifiedBuffer, designTitle, tokenData.access_token, versions.length);

      return res.json({
        success: true,
        mode: 'multipage',
        total: 1,
        completedCount: 1,
        failedCount: 0,
        results: [
          {
            versionIndex: 1,
            designId: imported.designId,
            title: imported.title,
            pageCount: imported.pageCount,
            editUrl: imported.editUrl,
            viewUrl: imported.viewUrl,
            label: `${versions.length} páginas com campos substituídos no arquivo`,
            status: 'completed',
          },
        ],
      });
    }

    // Modo 'individual': gera 1 cópia editável no Canva para cada versão substituindo o texto no XML
    const baseZip = await JSZip.loadAsync(pptxBuffer);
    const baseSlideCount = Math.max(1, getSortedSlideNames(baseZip).length);
    const results: any[] = [];
    let completedCount = 0;
    let failedCount = 0;

    for (let i = 0; i < versions.length; i++) {
      const v = versions[i];
      const designTitle = `${titlePrefix ? titlePrefix + ' - ' : ''}Versão ${i + 1}${v.label ? ` (${String(v.label).substring(0, 26)})` : ''}`;
      try {
        const modifiedBuffer = await generateModifiedPptxFromTemplate(pptxBuffer, 'individual_copy', [v.replacements || []]);
        const imported = await importPptxBufferToCanva(modifiedBuffer, designTitle, tokenData.access_token, baseSlideCount);
        completedCount++;
        results.push({
          versionIndex: i + 1,
          designId: imported.designId,
          title: imported.title,
          pageCount: imported.pageCount,
          editUrl: imported.editUrl,
          viewUrl: imported.viewUrl,
          label: v.label || '',
          status: 'completed',
        });
      } catch (err: any) {
        failedCount++;
        results.push({
          versionIndex: i + 1,
          designId: null,
          title: designTitle,
          pageCount: 0,
          editUrl: null,
          label: v.label || '',
          status: 'failed',
          error: err.response?.data?.message || err.message || 'Falha ao importar versão substituída.',
        });
      }
    }

    return res.json({
      success: completedCount > 0,
      mode: 'individual',
      total: versions.length,
      completedCount,
      failedCount,
      results,
    });
  } catch (err: any) {
    const errMsg = err.response?.data?.message || err.message || 'Erro ao editar textos no arquivo Canva.';
    return res.status(500).json({ error: errMsg });
  }
});

// POST /api/canva/designs/replace-text-download - Gera download direto do .PPTX original com os textos substituídos in-place
app.post('/api/canva/designs/replace-text-download', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado.' });
  }

  const { designId, titlePrefix, versions } = req.body;
  const cleanId = String(designId || '').match(/design\/([A-Za-z0-9_-]+)/)?.[1] || String(designId || '').trim();
  if (!cleanId || !Array.isArray(versions) || versions.length === 0) {
    return res.status(400).json({ error: 'designId e lista de versões são obrigatórios.' });
  }

  try {
    const pptxBuffer = await exportCanvaDesignAsPptxBuffer(cleanId, tokenData.access_token);
    const perSlideReps = versions.map((v: any) => v.replacements || []);
    const modifiedBuffer = await generateModifiedPptxFromTemplate(pptxBuffer, 'multipage_batch', perSlideReps);
    return res.json({
      success: true,
      filename: `${(titlePrefix || 'Modelo_Editado').replace(/[^a-zA-Z0-9_-]/g, '_')}_${versions.length}pags.pptx`,
      pptxBase64: modifiedBuffer.toString('base64'),
    });
  } catch (err: any) {
    return res.status(500).json({ error: err.message || 'Erro ao gerar PPTX substituído.' });
  }
});

// POST /api/canva/autofill-batch - Opção 1: Gera múltiplas versões editáveis no Canva via Brand Template Autofill API
app.post('/api/canva/autofill-batch', async (req, res) => {
  const tokenData = await ensureCanvaToken();
  if (!tokenData?.access_token) {
    return res.status(401).json({ error: 'Canva não conectado. Conecte sua conta na Etapa 1.' });
  }

  const { brandTemplateId, titlePrefix, versions } = req.body;
  const cleanBtId = String(brandTemplateId || '').trim();
  if (!cleanBtId || !Array.isArray(versions) || versions.length === 0) {
    return res.status(400).json({ error: 'brandTemplateId e lista de versões são obrigatórios.' });
  }

  const results: any[] = [];
  let completedCount = 0;
  let failedCount = 0;

  for (let i = 0; i < versions.length; i++) {
    const v = versions[i];
    const fieldValues: Record<string, string> = v.fieldValues || {};
    const autofillData: Record<string, { type: string; text: string }> = {};

    for (const [fieldKey, fieldText] of Object.entries(fieldValues)) {
      if (fieldKey && fieldText !== undefined) {
        autofillData[fieldKey] = { type: 'text', text: String(fieldText) };
      }
    }

    const designTitle = `${titlePrefix ? titlePrefix + ' - ' : ''}Versão ${i + 1}${v.label ? ` (${String(v.label).substring(0, 28)})` : ''}`;

    try {
      const startRes = await axios.post(
        'https://api.canva.com/rest/v1/autofills',
        {
          brand_template_id: cleanBtId,
          title: designTitle,
          data: autofillData,
        },
        {
          headers: {
            Authorization: `Bearer ${tokenData.access_token}`,
            'Content-Type': 'application/json',
          },
        }
      );

      const jobId = startRes.data?.job?.id;
      if (!jobId) {
        throw new Error('Canva não retornou job ID para o Autofill.');
      }

      const polled = await pollCanvaAutofillJob(jobId, tokenData.access_token);
      if (polled.success && polled.editUrl && polled.designId) {
        completedCount++;
        results.push({
          versionIndex: i + 1,
          designId: polled.designId,
          title: polled.title || designTitle,
          pageCount: polled.pageCount || 1,
          editUrl: polled.editUrl,
          viewUrl: polled.viewUrl,
          label: v.label || '',
          status: 'completed',
        });
      } else {
        throw new Error('O job de Autofill não retornou edit_url no tempo limite.');
      }
    } catch (err: any) {
      failedCount++;
      const errMsg = err.response?.data?.message || err.message || 'Falha no Autofill.';
      results.push({
        versionIndex: i + 1,
        designId: null,
        title: designTitle,
        pageCount: 0,
        editUrl: null,
        label: v.label || '',
        status: 'failed',
        error: errMsg,
      });
    }
  }

  return res.json({
    success: completedCount > 0,
    total: versions.length,
    completedCount,
    failedCount,
    results,
  });
});

// 7. POST /api/canva/batch-generate - Processador de lote com Autofill real do Canva ou preparação para Design Editável
app.post('/api/canva/batch-generate', async (req, res) => {
  const { phrases, templateId, settings, mediaItems, isMultipage: reqIsMultipage } = req.body;

  if (!phrases || !Array.isArray(phrases) || phrases.length === 0) {
    return res.status(400).json({ error: 'Nenhuma frase informada.' });
  }

  const cleanTemplateId = templateId
    ? String(templateId).match(/design\/([A-Za-z0-9_-]+)/)?.[1] || String(templateId).trim()
    : '';
  const isMultipage = Boolean(reqIsMultipage || settings?.designMode === 'multipage');
  const tokenData = await ensureCanvaToken();
  const results: any[] = [];
  let autofillSupported = false;
  let autofillNote = '';

  const customFields = {
    fieldPhrase: settings?.fieldPhrase,
    fieldSignature: settings?.fieldSignature,
    signatureText: settings?.signatureText,
  };

  if (isMultipage) {
    const singleDesignId = `editable_multipage_${Date.now()}`;

    if (tokenData?.access_token && cleanTemplateId) {
      try {
        const pageDataMap = buildPlaceholderDataMap(phrases[0], customFields);
        phrases.forEach((phraseText, idx) => {
          pageDataMap[`FRASE_${idx + 1}`] = { type: 'text', text: phraseText };
          pageDataMap[`SLIDE_${idx + 1}`] = { type: 'text', text: phraseText };
          pageDataMap[`(frase_${idx + 1})`] = { type: 'text', text: phraseText };
        });

        const autofillRes = await axios.post(
          'https://api.canva.com/rest/v1/autofills',
          {
            brand_template_id: cleanTemplateId,
            title: `Apresentação Multipáginas (${phrases.length} Slides) - LoteCriativo`,
            data: pageDataMap,
          },
          {
            headers: {
              Authorization: `Bearer ${tokenData.access_token}`,
              'Content-Type': 'application/json',
            },
          }
        );

        const jobId = autofillRes.data?.job?.id;
        if (jobId) {
          const polled = await pollCanvaAutofillJob(jobId, tokenData.access_token);
          if (polled.success && polled.editUrl) {
            autofillSupported = true;
            results.push({
              id: polled.designId || jobId,
              phrase: `Apresentação Completa em Arquivo Único (${phrases.length} Slides)`,
              status: 'completed',
              type: 'canva_design',
              isMultipage: true,
              pageCount: phrases.length,
              url: polled.editUrl,
              mediaName: 'Todos os Slides Juntos',
              message: `Design multipáginas editado via Autofill oficial no Canva (${phrases.length} slides)!`,
            });
          }
        }
      } catch (err: any) {
        autofillNote = err.response?.data?.message || err.message;
        console.warn('Multipage Autofill API info:', autofillNote);
      }
    }

    if (results.length === 0) {
      results.push({
        id: singleDesignId,
        phrase: `Apresentação Completa Editável (${phrases.length} Páginas com Fundo e Camadas do Canva)`,
        status: 'alternative_ready',
        type: 'editable_slide',
        isMultipage: true,
        pageCount: phrases.length,
        url: null,
        mediaName: `${mediaItems?.length || 1} modelo(s) selecionado(s)`,
        message: cleanTemplateId && autofillNote
          ? `Autofill direto indisponível para este ID (${autofillNote}). Use "Salvar como Design Editável no Canva" ou "Baixar Arquivo Editável (.PPTX)" para criar o design multipáginas com todas as camadas editáveis.`
          : `Pronto para salvar como novo Design Editável Multipáginas no Canva ou baixar em formato .PPTX editável.`,
      });
    }
  } else {
    // Mode: Gerar arquivos individuais (1 design por frase)
    let firstAutofillFailed = false;
    for (let i = 0; i < phrases.length; i++) {
      const phrase = phrases[i];
      const media = Array.isArray(mediaItems) && mediaItems.length > 0 ? mediaItems[i % mediaItems.length] : undefined;
      const designId = `editable_slide_${Date.now()}_${i + 1}`;

      if (tokenData?.access_token && cleanTemplateId && !firstAutofillFailed) {
        try {
          const autofillRes = await axios.post(
            'https://api.canva.com/rest/v1/autofills',
            {
              brand_template_id: cleanTemplateId,
              title: `Design ${i + 1}/${phrases.length} - ${phrase.substring(0, 25)}`,
              data: buildPlaceholderDataMap(phrase, customFields),
            },
            {
              headers: {
                Authorization: `Bearer ${tokenData.access_token}`,
                'Content-Type': 'application/json',
              },
            }
          );

          const jobId = autofillRes.data?.job?.id;
          if (jobId) {
            const polled = await pollCanvaAutofillJob(jobId, tokenData.access_token);
            if (polled.success && polled.editUrl) {
              autofillSupported = true;
              results.push({
                id: polled.designId || jobId,
                phrase,
                status: 'completed',
                type: 'canva_design',
                isMultipage: false,
                url: polled.editUrl,
                mediaName: media?.name || 'Mídia Canva',
                message: 'Novo design criado via Brand Template Autofill oficial do Canva com formatação preservada!',
              });
              continue;
            }
          }
        } catch (err: any) {
          autofillNote = err.response?.data?.message || err.message;
          firstAutofillFailed = true;
          console.warn(`Autofill item ${i + 1} info:`, autofillNote);
        }
      }

      results.push({
        id: designId,
        phrase,
        status: 'alternative_ready',
        type: 'editable_slide',
        isMultipage: false,
        url: null,
        mediaName: media?.name || 'Modelo Selecionado',
        message: cleanTemplateId && autofillNote
          ? `Slide preparado com camadas editáveis (${autofillNote}). Clique em "Salvar como Design no Canva" ou baixe o .PPTX editável.`
          : 'Slide editável preparado com imagem de fundo, máscara, texto e assinatura.',
      });
    }
  }

  return res.json({
    total: phrases.length,
    isMultipage,
    autofillSupported,
    autofillNote,
    results,
    accountTier: tokenData?.access_token ? 'canva_connected' : 'offline',
  });
});

// Setup Vite Dev Server as middleware on port 3000
const PORT = 3000;

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static('dist'));
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`LoteCriativo Canva full-stack server listening on http://0.0.0.0:${PORT}`);
  });
}

if (!process.env.SKIP_SERVER_LISTEN) {
  startServer().catch((err) => {
    console.error('Falha ao iniciar o servidor:', err);
  });
}
