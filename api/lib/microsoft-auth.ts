import { createMiddleware } from "hono/factory";
import { HTTPException } from "hono/http-exception";

/* ---------- Bearer-token middleware ---------- */

export const microsoftBearerTokenAuthMiddleware = createMiddleware<{
  Bindings: Env;
}>(async (c, next) => {
  const auth = c.req.header("Authorization") ?? "";
  if (!auth.startsWith("Bearer ")) {
    throw new HTTPException(401, {
      message: "Missing or invalid access token",
    });
  }

  // Slice off "Bearer "
  const accessToken = auth.slice(7);
  const refreshToken = c.req.header("X-Microsoft-Refresh-Token") ?? "";

  // @ts-expect-error  – Cloudflare Workers executionCtx props
  c.executionCtx.props = { accessToken, refreshToken };
  await next();
});

/* ---------- Helpers ---------- */

export const MICROSOFT_GRAPH_DEFAULT_SCOPES = [
  "openid",
  "profile",
  "offline_access", // <-- needed to receive a refresh_token
  "Calendars.ReadWrite",
  "Mail.ReadWrite",
  "Mail.Send",
  "User.Read",
  "People.Read",
  "Contacts.ReadWrite",
  "MailboxSettings.Read",
];

export function getMicrosoftAuthEndpoint(
  tenantId: string,
  endpoint: "authorize" | "token"
): string {
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/${endpoint}`;
}

type TokenResponse = {
  access_token: string;
  token_type: string;
  scope: string;
  expires_in: number;
  refresh_token?: string;
  id_token?: string;
};

function form(params: Record<string, string | undefined>): URLSearchParams {
  const body = new URLSearchParams();
  Object.entries(params).forEach(([k, v]) => {
    if (v) body.append(k, v);
  });
  return body;
}

/* ---------- Code-grant exchange ---------- */

export async function exchangeCodeForToken(
  code: string,
  redirectUri: string,
  clientId: string,
  clientSecret: string | undefined,
  tenantId: string,
  codeVerifier?: string,
  scopes: string[] = MICROSOFT_GRAPH_DEFAULT_SCOPES
): Promise<TokenResponse> {
  const body = form({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    scope: scopes.join(" "),
    code_verifier: codeVerifier,
  });

  const res = await fetch(getMicrosoftAuthEndpoint(tenantId, "token"), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) throw new Error(`Token exchange failed: ${await res.text()}`);
  return res.json() as Promise<TokenResponse>;
}

/* ---------- Refresh-token flow ---------- */

export async function refreshAccessToken(
  refreshToken: string,
  clientId: string,
  clientSecret: string | undefined,
  tenantId: string,
  scopes: string[] = MICROSOFT_GRAPH_DEFAULT_SCOPES
): Promise<TokenResponse> {
  const body = form({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: scopes.join(" "),
  });

  const res = await fetch(getMicrosoftAuthEndpoint(tenantId, "token"), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) throw new Error(`Refresh failed: ${await res.text()}`);
  return res.json() as Promise<TokenResponse>;
}
