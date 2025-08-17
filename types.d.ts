// Environment variables and bindings
interface Env {
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_CLIENT_SECRET: string;
  MICROSOFT_TENANT_ID: string;
  MICROSOFT_WEBHOOK_SECRET: string;
  MICROSOFT_WEBHOOK_URL: string;
  MICROSOFT_MCP_OBJECT: DurableObjectNamespace;
  MICROSOFT_EVENT_CACHE: KVNamespace;
}

export type Todo = {
  id: string;
  text: string;
  completed: boolean;
};

// Context from the auth process, extracted from the Stytch auth token JWT
// and provided to the MCP Server as this.props
type AuthenticationContext = {
  claims: {
    iss: string;
    scope: string;
    sub: string;
    aud: string[];
    client_id: string;
    exp: number;
    iat: number;
    nbf: number;
    jti: string;
  };
  accessToken: string;
};

// Context from the Microsoft OAuth process
export type MicrosoftAuthContext = {
  accessToken: string;
  expiresIn?: number;
  tokenType?: string;
  scope?: string;
};

// Webhook response contract for proxied webhook handling
export type WebhookResponse = {
  /** HTTP status code to proxy back to the origin of the webhook */
  reqResponseCode: number;
  /** body string to proxy back; if JSON, stringify it */
  reqResponseContent: string;
  /** content type for reqResponseContent: "json" or "text" */
  reqResponseContentType?: "json" | "text";
  /** optional return to run with the agent to do something */
  promptContent?: string;
};
