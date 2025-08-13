import { MicrosoftMCP } from "./MicrosoftMCP.ts";
import {
  microsoftBearerTokenAuthMiddleware,
  getMicrosoftAuthEndpoint,
  exchangeCodeForToken,
  refreshAccessToken,
} from "./lib/microsoft-auth.ts";
import { cors } from "hono/cors";
import { Hono } from "hono";
import type { WebhookResponse } from "../types";

// Export the MicrosoftMCP class so the Worker runtime can find it
export { MicrosoftMCP };

// Store registered clients in memory (in production, use a database)
interface RegisteredClient {
  client_id: string;
  client_name: string;
  redirect_uris: string[];
  grant_types: string[];
  response_types: string[];
  scope?: string;
  token_endpoint_auth_method: string;
  created_at: number;
}
const registeredClients = new Map<string, RegisteredClient>();

export default new Hono<{ Bindings: Env }>()
  .use(cors())
  .onError((err, c) => {
    const e = err as unknown as {
      status?: number;
      wwwAuthenticate?: string;
      message?: string;
    };
    if (e && e.status === 401) {
      if (e.wwwAuthenticate) {
        c.header("WWW-Authenticate", e.wwwAuthenticate);
      } else {
        c.header(
          "WWW-Authenticate",
          'Bearer error="invalid_token", error_description="The access token is invalid or expired"'
        );
      }
      c.header("Cache-Control", "no-store");
      c.header("Pragma", "no-cache");
      return c.json(
        {
          error: "invalid_token",
          error_description:
            e.message || "The access token is invalid or expired",
        },
        401 as const
      );
    }
    return c.json({ error: "server_error" }, 500 as const);
  })

  // OAuth Authorization Server Discovery
  .get("/.well-known/oauth-authorization-server", async (c) => {
    const url = new URL(c.req.url);
    return c.json({
      issuer: url.origin,
      authorization_endpoint: `${url.origin}/authorize`,
      token_endpoint: `${url.origin}/token`,
      registration_endpoint: `${url.origin}/register`,
      response_types_supported: ["code"],
      response_modes_supported: ["query"],
      grant_types_supported: ["authorization_code", "refresh_token"],
      token_endpoint_auth_methods_supported: ["none"],
      code_challenge_methods_supported: ["S256", "plain"],
      scopes_supported: [
        "openid",
        "profile",
        "offline_access",
        // TODO: Add scopes for outlook
        "Calendars.ReadWrite",
        "Mail.ReadWrite",
        "Mail.Send",
        "User.Read",
        "People.Read",
        "Contacts.ReadWrite",
        "MailboxSettings.Read",
      ],
    });
  })

  // Dynamic Client Registration endpoint
  .post("/register", async (c) => {
    const body = await c.req.json();

    // Generate a client ID
    const clientId = crypto.randomUUID();

    // Store the client registration
    registeredClients.set(clientId, {
      client_id: clientId,
      client_name: body.client_name || "MCP Client",
      redirect_uris: body.redirect_uris || [],
      grant_types: body.grant_types || ["authorization_code", "refresh_token"],
      response_types: body.response_types || ["code"],
      scope: body.scope,
      token_endpoint_auth_method: "none",
      created_at: Date.now(),
    });

    // Return the client registration response
    return c.json(
      {
        client_id: clientId,
        client_name: body.client_name || "MCP Client",
        redirect_uris: body.redirect_uris || [],
        grant_types: body.grant_types || [
          "authorization_code",
          "refresh_token",
        ],
        response_types: body.response_types || ["code"],
        scope: body.scope,
        token_endpoint_auth_method: "none",
      },
      201
    );
  })

  // Authorization endpoint - redirects to Microsoft
  .get("/authorize", async (c) => {
    const url = new URL(c.req.url);

    const microsoftAuthUrl = new URL(
      getMicrosoftAuthEndpoint(c.env.MICROSOFT_TENANT_ID, "authorize")
    );

    // Copy all query parameters except client_id
    url.searchParams.forEach((value, key) => {
      if (key !== "client_id") {
        microsoftAuthUrl.searchParams.set(key, value);
      }
    });

    // Use our Microsoft app's client_id
    microsoftAuthUrl.searchParams.set("client_id", c.env.MICROSOFT_CLIENT_ID);

    const m = microsoftAuthUrl.searchParams.get("code_challenge_method"); // 'S256' | 'plain' | null
    const cc = microsoftAuthUrl.searchParams.get("code_challenge");
    const ru = microsoftAuthUrl.searchParams.get("redirect_uri");
    console.log("PKCE authorize:", {
      method: m,
      code_challenge: cc?.slice(0, 8) + "…",
      redirect_uri: ru,
    });

    // Redirect to Microsoft's authorization page
    return c.redirect(microsoftAuthUrl.toString());
  })

  // Token exchange endpoint
  .post("/token", async (c) => {
    const body = await c.req.parseBody();
    const cv = body.code_verifier as string | undefined;
    console.log("PKCE token:", {
      code_verifier_len: cv?.length,
      redirect_uri: body.redirect_uri,
    });

    try {
      if (body.grant_type === "authorization_code") {
        const result = await exchangeCodeForToken(
          body.code as string,
          body.redirect_uri as string,
          c.env.MICROSOFT_CLIENT_ID,
          c.env.MICROSOFT_CLIENT_SECRET,
          c.env.MICROSOFT_TENANT_ID,
          body.code_verifier as string | undefined,
          (body.scope as string | undefined) ||
            (typeof body.scope === "string"
              ? (body.scope as string)
              : undefined)
        );
        return c.json(result);
      } else if (body.grant_type === "refresh_token") {
        const result = await refreshAccessToken(
          body.refresh_token as string,
          c.env.MICROSOFT_CLIENT_ID,
          c.env.MICROSOFT_CLIENT_SECRET,
          c.env.MICROSOFT_TENANT_ID
        );
        return c.json(result);
      }
    } catch (err) {
      // Pass through OAuth errors from Microsoft
      const e = err as unknown as {
        status?: number;
        body?: unknown;
        name?: string;
      };
      if (e && e.name === "OAuthHttpError") {
        const allowedStatuses = [
          400, 401, 403, 404, 405, 409, 410, 415, 422, 429, 500, 502, 503, 504,
        ] as const;
        const statusCandidate = (e.status as number) || 400;
        const status = (allowedStatuses as readonly number[]).includes(
          statusCandidate
        )
          ? (statusCandidate as
              | 400
              | 401
              | 403
              | 404
              | 405
              | 409
              | 410
              | 415
              | 422
              | 429
              | 500
              | 502
              | 503
              | 504)
          : (400 as const);
        return c.json(e.body ?? { error: "invalid_request" }, { status });
      }
      throw err;
    }

    return c.json({ error: "unsupported_grant_type" }, 400);
  })

  // Microsoft MCP endpoints
  .use("/sse/*", microsoftBearerTokenAuthMiddleware)
  .route(
    "/sse",
    new Hono().mount(
      "/",
      MicrosoftMCP.serveSSE("/sse", { binding: "MICROSOFT_MCP_OBJECT" }).fetch
    )
  )

  .use("/mcp", microsoftBearerTokenAuthMiddleware)
  .route(
    "/mcp",
    new Hono().mount(
      "/",
      MicrosoftMCP.serve("/mcp", { binding: "MICROSOFT_MCP_OBJECT" }).fetch
    )
  )

  .route(
    "/webhooks",
    new Hono<{ Bindings: Env }>()

      // Notification payloads
      .post("/email-notify", async (c) => {
        const url = new URL(c.req.url);
        const validationToken = url.searchParams.get("validationToken");
        // Validation challenge from Microsoft Graph
        if (validationToken) {
          const response: WebhookResponse = {
            reqResponseCode: 200,
            reqResponseContent: validationToken,
            reqResponseContentType: "text",
          };

          return c.json(response);
        }
        const body = await c.req.json();
        const bodyValue = body.value[0];
        if (!bodyValue) {
          return c.json({ error: "No body value" }, 400);
        }
        const clientState = bodyValue.clientState;
        if (clientState !== c.env.MICROSOFT_WEBHOOK_SECRET) {
          return c.json({ error: "Invalid client state" }, 401);
        }
        const resourceId = bodyValue.resourceData?.id;
        if (!resourceId) {
          return c.json({ error: "No resource id" }, 400);
        }
        // name for the email, so the agent can use it to identify the email account
        // from header
        const name = c.req.header("x-mcp-name") ?? "outlook";

        const respData = {
          name,
          emailId: resourceId,
        };

        const prompt: WebhookResponse = {
          reqResponseCode: 202,
          reqResponseContent: JSON.stringify({ ok: true }),
          reqResponseContentType: "json",
          promptContent: `Outlook email received:\n\n\`\`\`json\n${JSON.stringify(
            respData,
            null,
            2
          )}\n\`\`\``,
        };

        return c.json(prompt);
      })

      // Lifecycle notifications (e.g., reauthorizationRequired, subscriptionRemoved)
      .post("/email-lifecycle", async (c) => {
        const url = new URL(c.req.url);
        const validationToken = url.searchParams.get("validationToken");
        // Validation challenge from Microsoft Graph
        if (validationToken) {
          const response: WebhookResponse = {
            reqResponseCode: 200,
            reqResponseContent: validationToken,
            reqResponseContentType: "text",
          };

          return c.json(response);
        }
        const body = await c.req.json();
        const bodyValue = body.value[0];
        if (!bodyValue) {
          return c.json({ error: "No body value" }, 400);
        }
        const clientState = bodyValue.clientState;
        if (clientState !== c.env.MICROSOFT_WEBHOOK_SECRET) {
          return c.json({ error: "Invalid client state" }, 401);
        }
        // types of events:
        // - subscriptionRenewalRequired (need to refresh the subscription)
        // - missedNotifications (sign to run delta, but ok to ignore here for now)
        // - subscriptionRemoved (create new subscription)
        // - reauthorizationRequired (need to re-auth with oauth, leave this up to the UI)
        // no need to respond with any prompt info to the agent

        // get the subscription id, the event type, and the name
        const subscriptionId = bodyValue.subscriptionId;
        const eventType = bodyValue.lifecycleEvent;
        const name = c.req.header("x-mcp-name") ?? "outlook";

        const respData = {
          name,
          eventType,
          subscriptionId,
        };

        const response: WebhookResponse = {
          reqResponseCode: 202,
          reqResponseContent: JSON.stringify({ ok: true }),
          reqResponseContentType: "json",
          promptContent: `Outlook email lifecycle notification received:\n\n\`\`\`json\n${JSON.stringify(
            respData,
            null,
            2
          )}\n\`\`\``,
        };

        return c.json(response);
      })
  )

  // Health check endpoint
  .get("/", (c) => c.text("Microsoft MCP Server is running"));
