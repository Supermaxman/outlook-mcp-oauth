/* eslint-disable @typescript-eslint/no-explicit-any */
import { decodeJwt } from "jose";

export class MicrosoftService {
  private env: Env;
  private accessToken: string;
  private refreshToken: string;
  private baseUrl = "https://graph.microsoft.com/v1.0";
  private userId: string;

  constructor(env: Env, accessToken: string, refreshToken: string) {
    this.env = env;
    this.accessToken = accessToken;
    this.refreshToken = refreshToken;
    this.userId = this.extractUserId(accessToken);
  }

  private async makeRequest(
    url: string,
    options: RequestInit = {}
  ): Promise<any> {
    try {
      const response = await fetch(url, {
        ...options,
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          "Content-Type": "application/json",
          ...options.headers,
        },
      });

      if (response.status === 401) {
        // Token expired, try to refresh
        await this.refreshAccessToken();

        // Retry the request with new token
        return fetch(url, {
          ...options,
          headers: {
            Authorization: `Bearer ${this.accessToken}`,
            "Content-Type": "application/json",
            ...options.headers,
          },
        }).then((res) => res.json());
      }

      if (!response.ok) {
        throw new Error(
          `Microsoft API error: ${response.status} ${response.statusText}`
        );
      }

      return response.json();
    } catch (error) {
      console.error("Microsoft API request failed:", error);
      throw error;
    }
  }

  private async makeEndpointRequest(
    endpoint: string,
    options: RequestInit = {}
  ): Promise<any> {
    const url = `${this.baseUrl}${endpoint}`;
    return this.makeRequest(url, options);
  }

  private async refreshAccessToken(): Promise<void> {
    const body = new URLSearchParams({
      client_id: this.env.MICROSOFT_CLIENT_ID,
      client_secret: this.env.MICROSOFT_CLIENT_SECRET,
      grant_type: "refresh_token",
      refresh_token: this.refreshToken,
      scope:
        "openid profile email offline_access Calendars.ReadWrite Mail.ReadWrite Mail.Send User.Read People.Read",
    });

    const res = await fetch(
      // use the same tenant you passed to the original exchange
      `https://login.microsoftonline.com/${this.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body,
      }
    );

    if (!res.ok) {
      throw new Error(`Failed to refresh access token: ${await res.text()}`);
    }

    const { access_token, refresh_token } = (await res.json()) as {
      access_token: string;
      refresh_token?: string;
    };

    this.accessToken = access_token;
    if (refresh_token) this.refreshToken = refresh_token;
  }

  extractUserId(accessToken: string): string {
    const decoded = decodeJwt(accessToken);
    return decoded.oid as string;
  }

  // Calendars
  async getUserCalendarEvents(
    startDate: string,
    endDate: string,
    limit: number = 100
  ): Promise<any> {
    const params = {
      // TODO support finding recurring events, not sure if it find them by default
      $filter: `start/dateTime lt '${endDate}' and end/dateTime ge '${startDate}'`,
      $orderby: "start/dateTime desc",
      $top: limit.toString(),
    };
    const initialUrl = `${this.baseUrl}/users/${
      this.userId
    }/events?${new URLSearchParams(params).toString()}`;
    let url: string | null = initialUrl;
    const events: any[] = [];
    while (url) {
      const data = await this.makeRequest(url, {
        method: "GET",
      });
      const newEvents = data["value"];
      // TODO consider some better formatting for the events
      if (newEvents) {
        events.push(...newEvents);
      }
      url = data["@odata.nextLink"] || null;
    }
    return events;
  }
}
