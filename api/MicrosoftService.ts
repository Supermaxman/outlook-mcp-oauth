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
    endpoint: string,
    options: RequestInit = {}
  ): Promise<any> {
    const url = `${this.baseUrl}${endpoint}`;

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

  private async refreshAccessToken(): Promise<void> {
    const response = await fetch(
      `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
      {
        method: "POST",
        body: JSON.stringify({
          client_id: this.env.MICROSOFT_CLIENT_ID,
          client_secret: this.env.MICROSOFT_CLIENT_SECRET,
          refresh_token: this.refreshToken,
          grant_type: "refresh_token",
          scope:
            "openid profile email offline_access Calendars.ReadWrite Mail.ReadWrite Mail.Send User.Read People.Read",
        }),
      }
    );

    if (!response.ok) {
      throw new Error("Failed to refresh access token");
    }

    const data = (await response.json()) as {
      access_token: string;
      refresh_token?: string;
      expires_in: number;
    };
    this.accessToken = data.access_token;
    if (data.refresh_token) {
      this.refreshToken = data.refresh_token;
    }
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
    const url = `/users/${this.userId}/events?${new URLSearchParams(
      params
    ).toString()}`;
    return this.makeRequest(url, {
      method: "GET",
    });
  }
}
