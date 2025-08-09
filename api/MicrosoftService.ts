/* eslint-disable @typescript-eslint/no-explicit-any */
import { decodeJwt } from "jose";
import {
  CalendarEventSchema,
  CalendarEventStrict,
  EmailSchema,
  EmailMessageSchema,
} from "./lib/microsoft-types";

export interface DateTimeWithZone {
  /** ISO-8601 date-time string, e.g. "2025-08-05T14:00:00.0000000" */
  dateTime: string;
  /** IANA or Windows time-zone identifier, e.g. "UTC" or "Central Standard Time" */
  timeZone: string;
}

export interface EmailAddress {
  name: string;
  address: string;
}

export interface ResponseStatus {
  response:
    | "none"
    | "organizer"
    | "accepted"
    | "tentativelyAccepted"
    | "declined"
    | "notResponded"
    | string;
  /** ISO-8601 date-time string */
  time: string;
}

export interface ItemBody {
  contentType: "html" | "text";
  content: string;
}

export interface PhysicalAddress {
  street: string;
  city: string;
  state: string;
  countryOrRegion: string;
  postalCode: string;
}

export interface Coordinates {
  latitude: number;
  longitude: number;
  altitude?: number;
  accuracy?: number;
  altitudeAccuracy?: number;
}

/* ---------- complex helpers ---------- */

export interface Location {
  displayName: string;
  locationUri?: string;
  locationType?:
    | "default"
    | "conferenceRoom"
    | "homeAddress"
    | "businessAddress"
    | "geoCoordinates"
    | "streetAddress"
    | "hotel"
    | "restaurant"
    | "localBusiness"
    | "postalAddress"
    | string;
  uniqueId?: string;
  uniqueIdType?:
    | "unknown"
    | "locationStore"
    | "directory"
    | "private"
    | "bing"
    | string;
  address?: PhysicalAddress;
  coordinates?: Coordinates;
}

export interface RecurrencePattern {
  type:
    | "daily"
    | "weekly"
    | "absoluteMonthly"
    | "relativeMonthly"
    | "absoluteYearly"
    | "relativeYearly";
  interval: number;
  month: number;
  dayOfMonth: number;
  daysOfWeek?: (
    | "sunday"
    | "monday"
    | "tuesday"
    | "wednesday"
    | "thursday"
    | "friday"
    | "saturday"
  )[];
  firstDayOfWeek?:
    | "sunday"
    | "monday"
    | "tuesday"
    | "wednesday"
    | "thursday"
    | "friday"
    | "saturday";
  index?: "first" | "second" | "third" | "fourth" | "last";
}

export interface RecurrenceRange {
  type: "noEnd" | "endDate" | "numbered";
  /** YYYY-MM-DD */
  startDate: string;
  /** YYYY-MM-DD */
  endDate?: string;
  recurrenceTimeZone?: string;
  numberOfOccurrences?: number;
}

export interface Recurrence {
  pattern: RecurrencePattern;
  range: RecurrenceRange;
}

export interface Attendee {
  emailAddress: EmailAddress;
  type?: "required" | "optional" | "resource" | string;
  status?: ResponseStatus;
}

/* ---------- main event ---------- */

export interface CalendarEvent {
  /* identifiers & timestamps */
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;

  /* classification & misc. flags */
  categories: string[];
  originalStartTimeZone: string;
  originalEndTimeZone: string;
  reminderMinutesBeforeStart: number;
  isReminderOn: boolean;

  /* core content */
  subject: string;
  bodyPreview: string;
  body: ItemBody;

  importance: "low" | "normal" | "high";
  sensitivity: "normal" | "personal" | "private" | "confidential";

  isAllDay: boolean;
  isCancelled: boolean;
  isOrganizer: boolean;
  responseRequested: boolean;
  seriesMasterId: string | null;

  showAs:
    | "free"
    | "tentative"
    | "busy"
    | "oof"
    | "workingElsewhere"
    | "unknown"
    | string;
  type: "singleInstance" | "occurrence" | "exception" | "seriesMaster" | string;

  /* links & online-meeting info */
  webLink: string;
  onlineMeetingUrl: string;
  isOnlineMeeting: boolean;
  onlineMeetingProvider:
    | "unknown"
    | "teamsForBusiness"
    | "skypeForBusiness"
    | "skypeForConsumer"
    | string;

  allowNewTimeProposals: boolean;
  occurrenceId: string | null;
  isDraft: boolean;
  hideAttendees: boolean;

  /* composite objects */
  responseStatus: ResponseStatus;
  start: DateTimeWithZone;
  end: DateTimeWithZone;

  location: Location;
  locations: Location[];

  recurrence?: Recurrence;

  attendees: Attendee[];
  organizer: { emailAddress: EmailAddress };
}

type ODataPage<T> = {
  value: T[];
  "@odata.nextLink"?: string;
};

export class MicrosoftService {
  private env: Env;
  private accessToken: string;
  private refreshToken: string;
  private baseUrl = "https://graph.microsoft.com/v1.0";
  private userId: string;
  private webhookSecret: string;

  constructor(env: Env, accessToken: string, refreshToken: string) {
    this.env = env;
    this.accessToken = accessToken;
    this.refreshToken = refreshToken;
    this.userId = this.extractUserId(accessToken);
    this.webhookSecret = env.MICROSOFT_WEBHOOK_SECRET;
  }

  private async makeRequest<T>(
    url: string,
    options: RequestInit = {}
  ): Promise<T> {
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
        const errorText = await response.text().catch(() => "");
        throw new Error(
          `Microsoft API error: ${response.status} ${response.statusText}${
            errorText ? ` - ${errorText}` : ""
          }`
        );
      }

      const data = await response.json();
      return data as T;
    } catch (error) {
      console.error("Microsoft API request failed:", error);
      throw error;
    }
  }

  private async makeRequestIgnoreResponse(
    url: string,
    options: RequestInit = {}
  ): Promise<void> {
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
        const errorText = await response.text().catch(() => "");
        throw new Error(
          `Microsoft API error: ${response.status} ${response.statusText}${
            errorText ? ` - ${errorText}` : ""
          }`
        );
      }
    } catch (error) {
      console.error("Microsoft API request failed:", error);
      throw error;
    }
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
    endDate: string
  ): Promise<CalendarEventStrict[]> {
    // const params = {
    //   // TODO support finding recurring events, not sure if it find them by default
    //   $filter: `start/dateTime lt '${endDate}' and end/dateTime ge '${startDate}'`,
    //   $orderby: "start/dateTime desc",
    //   $top: limit.toString(),
    // };
    // const initialUrl = `${this.baseUrl}/users/${
    //   this.userId
    // }/calendarView?${new URLSearchParams(params).toString()}`;

    const params = {
      startDateTime: startDate,
      endDateTime: endDate,
    };

    const initialUrl = `${this.baseUrl}/users/${
      this.userId
    }/calendarView?${new URLSearchParams(params).toString()}`;

    let url: string | undefined = initialUrl;
    const events: CalendarEventStrict[] = [];

    while (url) {
      // tell makeRequest what the payload looks like:
      const page = await this.makeRequest<ODataPage<unknown>>(url, {
        method: "GET",
        // headers: {
        //   // TODO make this dynamic and configurable
        //   Prefer: 'outlook.timezone="Central Standard Time"',
        // },
      });

      // ⚡ validate every raw event – throws if anything is “illegal”
      events.push(...page.value.map((raw) => CalendarEventSchema.parse(raw)));

      url = page["@odata.nextLink"] as string | undefined;
    }
    // consider normalizing times to Date objects here, if you like
    return events;
  }

  async createCalendarEvent(
    subject: string,
    startDate: string,
    endDate: string,
    reminderMinutesBeforeStart: number,
    body?: string,
    location?: string,
    isAllDay?: boolean,
    categories?: string[],
    attendees?: string[]
  ) {
    const eventData = {
      subject,
      start: { dateTime: startDate, timeZone: "UTC" },
      end: { dateTime: endDate, timeZone: "UTC" },
      reminderMinutesBeforeStart,
      body: body ? { contentType: "text", content: body } : undefined,
      location: location ? { displayName: location } : undefined,
      isAllDay: isAllDay ?? false,
      categories: categories ?? undefined,
      attendees:
        attendees?.map((email) => ({ emailAddress: { address: email } })) ??
        undefined,
    };

    const eventRaw = await this.makeRequest<unknown>(
      `${this.baseUrl}/users/${this.userId}/events`,
      {
        method: "POST",
        body: JSON.stringify(eventData),
      }
    );
    const event = CalendarEventSchema.parse(eventRaw);
    return event;
  }

  async deleteCalendarEvent(eventId: string) {
    await this.makeRequestIgnoreResponse(
      `${this.baseUrl}/users/${this.userId}/events/${eventId}`,
      {
        method: "DELETE",
      }
    );
  }

  async getCalendarEvent(eventId: string) {
    const eventRaw = await this.makeRequest<unknown>(
      `${this.baseUrl}/users/${this.userId}/events/${eventId}`,
      {
        method: "GET",
      }
    );
    const event = CalendarEventSchema.parse(eventRaw);
    return event;
  }

  // allow for partial updates, but don't allow for removing fields, only updating them.
  // TODO: Consider null for fields that should be removed.
  async updateCalendarEvent(
    eventId: string,
    subject?: string,
    startDate?: string,
    endDate?: string,
    reminderMinutesBeforeStart?: number,
    body?: string,
    location?: string,
    isAllDay?: boolean,
    categories?: string[],
    attendees?: string[]
  ) {
    const eventData = {
      subject: subject ?? undefined,
      start: startDate ? { dateTime: startDate, timeZone: "UTC" } : undefined,
      end: endDate ? { dateTime: endDate, timeZone: "UTC" } : undefined,
      reminderMinutesBeforeStart:
        reminderMinutesBeforeStart !== undefined
          ? reminderMinutesBeforeStart
          : undefined,
      body: body ? { contentType: "text", content: body } : undefined,
      location: location ? { displayName: location } : undefined,
      isAllDay: isAllDay !== undefined ? isAllDay : undefined,
      categories: categories !== undefined ? categories : undefined,
      attendees:
        attendees !== undefined
          ? attendees.map((email) => ({ emailAddress: { address: email } }))
          : undefined,
    };

    const eventRaw = await this.makeRequest<unknown>(
      `${this.baseUrl}/users/${this.userId}/events/${eventId}`,
      {
        method: "PATCH",
        body: JSON.stringify(eventData),
      }
    );
    const event = CalendarEventSchema.parse(eventRaw);
    return event;
  }

  async getInboxEmails() {
    const emailsRaw = await this.makeRequest<unknown>(
      `${this.baseUrl}/me/mailFolders('Inbox')/messages`,
      {
        method: "GET",
      }
    );
    const emails = EmailSchema.parse(emailsRaw);
    return emails;
  }

  async getEmail(emailId: string) {
    const emailRaw = await this.makeRequest<unknown>(
      `${this.baseUrl}/me/messages/${emailId}`,
      {
        method: "GET",
      }
    );
    const email = EmailMessageSchema.parse(emailRaw);
    return email;
  }

  async draftEmail(
    subject: string,
    body: string,
    toRecipients: string[],
    ccRecipients?: string[],
    bccRecipients?: string[]
  ) {
    const draftData = {
      subject,
      body: { contentType: "text" as const, content: body },
      toRecipients: toRecipients.map((address) => ({
        emailAddress: { address },
      })),
      ccRecipients:
        ccRecipients && ccRecipients.length > 0
          ? ccRecipients.map((address) => ({ emailAddress: { address } }))
          : undefined,
      bccRecipients:
        bccRecipients && bccRecipients.length > 0
          ? bccRecipients.map((address) => ({ emailAddress: { address } }))
          : undefined,
    };

    const emailRaw = await this.makeRequest<unknown>(
      `${this.baseUrl}/me/messages`,
      {
        method: "POST",
        body: JSON.stringify(draftData),
      }
    );
    const email = EmailMessageSchema.parse(emailRaw);
    return email;
  }

  async createSubscription() {
    // max allowed for message subscriptions is 4230 minutes (~2.94 days)
    const expirationMinutes = 4230;
    const expirationDateTime = new Date(
      Date.now() + expirationMinutes * 60 * 1000
    );
    await this.makeRequest<unknown>(`${this.baseUrl}/subscriptions`, {
      method: "POST",
      body: JSON.stringify({
        changeType: "created",
        clientState: this.webhookSecret,
        // TODO: make this dynamic and configurable
        notificationUrl:
          "https://chat.aiescape.io/api/webhooks/outlook-email-notify",
        // TODO add this so we can refresh the subscription
        // lifecycleNotificationUrl: "",
        resource: "/me/mailFolders('Inbox')/messages",
        expirationDateTime: expirationDateTime.toISOString(),
        // TODO: add this so we can get the resource data
        // includeResourceData: true,
      }),
    });
  }
}
