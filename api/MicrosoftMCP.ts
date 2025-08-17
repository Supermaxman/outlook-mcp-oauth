import { McpAgent } from "agents/mcp";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { MicrosoftService } from "./MicrosoftService.ts";
import { MicrosoftAuthContext } from "../types";

/**
 * The `MicrosoftMCP` class exposes the Microsoft's Outlook API via the Model Context Protocol
 * for consumption by API Agents
 */
export class MicrosoftMCP extends McpAgent<Env, unknown, MicrosoftAuthContext> {
  async init() {
    // Initialize any necessary state
  }

  get microsoftService() {
    return new MicrosoftService(this.env, this.props.accessToken);
  }

  formatResponse = (
    description: string,
    data: unknown
  ): {
    content: Array<{ type: "text"; text: string }>;
  } => {
    return {
      content: [
        {
          type: "text",
          text: `Success! ${description}\n\nResult:\n${JSON.stringify(
            data,
            null,
            2
          )}`,
        },
      ],
    };
  };

  get server() {
    const server = new McpServer(
      {
        name: "Microsoft Service",
        description: "Microsoft MCP Server for Outlook",
        version: "1.0.0",
      },
      {
        instructions:
          "This MCP server is for the Microsoft's Outlook API. It can be used to get the user's calendar events, send emails, and more.",
      }
    );

    server.tool(
      "getUserCalendarEvents",
      "Get the user's calendar events",
      {
        startDate: z
          .string()
          .describe("Start date for the events in ISO 8601 format"),
        endDate: z
          .string()
          .describe("End date for the events in ISO 8601 format"),
      },
      async ({ startDate, endDate }) => {
        const events = await this.microsoftService.getUserCalendarEvents(
          startDate,
          endDate
        );
        return this.formatResponse("Calendar events retrieved", events);
      }
    );

    server.tool(
      "createCalendarEvent",
      "Create a new calendar event for the user",
      {
        subject: z.string().describe("The subject of the event"),
        startDate: z
          .string()
          .describe("The start date of the event in ISO 8601 format"),
        endDate: z
          .string()
          .describe("The end date of the event in ISO 8601 format"),
        reminderMinutesBeforeStart: z
          .number()
          .default(15)
          .describe(
            "The number of minutes before the event start to send a reminder"
          ),
        body: z
          .string()
          .optional()
          .describe("The body of the event, in text format"),
        location: z
          .string()
          .optional()
          .describe("The location of the event (or meeting link)"),
        isAllDay: z
          .boolean()
          .optional()
          .describe("Whether the event is all day (default: false)"),
        categories: z
          .array(z.string())
          .optional()
          .describe("The categories of the event (default: no categories)"),
        attendees: z
          .array(z.string())
          .optional()
          // TODO allow required and optional attendees
          .describe(
            "The email addresses of the attendees of the event (default: just the user)"
          ),
      },
      async ({
        subject,
        startDate,
        endDate,
        reminderMinutesBeforeStart,
        body,
        location,
        isAllDay,
        categories,
        attendees,
      }) => {
        const event = await this.microsoftService.createCalendarEvent(
          subject,
          startDate,
          endDate,
          reminderMinutesBeforeStart,
          body,
          location,
          isAllDay,
          categories,
          attendees
        );
        return this.formatResponse("Calendar event created", event);
      }
    );

    server.tool(
      "deleteCalendarEvent",
      "Delete a calendar event for the user",
      {
        eventId: z.string().describe("The ID of the event to delete"),
      },
      async ({ eventId }) => {
        await this.microsoftService.deleteCalendarEvent(eventId);
        return this.formatResponse("Calendar event deleted", {
          eventId,
        });
      }
    );

    server.tool(
      "getCalendarEvent",
      "Get a calendar event for the user",
      {
        eventId: z.string().describe("The ID of the event to get"),
      },
      async ({ eventId }) => {
        const event = await this.microsoftService.getCalendarEvent(eventId);
        return this.formatResponse("Calendar event retrieved", event);
      }
    );

    server.tool(
      "updateCalendarEvent",
      "Update a calendar event for the user. Only provided fields will be updated, other fields will be left unchanged.",
      {
        eventId: z.string().describe("The ID of the event to update"),
        subject: z.string().optional().describe("The subject of the event"),
        startDate: z
          .string()
          .optional()
          .describe("The start date of the event in ISO 8601 format"),
        endDate: z
          .string()
          .optional()
          .describe("The end date of the event in ISO 8601 format"),
        reminderMinutesBeforeStart: z
          .number()
          .optional()
          .describe(
            "The number of minutes before the event start to send a reminder"
          ),
        body: z
          .string()
          .optional()
          .describe("The body of the event, in text format"),
        location: z
          .string()
          .optional()
          .describe("The location of the event (or meeting link)"),
        isAllDay: z
          .boolean()
          .optional()
          .describe("Whether the event is all day (default: false)"),
        categories: z
          .array(z.string())
          .optional()
          .describe("The categories of the event (default: no categories)"),
        attendees: z
          .array(z.string())
          .optional()
          // TODO allow required and optional attendees
          .describe(
            "The email addresses of the attendees of the event (default: just the user)"
          ),
      },
      async ({
        eventId,
        subject,
        startDate,
        endDate,
        reminderMinutesBeforeStart,
        body,
        location,
        isAllDay,
        categories,
        attendees,
      }) => {
        const event = await this.microsoftService.updateCalendarEvent(
          eventId,
          subject,
          startDate,
          endDate,
          reminderMinutesBeforeStart,
          body,
          location,
          isAllDay,
          categories,
          attendees
        );
        return this.formatResponse("Calendar event updated", event);
      }
    );

    server.tool(
      "searchEmails",
      "Search emails in a folder by date range with optional filters. Uses $search when a free-text query is provided, otherwise uses server-side filters and client-side refinement.",
      {
        folder: z
          .enum(["inbox", "sentitems", "drafts", "archive"])
          .default("inbox")
          .describe(
            "Folder to search: 'inbox', 'sentitems', 'drafts', or 'archive'"
          ),
        startDate: z
          .string()
          .describe("Start of the date range in ISO 8601 format"),
        endDate: z
          .string()
          .describe("End of the date range in ISO 8601 format"),
        fromAddress: z
          .string()
          .optional()
          .describe("Filter by sender email address (contains match)"),
        toAddress: z
          .string()
          .optional()
          .describe("Filter by recipient email address (contains match)"),
        conversationId: z
          .string()
          .optional()
          .describe("Filter by exact conversation ID"),
        query: z
          .string()
          .optional()
          .describe(
            "Free-text search query. When set, server uses $search (no $filter/$orderby)."
          ),
      },
      async ({
        folder,
        startDate,
        endDate,
        fromAddress,
        toAddress,
        conversationId,
        query,
      }) => {
        const emails = await this.microsoftService.searchEmails(
          folder,
          startDate,
          endDate,
          fromAddress,
          toAddress,
          conversationId,
          query
        );
        return this.formatResponse("Emails retrieved", emails);
      }
    );

    server.tool(
      "markEmailAsRead",
      "Mark an email as read by ID.",
      {
        emailId: z.string().describe("The ID of the email to mark read"),
      },
      async ({ emailId }) => {
        await this.microsoftService.markEmailAsRead(emailId);
        return this.formatResponse("Email marked as read", { emailId });
      }
    );

    server.tool(
      "archiveEmail",
      "Move an email to the Archive folder.",
      {
        emailId: z.string().describe("The ID of the email to archive"),
      },
      async ({ emailId }) => {
        const moved = await this.microsoftService.archiveEmail(emailId);
        return this.formatResponse("Email archived", moved);
      }
    );

    server.tool(
      "getEmail",
      "Get a single email by its ID",
      {
        emailId: z.string().describe("The ID of the email to retrieve"),
      },
      async ({ emailId }) => {
        const email = await this.microsoftService.getEmail(emailId);
        return this.formatResponse("Email retrieved", email);
      }
    );

    server.tool(
      "draftEmail",
      "Create a draft email with recipients, subject, and plaintext body",
      {
        subject: z.string().describe("The subject of the draft email"),
        body: z
          .string()
          .describe("The body of the draft email in plaintext (no HTML)"),
        toRecipients: z
          .array(z.string())
          .min(1)
          .describe("List of recipient email addresses for the To field"),
        ccRecipients: z
          .array(z.string())
          .optional()
          .describe("Optional list of CC recipient email addresses"),
        bccRecipients: z
          .array(z.string())
          .optional()
          .describe("Optional list of BCC recipient email addresses"),
      },
      async ({ subject, body, toRecipients, ccRecipients, bccRecipients }) => {
        const draft = await this.microsoftService.draftEmail(
          subject,
          body,
          toRecipients,
          ccRecipients,
          bccRecipients
        );
        return this.formatResponse("Draft email created", draft);
      }
    );

    server.tool(
      "createReplyDraft",
      "Create a reply (or reply-all) draft to an existing email. This ensures proper threading in Outlook.",
      {
        originalEmailId: z
          .string()
          .describe("The ID of the original email to reply to"),
        replyAll: z
          .boolean()
          .optional()
          .default(false)
          .describe("Whether to reply to all recipients or just the sender"),
        body: z
          .string()
          .optional()
          .describe(
            "Optional plaintext body to set on the draft after creating the reply"
          ),
      },
      async ({ originalEmailId, replyAll, body }) => {
        const draft = await this.microsoftService.createReplyDraft(
          originalEmailId,
          replyAll,
          body
        );
        return this.formatResponse("Reply draft created", draft);
      }
    );

    server.tool(
      "updateEmailDraft",
      "Update fields of an existing email draft (subject, body, recipients).",
      {
        emailId: z.string().describe("The ID of the draft email to update"),
        subject: z.string().optional().describe("New subject"),
        body: z.string().optional().describe("New plaintext body"),
        toRecipients: z
          .array(z.string())
          .optional()
          .describe("Replace the To recipients list"),
        ccRecipients: z
          .array(z.string())
          .optional()
          .describe("Replace the CC recipients list"),
        bccRecipients: z
          .array(z.string())
          .optional()
          .describe("Replace the BCC recipients list"),
      },
      async ({
        emailId,
        subject,
        body,
        toRecipients,
        ccRecipients,
        bccRecipients,
      }) => {
        const updated = await this.microsoftService.updateEmailDraft(
          emailId,
          subject,
          body,
          toRecipients,
          ccRecipients,
          bccRecipients
        );
        return this.formatResponse("Draft updated", updated);
      }
    );

    server.tool(
      "sendEmail",
      "Send a draft email by ID.",
      {
        emailId: z.string().describe("The ID of the draft email to send"),
      },
      async ({ emailId }) => {
        await this.microsoftService.sendEmail(emailId);
        return this.formatResponse("Email sent", { emailId });
      }
    );

    server.tool(
      "deleteEmail",
      "Delete an email by ID (draft or message).",
      {
        emailId: z.string().describe("The ID of the email to delete"),
      },
      async ({ emailId }) => {
        await this.microsoftService.deleteEmail(emailId);
        return this.formatResponse("Email deleted", { emailId });
      }
    );

    server.tool(
      "createSubscription",
      "Create a subscription to a resource and set up a webhook to receive notifications when the resource changes.",
      {
        serverName: z
          .string()
          .describe(
            "The name of the MCP server to create the subscription for. This is used to identify the server in the webhook URL."
          ),
        resource: z
          .enum(["email", "calendar"])
          .describe("The resource to subscribe to: 'email' or 'calendar'"),
      },
      async ({ serverName, resource }) => {
        if (resource === "email") {
          await this.microsoftService.createEmailSubscription(serverName);
        } else if (resource === "calendar") {
          await this.microsoftService.createCalendarSubscription(serverName);
        } else {
          throw new Error(`Invalid resource: ${resource}`);
        }
        return this.formatResponse("Subscription created", {
          success: true,
        });
      }
    );

    server.tool(
      "refreshSubscription",
      "Refresh a subscription to a resource to avoid it expiring.",
      {
        subscriptionId: z
          .string()
          .describe("The ID of the subscription to refresh"),
      },
      async ({ subscriptionId }) => {
        await this.microsoftService.refreshSubscription(subscriptionId);
        return this.formatResponse("Subscription refreshed", {
          success: true,
        });
      }
    );

    server.tool(
      "listSubscriptions",
      "List all subscriptions to resources.",
      {},
      async () => {
        const subscriptions = await this.microsoftService.listSubscriptions();
        return this.formatResponse("Subscriptions retrieved", subscriptions);
      }
    );

    server.tool(
      "deleteSubscription",
      "Delete a subscription to a resource.",
      {
        subscriptionId: z
          .string()
          .describe("The ID of the subscription to delete"),
      },
      async ({ subscriptionId }) => {
        await this.microsoftService.deleteSubscription(subscriptionId);
        return this.formatResponse("Subscription deleted", {
          success: true,
        });
      }
    );

    server.tool(
      "searchPeople",
      "Search for people in the user's contacts and emails. Returns a list of people with their name, email address, etc.",
      {
        query: z
          .string()
          .describe(
            "The query to search for. This can be a name, email address, or other identifier."
          ),
        top: z
          .number()
          .min(1)
          .max(20)
          .default(10)
          .optional()
          .describe("The number of people to return"),
      },
      async ({ query, top }) => {
        const people = await this.microsoftService.searchPeople(query, top);
        return this.formatResponse("People retrieved", people);
      }
    );

    return server;
  }
}
