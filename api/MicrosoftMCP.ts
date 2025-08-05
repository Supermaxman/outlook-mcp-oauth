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
    return new MicrosoftService(
      this.env,
      this.props.accessToken,
      this.props.refreshToken
    );
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

    // TODO allow for partial updates
    server.tool(
      "updateCalendarEvent",
      "Update a calendar event for the user",
      {
        eventId: z.string().describe("The ID of the event to update"),
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

    return server;
  }
}
