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

    return server;
  }
}
