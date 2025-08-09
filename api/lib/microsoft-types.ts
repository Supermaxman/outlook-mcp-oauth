/* --------------------------------- zod schema -------------------------------- */

import { z } from "zod";

export const CalendarEventSchema = z
  .object({
    /* ----------------------------------------------------
       Everything must appear here EXACTLY as in CalendarEvent.
       Optional properties => .optional(), nullable => z.null()
       ---------------------------------------------------- */
    id: z.string(),
    createdDateTime: z.string(),
    lastModifiedDateTime: z.string(),

    categories: z.array(z.string()),
    originalStartTimeZone: z.string(),
    originalEndTimeZone: z.string(),
    reminderMinutesBeforeStart: z.number(),
    isReminderOn: z.boolean(),

    subject: z.string(),
    bodyPreview: z.string(),
    body: z.object({
      contentType: z.union([z.literal("html"), z.literal("text")]),
      content: z.string(),
    }),

    importance: z.union([
      z.literal("low"),
      z.literal("normal"),
      z.literal("high"),
    ]),
    sensitivity: z.union([
      z.literal("normal"),
      z.literal("personal"),
      z.literal("private"),
      z.literal("confidential"),
    ]),

    isAllDay: z.boolean(),
    isCancelled: z.boolean(),
    isOrganizer: z.boolean(),
    responseRequested: z.boolean(),
    seriesMasterId: z.string().nullish(),

    showAs: z.string(), // keep broad for now
    type: z.string(),

    webLink: z.string().url(),
    onlineMeetingUrl: z.string().nullish(),
    isOnlineMeeting: z.boolean(),
    onlineMeetingProvider: z.string(),

    allowNewTimeProposals: z.boolean(),
    occurrenceId: z.string().nullish(),
    isDraft: z.boolean(),
    hideAttendees: z.boolean(),

    responseStatus: z.object({
      response: z.string(),
      time: z.string(),
    }),

    start: z.object({ dateTime: z.string(), timeZone: z.string() }),
    end: z.object({ dateTime: z.string(), timeZone: z.string() }),

    location: z.any(), // ←-- give Location its own schema exactly as above if you want
    locations: z.array(z.any()),
    recurrence: z.any().optional(),

    attendees: z.array(z.any()),
    organizer: z.object({ emailAddress: z.any() }),
  })
  .strip(); // <-- ✨ rejects *any* unknown keys

export type CalendarEventStrict = z.infer<typeof CalendarEventSchema>;

/* --------------------------------- email schema -------------------------------- */

const EmailAddressSchema = z
  .object({
    name: z.string().nullable().optional(),
    address: z.string().nullable().optional(),
  })
  .strip();

const RecipientSchema = z
  .object({
    emailAddress: EmailAddressSchema,
  })
  .strip();

const ItemBodySchema = z
  .object({
    contentType: z.union([z.literal("html"), z.literal("text")]).optional(),
    content: z.string().nullable().optional(),
  })
  .strip();

export const EmailMessageSchema = z
  .object({
    /* identifiers & timestamps */
    id: z.string(),
    createdDateTime: z.string().optional(),
    lastModifiedDateTime: z.string().optional(),
    changeKey: z.string().optional(),

    /* classification & flags */
    categories: z.array(z.string()).optional(),
    receivedDateTime: z.string().optional(),
    sentDateTime: z.string().optional(),
    hasAttachments: z.boolean().optional(),
    internetMessageId: z.string().nullable().optional(),
    importance: z
      .union([z.literal("low"), z.literal("normal"), z.literal("high")])
      .optional(),
    inferenceClassification: z.string().optional(),

    /* core content */
    subject: z.string().nullable().optional(),
    bodyPreview: z.string().optional(),
    body: ItemBodySchema.optional(),

    /* read state */
    isRead: z.boolean().optional(),
    isDraft: z.boolean().optional(),
    isReadReceiptRequested: z.boolean().optional(),
    isDeliveryReceiptRequested: z.boolean().optional(),

    /* web link */
    webLink: z.string().url().optional(),

    /* conversation & folder */
    parentFolderId: z.string().optional(),
    conversationId: z.string().optional(),

    /* participants */
    from: RecipientSchema.optional(),
    sender: RecipientSchema.optional(),
    toRecipients: z.array(RecipientSchema).optional(),
    ccRecipients: z.array(RecipientSchema).optional(),
    bccRecipients: z.array(RecipientSchema).optional(),
    replyTo: z.array(RecipientSchema).optional(),

    /* misc fields we do not strictly validate right now */
    flag: z.any().optional(),
    attachments: z.any().optional(),
    internetMessageHeaders: z.any().optional(),
  })
  .strip();

/* OData page wrapper for list messages endpoint */
export const EmailSchema = z
  .object({
    value: z.array(EmailMessageSchema),
    "@odata.nextLink": z.string().url().optional(),
  })
  .strip();
