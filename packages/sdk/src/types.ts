import { z } from "zod";

// Config
export const OutlookConfigSchema = z.object({
  clientId: z.string().min(1, "OUTLOOK_CLIENT_ID is required"),
  tenantId: z.string().min(1, "OUTLOOK_TENANT_ID is required"),
});
export type OutlookConfig = z.infer<typeof OutlookConfigSchema>;

// Token storage
export const TokenDataSchema = z.object({
  accessToken: z.string(),
  refreshToken: z.string(),
  accessTokenExpiry: z.number(),
  refreshTokenExpiry: z.number(),
  userEmail: z.string().optional(),
});
export type TokenData = z.infer<typeof TokenDataSchema>;

// Auth status (non-sensitive, safe to display)
export const AuthStatusSchema = z.object({
  authenticated: z.boolean(),
  userEmail: z.string().optional(),
  accessTokenExpiry: z.number().optional(),
  refreshTokenExpiry: z.number().optional(),
});
export type AuthStatus = z.infer<typeof AuthStatusSchema>;

// Graph API — message types
export const MessageBodySchema = z.object({
  contentType: z.enum(["text", "HTML"]),
  content: z.string(),
});
export type MessageBody = z.infer<typeof MessageBodySchema>;

export const EmailAddressSchema = z.object({
  address: z.string(),
  name: z.string().optional(),
});
export type EmailAddress = z.infer<typeof EmailAddressSchema>;

export const RecipientSchema = z.object({
  emailAddress: EmailAddressSchema,
});
export type Recipient = z.infer<typeof RecipientSchema>;

export const MessageSchema = z
  .object({
    id: z.string(),
    subject: z.string().nullable().optional(),
    bodyPreview: z.string().optional(),
    body: MessageBodySchema.optional(),
    from: RecipientSchema.optional(),
    toRecipients: z.array(RecipientSchema).optional(),
    receivedDateTime: z.string().optional(),
    sentDateTime: z.string().optional(),
    isRead: z.boolean().optional(),
    isDraft: z.boolean().optional(),
    conversationId: z.string().optional(),
  })
  .passthrough();
export type Message = z.infer<typeof MessageSchema>;

export const MailListResponseSchema = z.object({
  value: z.array(MessageSchema),
  "@odata.nextLink": z.string().optional(),
});
export type MailListResponse = z.infer<typeof MailListResponseSchema>;

export const DeltaResponseSchema = z.object({
  value: z.array(MessageSchema),
  "@odata.nextLink": z.string().optional(),
  "@odata.deltaLink": z.string().optional(),
});
export type DeltaResponse = z.infer<typeof DeltaResponseSchema>;

// Mail operation params
export const ListMailParamsSchema = z.object({
  folder: z.string().default("inbox"),
  limit: z.number().int().positive().max(999).default(25),
  cursor: z.string().optional(),
  filter: z.string().optional(),
  select: z.string().optional(),
  orderby: z.string().optional(),
});
export type ListMailParams = z.infer<typeof ListMailParamsSchema>;

export const SendMailParamsSchema = z.object({
  to: z.string().email(),
  subject: z.string(),
  body: z.string(),
  contentType: z.enum(["text", "HTML"]).default("HTML"),
});
export type SendMailParams = z.infer<typeof SendMailParamsSchema>;

export const ReplyParamsSchema = z.object({
  body: z.string(),
  contentType: z.enum(["text", "HTML"]).default("HTML"),
});
export type ReplyParams = z.infer<typeof ReplyParamsSchema>;

export const DraftParamsSchema = z.object({
  to: z.string().email(),
  subject: z.string(),
  body: z.string(),
  contentType: z.enum(["text", "HTML"]).default("HTML"),
});
export type DraftParams = z.infer<typeof DraftParamsSchema>;

// Profile storage
export const ProfileSchema = z.object({
  clientId: z.string().min(1),
  tenantId: z.string().min(1),
});
export type Profile = z.infer<typeof ProfileSchema>;

export const ProfilesFileSchema = z.record(z.string(), ProfileSchema);
export type ProfilesFile = z.infer<typeof ProfilesFileSchema>;
