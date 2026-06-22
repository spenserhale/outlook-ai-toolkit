import { z } from "zod";

// Output + body format enums
export const BodyFormatSchema = z.enum(["text", "markdown", "html"]);
export type BodyFormat = z.infer<typeof BodyFormatSchema>;

export const ListBodyModeSchema = z.enum(["none", "preview", "full"]);
export type ListBodyMode = z.infer<typeof ListBodyModeSchema>;

export const OutputFormatSchema = z.enum(["toon", "json"]);
export type OutputFormat = z.infer<typeof OutputFormatSchema>;

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
  contentType: z.string(),
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
  body: ListBodyModeSchema.default("preview"),
  bodyFormat: BodyFormatSchema.default("text"),
});
export type ListMailParams = z.infer<typeof ListMailParamsSchema>;

export const GetMailParamsSchema = z.object({
  bodyFormat: BodyFormatSchema.default("text"),
});
export type GetMailParams = z.infer<typeof GetMailParamsSchema>;

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

// Graph API — messageRules (inbox rules)
export const MessageRulePredicatesSchema = z
  .object({
    senderContains: z.array(z.string()).optional(),
    subjectContains: z.array(z.string()).optional(),
    bodyContains: z.array(z.string()).optional(),
    fromAddresses: z.array(RecipientSchema).optional(),
  })
  .passthrough();
export type MessageRulePredicates = z.infer<typeof MessageRulePredicatesSchema>;

export const MessageRuleActionsSchema = z
  .object({
    moveToFolder: z.string().optional(),
    delete: z.boolean().optional(),
    markAsRead: z.boolean().optional(),
    forwardTo: z.array(RecipientSchema).optional(),
    stopProcessingRules: z.boolean().optional(),
  })
  .passthrough();
export type MessageRuleActions = z.infer<typeof MessageRuleActionsSchema>;

export const MessageRuleSchema = z
  .object({
    id: z.string(),
    displayName: z.string().optional(),
    sequence: z.number().optional(),
    isEnabled: z.boolean().optional(),
    isReadOnly: z.boolean().optional(),
    hasError: z.boolean().optional(),
    conditions: MessageRulePredicatesSchema.optional(),
    actions: MessageRuleActionsSchema.optional(),
    exceptions: MessageRulePredicatesSchema.optional(),
  })
  .passthrough();
export type MessageRule = z.infer<typeof MessageRuleSchema>;

export const MessageRuleListResponseSchema = z.object({
  value: z.array(MessageRuleSchema),
});
export type MessageRuleListResponse = z.infer<typeof MessageRuleListResponseSchema>;

export const CreateMessageRuleParamsSchema = z.object({
  displayName: z.string().min(1),
  sequence: z.number().int().default(1),
  isEnabled: z.boolean().default(true),
  conditions: MessageRulePredicatesSchema.optional(),
  actions: MessageRuleActionsSchema,
  exceptions: MessageRulePredicatesSchema.optional(),
});
export type CreateMessageRuleParams = z.infer<typeof CreateMessageRuleParamsSchema>;

export const UpdateMessageRuleParamsSchema = z
  .object({
    displayName: z.string(),
    sequence: z.number().int(),
    isEnabled: z.boolean(),
    conditions: MessageRulePredicatesSchema,
    actions: MessageRuleActionsSchema,
    exceptions: MessageRulePredicatesSchema,
  })
  .partial();
export type UpdateMessageRuleParams = z.infer<typeof UpdateMessageRuleParamsSchema>;

// Profile storage
export const ProfileSchema = z.object({
  clientId: z.string().min(1),
  tenantId: z.string().min(1),
});
export type Profile = z.infer<typeof ProfileSchema>;

export const ProfilesFileSchema = z.record(z.string(), ProfileSchema);
export type ProfilesFile = z.infer<typeof ProfilesFileSchema>;
