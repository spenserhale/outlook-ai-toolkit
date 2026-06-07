// Config
export { resolveConfig, resolveAuthority, SCOPES, GRAPH_BASE_URL, AUTH_BASE_URL } from "./config.js";

// Types
export type {
  OutlookConfig,
  TokenData,
  AuthStatus,
  Message,
  MessageBody,
  EmailAddress,
  Recipient,
  MailListResponse,
  DeltaResponse,
  ListMailParams,
  GetMailParams,
  SendMailParams,
  ReplyParams,
  DraftParams,
  Profile,
  ProfilesFile,
  BodyFormat,
  ListBodyMode,
  OutputFormat,
} from "./types.js";
export {
  OutlookConfigSchema,
  TokenDataSchema,
  AuthStatusSchema,
  MessageSchema,
  MailListResponseSchema,
  DeltaResponseSchema,
  ListMailParamsSchema,
  GetMailParamsSchema,
  SendMailParamsSchema,
  ReplyParamsSchema,
  DraftParamsSchema,
  ProfileSchema,
  ProfilesFileSchema,
  BodyFormatSchema,
  ListBodyModeSchema,
  OutputFormatSchema,
} from "./types.js";

// Errors
export {
  OutlookError,
  OutlookConfigError,
  OutlookAuthError,
  OutlookNotFoundError,
  OutlookRateLimitError,
} from "./errors.js";

// Auth
export {
  OutlookAuth,
  generateCodeVerifier,
  generateCodeChallenge,
  generateState,
  buildAuthorizationUrl,
} from "./auth.js";

// Token + Profile storage
export { TokenStore } from "./token-store.js";
export { ProfileStore } from "./profile-store.js";

// Graph + Mail
export { GraphClient } from "./graph-client.js";
export type { ODataOptions, ODataListResponse } from "./graph-client.js";
export { MailClient } from "./mail-client.js";

// Output + body rendering
export { renderBody } from "./body.js";
export { renderOutput } from "./format.js";
