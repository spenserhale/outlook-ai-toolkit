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
  SendMailParams,
  ReplyParams,
  DraftParams,
  Profile,
  ProfilesFile,
} from "./types.js";
export {
  OutlookConfigSchema,
  TokenDataSchema,
  AuthStatusSchema,
  MessageSchema,
  MailListResponseSchema,
  DeltaResponseSchema,
  ListMailParamsSchema,
  SendMailParamsSchema,
  ReplyParamsSchema,
  DraftParamsSchema,
  ProfileSchema,
  ProfilesFileSchema,
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
