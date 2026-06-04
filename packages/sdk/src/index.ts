export { OutlookClient } from "./client.js";
export { resolveConfig } from "./config.js";
export { OutlookError, OutlookAuthError, OutlookNotFoundError } from "./errors.js";
export type {
  OutlookConfig,
  Resource,
  ListResourcesParams,
  CreateResourceParams,
  PaginatedResponse,
  ErrorResponse,
} from "./types.js";
export {
  OutlookConfigSchema,
  ResourceSchema,
  ListResourcesParamsSchema,
  CreateResourceParamsSchema,
  ErrorResponseSchema,
} from "./types.js";
