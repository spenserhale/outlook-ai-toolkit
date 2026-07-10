export class OutlookError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly statusCode?: number
  ) {
    super(message);
    this.name = "OutlookError";
  }
}

export class OutlookConfigError extends OutlookError {
  constructor(message: string) {
    super(message, "CONFIG_ERROR");
    this.name = "OutlookConfigError";
  }
}

export class OutlookAuthError extends OutlookError {
  constructor(message = "Not authenticated. Run `outlook login` first.") {
    super(message, "AUTH_ERROR", 401);
    this.name = "OutlookAuthError";
  }
}

export class OutlookNotFoundError extends OutlookError {
  constructor(resource: string, id: string) {
    super(`${resource} "${id}" not found`, "NOT_FOUND", 404);
    this.name = "OutlookNotFoundError";
  }
}

export class OutlookRateLimitError extends OutlookError {
  constructor(retryAfterSeconds?: number) {
    const msg = retryAfterSeconds
      ? `Rate limited by Microsoft Graph. Retry after ${retryAfterSeconds}s.`
      : "Rate limited by Microsoft Graph API.";
    super(msg, "RATE_LIMIT", 429);
    this.name = "OutlookRateLimitError";
  }
}
