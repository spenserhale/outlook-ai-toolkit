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

export class OutlookAuthError extends OutlookError {
  constructor(message = "Authentication failed. Check your API key.") {
    super(message, "AUTH_ERROR", 401);
    this.name = "OutlookAuthError";
  }
}

export class OutlookNotFoundError extends OutlookError {
  constructor(resource: string, id: string) {
    super(`${resource} with id "${id}" not found`, "NOT_FOUND", 404);
    this.name = "OutlookNotFoundError";
  }
}
