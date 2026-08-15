/**
 * Shared HTTP client for external board-game services.
 *
 * Centralizing UrlFetchApp usage makes optional auth headers and future request
 * defaults consistent across BoardGameGeek, Board Game Arena, and Bodoge.
 */
class HttpClient {
  /**
   * Sends one request through the shared Apps Script error-handling boundary.
   *
   * Callers need response codes to decide whether an external snapshot is safe
   * to publish, so every request must keep HTTP exceptions muted consistently.
   */
  private static fetch(
    url: string,
    options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {},
  ): GoogleAppsScript.URL_Fetch.HTTPResponse {
    return UrlFetchApp.fetch(url, { ...options, muteHttpExceptions: true });
  }

  /**
   * Fetches a URL with the supplied Apps Script request options.
   *
   * Always mutes HTTP exceptions so callers can inspect response codes and apply
   * their own error handling or throttling. UrlFetchApp throws on non-2xx by
   * default, which would skip those branches and skip post-fetch delays.
   */
  static get(
    url: string,
    options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {},
  ): GoogleAppsScript.URL_Fetch.HTTPResponse {
    return HttpClient.fetch(url, options);
  }

  /**
   * Fetches a URL and adds a Bearer token when one has been configured.
   *
   * BoardGameGeek allows anonymous access but applies stricter rate limits.
   * Callers still work without a token; installations with one avoid those
   * anonymous ceilings without changing call sites.
   */
  static getWithOptionalBearerToken(
    url: string,
    token: string | null,
    options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {},
  ): GoogleAppsScript.URL_Fetch.HTTPResponse {
    const headers = {
      ...(options.headers as Record<string, string> | undefined),
    };

    if (token !== null && token.length > 0) {
      headers.Authorization = `Bearer ${token}`;
    }

    return HttpClient.fetch(url, {
      ...options,
      headers,
    });
  }
}
