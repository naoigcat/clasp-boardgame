function getToken(): string | null {
  return PropertiesService.getScriptProperties().getProperty('TOKEN');
}

function getBodogeUserId(): string | null {
  return PropertiesService.getScriptProperties().getProperty('BODOGE_USER_ID');
}

function urlFetch(
  url: string,
  params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {},
): GoogleAppsScript.URL_Fetch.HTTPResponse {
  return UrlFetchApp.fetch(url, params);
}

function urlFetchWithAuth(
  url: string,
  params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {},
): GoogleAppsScript.URL_Fetch.HTTPResponse {
  const token = getToken();
  const headers: { [key: string]: string } = {
    ...(params.headers as { [key: string]: string } | undefined),
  };
  if (token) {
    headers['Authorization'] = `Bearer ${token}`;
  }
  return UrlFetchApp.fetch(url, { ...params, headers });
}
