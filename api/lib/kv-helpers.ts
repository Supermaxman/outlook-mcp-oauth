export const getEventCache = (
  env: Env,
  serverName: string,
  eventType: string,
  eventId: string
) => env.MICROSOFT_EVENT_CACHE.get(`${serverName}:${eventType}:${eventId}`);

export const putEventCache = (
  env: Env,
  serverName: string,
  eventType: string,
  eventId: string,
  eventData: string,
  // default to 1 minute
  expirationTtl?: number | null
) =>
  env.MICROSOFT_EVENT_CACHE.put(
    `${serverName}:${eventType}:${eventId}`,
    eventData,
    {
      // if undefined, default to 1 minute
      // if null, don't set an expiration (explicitly no expiration)
      // if a number, use that as the expiration ttl
      expirationTtl:
        expirationTtl === undefined
          ? 60
          : expirationTtl === null
          ? undefined
          : expirationTtl,
    }
  );
