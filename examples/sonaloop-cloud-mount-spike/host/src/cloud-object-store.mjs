export async function loadCloudObjectBytes({ presignGetUrl, objectKey, fetchImpl = globalThis.fetch }) {
  assertFetch(fetchImpl);
  const ticketUrl = new URL(presignGetUrl);
  ticketUrl.searchParams.set("key", objectKey);

  const ticketResponse = await fetchImpl(ticketUrl, { method: "GET" });
  assertOk(ticketResponse, "presigned GET ticket");
  const ticket = await ticketResponse.json();
  assertTicket(ticket, "presigned GET ticket");

  const objectResponse = await fetchImpl(ticket.url, {
    method: "GET",
    headers: ticket.headers ?? {},
  });
  assertOk(objectResponse, "cloud object GET");

  return {
    bytes: new Uint8Array(await objectResponse.arrayBuffer()),
    etag: objectResponse.headers.get("etag") ?? ticket.etag ?? null,
    filename:
      typeof ticket.filename === "string" && ticket.filename.length > 0
        ? ticket.filename
        : filenameFromKey(objectKey),
  };
}

export async function saveCloudObjectBytes({
  presignPutUrl,
  objectKey,
  bytes,
  mime,
  filename,
  etag,
  fetchImpl = globalThis.fetch,
}) {
  assertFetch(fetchImpl);
  const ticketResponse = await fetchImpl(presignPutUrl, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      key: objectKey,
      contentType: mime,
      filename,
      ifMatch: etag,
    }),
  });
  assertOk(ticketResponse, "presigned PUT ticket");
  const ticket = await ticketResponse.json();
  assertTicket(ticket, "presigned PUT ticket");

  const putResponse = await fetchImpl(ticket.url, {
    method: "PUT",
    headers: {
      ...(ticket.headers ?? {}),
      "content-type": mime,
      ...(etag ? { "if-match": etag } : {}),
    },
    body: bytes,
  });
  assertOk(putResponse, "cloud object PUT");

  return {
    etag: putResponse.headers.get("etag") ?? ticket.etag ?? null,
  };
}

function assertFetch(fetchImpl) {
  if (typeof fetchImpl !== "function") {
    throw new Error("A fetch implementation is required for cloud object access");
  }
}

function assertOk(response, label) {
  if (!response.ok) {
    throw new Error(`${label} failed with HTTP ${response.status}`);
  }
}

function assertTicket(ticket, label) {
  if (!ticket || typeof ticket !== "object" || typeof ticket.url !== "string" || ticket.url.length === 0) {
    throw new Error(`${label} did not return a usable URL`);
  }
  if (
    ticket.headers !== undefined &&
    (!ticket.headers || typeof ticket.headers !== "object" || Array.isArray(ticket.headers))
  ) {
    throw new Error(`${label} returned invalid headers`);
  }
}

function filenameFromKey(objectKey) {
  const last = objectKey.split("/").filter(Boolean).pop();
  return last && last.length > 0 ? last : "document.xlsx";
}
