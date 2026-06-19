export interface HttpRequestOptions {
  url: string;
  method: string;
  headers: Record<string, string>;
  body: string;
}

export function shouldUseXhrSse(): boolean {
  return typeof ReadableStream === "undefined";
}

function supportsResponseStreaming(response: Response): boolean {
  return Boolean(response.body && typeof response.body.getReader === "function");
}

async function* iterateSseLinesFromXhr(options: HttpRequestOptions): AsyncIterable<string> {
  var lineQueue: string[] = [];
  var done = false;
  var errorMsg: string | null = null;
  var wake: (() => void) | null = null;

  function poke(): void {
    if (wake) {
      var resolve = wake;
      wake = null;
      resolve();
    }
  }

  var xhr = new XMLHttpRequest();
  var lastIndex = 0;
  var pending = "";

  function processNewText(): void {
    var text = xhr.responseText || "";
    var chunk = text.substring(lastIndex);
    lastIndex = text.length;
    pending += chunk;

    var newlineIndex = pending.indexOf("\n");
    while (newlineIndex >= 0) {
      var line = pending.substring(0, newlineIndex);
      pending = pending.substring(newlineIndex + 1);
      lineQueue.push(line);
      newlineIndex = pending.indexOf("\n");
    }
    poke();
  }

  xhr.open(options.method, options.url, true);
  var headerKeys = Object.keys(options.headers);
  for (var i = 0; i < headerKeys.length; i++) {
    var key = headerKeys[i];
    xhr.setRequestHeader(key, options.headers[key]);
  }

  xhr.onprogress = processNewText;
  xhr.onload = function () {
    processNewText();
    if (pending.length > 0) {
      lineQueue.push(pending);
      pending = "";
    }

    if (xhr.status < 200 || xhr.status >= 300) {
      errorMsg = "HTTP " + xhr.status + ": " + (xhr.responseText || xhr.statusText);
    }
    done = true;
    poke();
  };
  xhr.onerror = function () {
    errorMsg = "Network request failed";
    done = true;
    poke();
  };
  xhr.send(options.body);

  while (true) {
    if (errorMsg) {
      throw new Error(errorMsg);
    }
    if (lineQueue.length > 0) {
      yield lineQueue.shift() as string;
      continue;
    }
    if (done) {
      break;
    }
    await new Promise<void>(function (resolve) {
      wake = resolve;
    });
  }
}

async function* iterateSseLinesFromFetch(options: HttpRequestOptions): AsyncIterable<string> {
  var response = await fetch(options.url, {
    method: options.method,
    headers: options.headers,
    body: options.body,
  });

  if (!response.ok) {
    var body = await response.text();
    throw new Error("HTTP " + response.status + ": " + body);
  }

  if (!supportsResponseStreaming(response)) {
    var fullText = await response.text();
    var lines = fullText.split("\n");
    for (var i = 0; i < lines.length; i++) {
      yield lines[i];
    }
    return;
  }

  var reader = response.body!.getReader();
  var decoder = typeof TextDecoder !== "undefined" ? new TextDecoder() : null;
  var buffer = "";

  try {
    while (true) {
      var chunk = await reader.read();
      if (chunk.done) {
        break;
      }

      if (decoder) {
        buffer += decoder.decode(chunk.value, { stream: true });
      } else {
        buffer += String.fromCharCode.apply(null, Array.from(chunk.value) as number[]);
      }

      var newlineAt = buffer.indexOf("\n");
      while (newlineAt >= 0) {
        yield buffer.substring(0, newlineAt);
        buffer = buffer.substring(newlineAt + 1);
        newlineAt = buffer.indexOf("\n");
      }
    }

    if (buffer.length > 0) {
      yield buffer;
    }
  } finally {
    reader.releaseLock();
  }
}

export async function* iterateSseLines(options: HttpRequestOptions): AsyncIterable<string> {
  if (shouldUseXhrSse()) {
    yield* iterateSseLinesFromXhr(options);
    return;
  }
  yield* iterateSseLinesFromFetch(options);
}