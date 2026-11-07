import http from "node:http";
import { readFile, stat } from "node:fs/promises";
import { resolve, extname, sep } from "node:path";
import { pathToFileURL } from "node:url";
import {
  aiProvider,
  askAI,
  isSharedOpenAI,
  openAIUserRef,
  validateRequest,
} from "./ai.mjs";
const mime = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript",
  ".css": "text/css",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
  ".json": "application/json",
};
export function createServer(config = process.env) {
  const root = resolve("dist"),
    limits = new Map();
  const sharedOpenAI = isSharedOpenAI(config);
  const publicOrigin = config.PUBLIC_ORIGIN?.trim();
  const configuredRate = Number(config.OPENAI_SHARED_RPM);
  const sharedRate =
    Number.isSafeInteger(configuredRate) && configuredRate > 0
      ? configuredRate
      : 30;
  let sharedWindow = { count: 0, reset: 0 };
  const json = (res, status, body) => {
    if (res.writableEnded || res.destroyed) return;
    res.writeHead(status, {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
    });
    res.end(JSON.stringify(body));
  };
  return http.createServer(
    { requestTimeout: 15000, headersTimeout: 10000, keepAliveTimeout: 5000 },
    async (req, res) => {
      res.setHeader("X-Content-Type-Options", "nosniff");
      res.setHeader("Referrer-Policy", "strict-origin-when-cross-origin");
      res.setHeader("X-Frame-Options", "SAMEORIGIN");
      try {
        const url = new URL(req.url, "http://localhost");
        const managed = config.CODERS_NATIVE === "true";
        // In native mode the Coders gate strips user-supplied identity before stamping this header.
        // A standalone/local server must never forward arbitrary client identity to the provider.
        const rawUser = req.headers["x-coders-user"];
        const user =
          managed &&
          typeof rawUser === "string" &&
          /^[a-zA-Z0-9_-]{1,128}$/.test(rawUser)
            ? rawUser
            : undefined;
        if (["GET", "HEAD"].includes(req.method) && url.pathname === "/healthz")
          return json(res, 200, { status: "ok", version: "0.7.0" });
        if (req.method === "GET" && url.pathname === "/api/status") {
          const provider = aiProvider(config, user);
          return json(res, 200, {
            aiAvailable: Boolean(provider),
            provider,
            signedIn: Boolean(user),
            managed,
            sharedOpenAI,
            openAIUserRef: sharedOpenAI ? null : openAIUserRef(user),
          });
        }
        if (url.pathname === "/api/ai" && req.method === "POST") {
          const origins = managed
            ? [publicOrigin || `https://${req.headers.host}`]
            : [
                "http://127.0.0.1:7730",
                "http://localhost:7730",
                `https://${req.headers.host}`,
                `http://${req.headers.host}`,
              ];
          if (req.headers.origin && !origins.includes(req.headers.origin))
            return json(res, 403, {
              error: "이 사이트에서 다시 요청해 주세요.",
            });
          if (!req.headers["content-type"]?.startsWith("application/json"))
            return json(res, 415, { error: "JSON 요청이 필요합니다." });
          if (managed && !user)
            return json(res, 401, {
              error: "AI를 사용하려면 Coders에 로그인해 주세요.",
            });
          const now = Date.now();
          for (const [k, value] of limits)
            if (value.reset < now) limits.delete(k);
          const identity = user || req.socket.remoteAddress;
          const limit = limits.get(identity) ?? {
            count: 0,
            reset: now + 60000,
          };
          if (
            limit.count >= 6 ||
            (limits.size >= 5000 && !limits.has(identity))
          ) {
            res.setHeader(
              "Retry-After",
              String(Math.max(1, Math.ceil((limit.reset - now) / 1000))),
            );
            return json(res, 429, {
              error: "요청이 많습니다. 1분 후 다시 시도해 주세요.",
            });
          }
          limit.count++;
          limits.set(identity, limit);
          // Decode UTF-8 once, after collection: a Korean character can span TCP chunks.
          let bytes = 0;
          const chunks = [];
          for await (const chunk of req.iterator({ destroyOnReturn: false })) {
            bytes += chunk.length;
            if (bytes > 65536) {
              res.setHeader("Connection", "close");
              req.resume();
              return json(res, 413, {
                error: "질문이나 선택 범위를 줄여 주세요.",
              });
            }
            chunks.push(chunk);
          }
          const body = Buffer.concat(chunks).toString("utf8");
          let input;
          try {
            input = validateRequest(JSON.parse(body));
          } catch (error) {
            return json(res, 400, {
              error:
                error instanceof SyntaxError
                  ? "JSON 형식을 확인해 주세요."
                  : error.message,
            });
          }
          if (sharedOpenAI) {
            const started = Date.now();
            if (started >= sharedWindow.reset)
              sharedWindow = { count: 0, reset: started + 60000 };
            if (sharedWindow.count >= sharedRate) {
              res.setHeader(
                "Retry-After",
                String(Math.ceil((sharedWindow.reset - started) / 1000)),
              );
              return json(res, 429, {
                error: "현재 AI 요청이 많습니다. 잠시 후 다시 시도해 주세요.",
              });
            }
            sharedWindow.count++;
          }
          const result = await askAI(input, user, config);
          return json(res, result.status, result.body);
        }
        if (url.pathname.startsWith("/api/"))
          return json(res, 404, { error: "요청한 API가 없습니다." });
        if (!["GET", "HEAD"].includes(req.method))
          return json(res, 405, { error: "지원하지 않는 요청입니다." });
        let pathname;
        try {
          pathname = decodeURIComponent(url.pathname);
        } catch {
          return json(res, 400, { error: "주소 형식을 확인해 주세요." });
        }
        const path = resolve(
          root,
          "." + (pathname === "/" ? "/index.html" : pathname),
        );
        if (!path.startsWith(root + sep))
          return json(res, 403, { error: "접근할 수 없습니다." });
        let file;
        try {
          if (!(await stat(path)).isFile()) throw new Error();
          file = await readFile(path);
        } catch {
          return json(res, 404, { error: "파일을 찾을 수 없습니다." });
        }
        res.writeHead(200, {
          "Content-Type": mime[extname(path)] || "application/octet-stream",
          "Cache-Control": pathname.startsWith("/assets/")
            ? "public,max-age=31536000,immutable"
            : "no-cache",
          "Content-Length": file.length,
        });
        res.end(req.method === "HEAD" ? undefined : file);
      } catch {
        json(res, 500, {
          error: "요청 처리에 실패했습니다. 잠시 후 다시 시도해 주세요.",
        });
      }
    },
  );
}
if (
  process.argv[1] &&
  import.meta.url === pathToFileURL(resolve(process.argv[1])).href
)
  createServer().listen(
    Number(process.env.PORT || 7730),
    process.env.HOST || "127.0.0.1",
    () => console.log("Excel Workspace server ready"),
  );
