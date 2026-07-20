import crypto from "node:crypto";

const SESSION_MAX_AGE_SECONDS = 60 * 60 * 12;
const FALLBACK_PASSWORD_SHA256 =
  "cd1152cf7d7e17e5aa3eea8f5b36e1cc4bd250b1b7f8dfc32b39f9b104b9f981";

function signValue(value, secret) {
  return crypto.createHmac("sha256", secret).update(value).digest("hex");
}

function buildSessionToken(secret) {
  const expiresAt = Date.now() + SESSION_MAX_AGE_SECONDS * 1000;
  const nonce = crypto.randomBytes(16).toString("hex");
  const payload = `${expiresAt}.${nonce}`;
  const signature = signValue(payload, secret);
  return `${payload}.${signature}`;
}

function hashPassword(password) {
  return crypto.createHash("sha256").update(password).digest("hex");
}

function parseRequestBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];

    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => {
      try {
        const raw = Buffer.concat(chunks).toString("utf8");
        resolve(raw ? JSON.parse(raw) : {});
      } catch (error) {
        reject(error);
      }
    });
    req.on("error", reject);
  });
}

function sendJson(res, statusCode, payload, extraHeaders = {}) {
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    ...extraHeaders
  });
  res.end(JSON.stringify(payload));
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    sendJson(res, 405, { error: "Method not allowed" }, { Allow: "POST" });
    return;
  }

  const sitePassword = process.env.SITE_PASSWORD;
  const sessionSecret = process.env.SESSION_SECRET;

  if (!sessionSecret) {
    sendJson(res, 500, { error: "Missing auth environment variables" });
    return;
  }

  let body;
  try {
    body = await parseRequestBody(req);
  } catch {
    sendJson(res, 400, { error: "Invalid request body" });
    return;
  }

  const submittedPassword = typeof body.password === "string" ? body.password : "";
  const matchesEnvPassword = sitePassword && submittedPassword === sitePassword;
  const matchesFallbackPassword = hashPassword(submittedPassword) === FALLBACK_PASSWORD_SHA256;

  if (!matchesEnvPassword && !matchesFallbackPassword) {
    sendJson(res, 401, { error: "Invalid password" });
    return;
  }

  const token = buildSessionToken(sessionSecret);
  const nextPath = typeof body.next === "string" && body.next.startsWith("/") ? body.next : "/";

  sendJson(
    res,
    200,
    { ok: true, next: nextPath },
    {
      "Set-Cookie": [
        `dashboard_session=${token}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${SESSION_MAX_AGE_SECONDS}`
      ]
    }
  );
}
