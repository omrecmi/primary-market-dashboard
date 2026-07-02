const PUBLIC_PATHS = new Set([
  "/login",
  "/login.html",
  "/api/login",
  "/api/logout",
  "/favicon.ico"
]);

function isPublicPath(pathname) {
  if (PUBLIC_PATHS.has(pathname)) {
    return true;
  }

  return pathname.startsWith("/_vercel") || pathname.startsWith("/.well-known");
}

async function signValue(value, secret) {
  const encoder = new TextEncoder();
  const key = await crypto.subtle.importKey(
    "raw",
    encoder.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const signature = await crypto.subtle.sign("HMAC", key, encoder.encode(value));
  const bytes = Array.from(new Uint8Array(signature));
  return bytes.map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

async function isValidSession(token, secret) {
  if (!token) {
    return false;
  }

  const lastDot = token.lastIndexOf(".");
  if (lastDot <= 0) {
    return false;
  }

  const payload = token.slice(0, lastDot);
  const providedSignature = token.slice(lastDot + 1);
  const expiryText = payload.split(".")[0];
  const expiry = Number.parseInt(expiryText, 10);

  if (!Number.isFinite(expiry) || Date.now() > expiry) {
    return false;
  }

  const expectedSignature = await signValue(payload, secret);
  return expectedSignature === providedSignature;
}

export default async function middleware(request) {
  const url = new URL(request.url);
  const pathname = url.pathname;
  const sessionSecret = process.env.SESSION_SECRET;

  if (!sessionSecret) {
    return new Response("Missing SESSION_SECRET environment variable.", { status: 500 });
  }

  const token = request.cookies.get("dashboard_session")?.value;
  const authenticated = await isValidSession(token, sessionSecret);

  if (pathname === "/login" || pathname === "/login.html") {
    if (authenticated) {
      return Response.redirect(new URL("/", request.url));
    }
    return;
  }

  if (isPublicPath(pathname)) {
    return;
  }

  if (!authenticated) {
    const loginUrl = new URL("/login", request.url);
    loginUrl.searchParams.set("next", pathname === "/" ? "/" : `${pathname}${url.search}`);
    return Response.redirect(loginUrl);
  }
}

export const config = {
  matcher: ["/:path*"]
};
