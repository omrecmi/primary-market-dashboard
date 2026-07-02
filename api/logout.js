export default function handler(req, res) {
  res.writeHead(302, {
    Location: "/login",
    "Set-Cookie": [
      "dashboard_session=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0"
    ]
  });
  res.end();
}
