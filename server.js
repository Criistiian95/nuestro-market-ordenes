const http = require("http");
const zlib = require("zlib");
const fs = require("fs");
const path = require("path");

const PORT = process.env.PORT || 3000;
const indexPath = path.join(__dirname, "index.html");

function send(res, status, type, body) {
  res.writeHead(status, {
    "Content-Type": type,
    "Cache-Control": type.includes("text/html") ? "no-cache" : "public, max-age=60"
  });
  res.end(body);
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, "http://localhost");

  if (url.pathname === "/health") {
    return send(res, 200, "text/plain; charset=utf-8", "ok");
  }

  if (url.pathname === "/data.json") {
    try {
      const b64 = process.env.DATA_BR_B64 || "";
      if (!b64) return send(res, 500, "application/json; charset=utf-8", JSON.stringify({error:"DATA_BR_B64 no configurada"}));
      const buf = Buffer.from(b64, "base64");
      const json = zlib.brotliDecompressSync(buf);
      return send(res, 200, "application/json; charset=utf-8", json);
    } catch (e) {
      return send(res, 500, "application/json; charset=utf-8", JSON.stringify({error:String(e.message || e)}));
    }
  }

  if (url.pathname === "/" || url.pathname === "/index.html") {
    try {
      return send(res, 200, "text/html; charset=utf-8", fs.readFileSync(indexPath));
    } catch (e) {
      return send(res, 500, "text/plain; charset=utf-8", "No se pudo cargar la aplicación.");
    }
  }

  send(res, 404, "text/plain; charset=utf-8", "No encontrado");
});

server.listen(PORT, "0.0.0.0", () => console.log("Nuestro Market escuchando en", PORT));
