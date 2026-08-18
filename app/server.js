const https = require("https");
const path = require("path");
const express = require("express");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;

async function main() {
  const app = express();
  app.use(express.static(path.join(__dirname, "..")));

  const httpsOptions = await devCerts.getHttpsServerOptions();

  https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`Serveur de dev disponible sur https://localhost:${PORT}`);
  });
}

main().catch((err) => {
  console.error("Erreur au demarrage du serveur:", err);
  process.exit(1);
});
