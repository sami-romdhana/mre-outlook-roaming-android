const fs = require("fs");
const https = require("https");
const path = require("path");

const programOptions = process.argv.slice(2);

try {
  const manifestRelPath = process.argv[2];
  const manifestPath = path.join(__dirname, manifestRelPath);
  const manifest = fs.readFileSync(manifestPath, "utf-8");

  const options = {
    hostname: "validationgateway.omex.office.net",
    port: 443,
    path: "/package/api/check?gates=DisableIconDimensionValidation",
    method: "POST",
    headers: {
      "Content-Type": "application/xml",
      "Content-Length": Buffer.byteLength(manifest),
    },
  };

  const req = https.request(options, (res) => {
    let data = "";

    res.on("data", (d) => {
      data += d;
    });

    res.on("end", () => {
      logResults(data);
    });
  });

  req.on("error", (error) => {
    console.error(error);
  });

  req.write(manifest);
  req.end();

  function logResults(data) {
    const results = JSON.parse(data);

    console.log(`"${manifestRelPath}" status: ${results.status}`);

    if (results.errors?.length) {
      console.error("Errors: ");
      results.errors.forEach(log(console.error));
      console.log();
    }

    if (results.warnings?.length) {
      console.error("Warnings: ");
      results.warnings.forEach(log(console.error));
      console.log();
    }

    if (results.notes?.length && programOptions.includes("--notes")) {
      console.error("Notes: ");
      results.notes.forEach(log(console.error));
      console.log();
    }
  }

  function log(level) {
    return (item) => level(`· ${item.title} : ${item.content}`);
  }
} catch (err) {
  console.error(err);
}
