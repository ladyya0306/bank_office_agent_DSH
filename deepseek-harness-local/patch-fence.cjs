const fs = require("fs");
const path = require("path");

const target = path.join(
  __dirname,
  "node_modules",
  "@deepseek-ai",
  "dsh-client-connection",
  "lib",
  "index.js",
);

const MARKER = "dsh-local-fence-patch";

if (!fs.existsSync(target)) {
  console.error("[patch-fence] target not found:", target);
  process.exit(1);
}

let src = fs.readFileSync(target, "utf8");

if (src.includes(MARKER)) {
  console.log("[patch-fence] already patched, nothing to do");
  process.exit(0);
}

const anchor = "function isTrustedApiRequest(request, trustedHosts) {";
const start = src.indexOf(anchor);
if (start === -1) {
  console.error("[patch-fence] could not find isTrustedApiRequest; aborting");
  process.exit(1);
}

let depth = 0;
let end = -1;
for (let i = src.indexOf("{", start); i < src.length; i += 1) {
  const c = src[i];
  if (c === "{") depth += 1;
  else if (c === "}") {
    depth -= 1;
    if (depth === 0) {
      end = i;
      break;
    }
  }
}
if (end === -1) {
  console.error("[patch-fence] could not find function end; aborting");
  process.exit(1);
}

const replacement = [
  "function isTrustedApiRequest(request, trustedHosts) {",
  '\tconst host = header$1(request.headers, "host");',
  "\tif (host === void 0) return false;",
  "\tconst hostUrl = parseAuthority(host);",
  "\tif (hostUrl === void 0) return false;",
  "\tif (!isLoopbackHostname(hostUrl.hostname) && !isTrustedAuthority(hostUrl, trustedHosts)) return false;",
  "\t// " + MARKER + ": Origin / sec-fetch-site checks relaxed for local use.",
  "\t// A loopback or trusted Host is enough; the authority-bound SameSite=Strict",
  "\t// browser-session cookie still gates every request (missing/invalid => 401).",
  "\treturn true;",
  "}",
].join("\n");

src = src.slice(0, start) + replacement + src.slice(end + 1);
fs.writeFileSync(target, src);
console.log("[patch-fence] patched", target);
