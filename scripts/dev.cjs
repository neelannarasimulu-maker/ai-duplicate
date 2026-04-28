const { spawn } = require("node:child_process");

const commands = [
  ["node", ["scripts/api-server.cjs"]],
  ["npx", ["vite", "--host", "0.0.0.0"]],
];

const children = commands.map(([command, args]) => {
  const child = spawn(command, args, {
    stdio: "inherit",
    shell: process.platform === "win32",
  });

  child.on("exit", (code) => {
    if (code) process.exitCode = code;
  });

  return child;
});

function shutdown() {
  for (const child of children) child.kill();
  process.exit();
}

process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);
