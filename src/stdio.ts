#!/usr/bin/env node
// stdio entry point for local MCP clients.
import { startStdio } from './index.js';

startStdio().catch((err) => {
  process.stderr.write(`Fatal error: ${err}\n`);
  process.exit(1);
});
