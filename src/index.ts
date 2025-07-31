#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { GraphService } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";
async function startMcpServer() {
  // Create MCP server
  const server = new McpServer({
    name: "teams-mcp",
    version: "0.3.3",
  });

  // Initialize Graph service (singleton)
  const graphService = GraphService.getInstance();

  // Register all tools
  registerAuthTools(server, graphService);
  registerUsersTools(server, graphService);
  registerTeamsTools(server, graphService);
  registerChatTools(server, graphService);
  registerSearchTools(server, graphService);

  // Start server
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Microsoft Graph MCP Server started");
}

// Main function to handle both CLI and MCP server modes
async function main() {
  const args = process.argv.slice(2);
  const command = args[0];

  // CLI commands
  switch (command) {
    case undefined:
      // No command = start MCP server
      await startMcpServer();
      return;
    default:
      console.error(`Unknown command: ${command}`);
      console.error("Use --help to see available commands");
      process.exit(1);
  }
}

// Handle uncaught errors
process.on("uncaughtException", (error) => {
  console.error("Uncaught exception:", error);
  process.exit(1);
});

process.on("unhandledRejection", (reason, promise) => {
  console.error("Unhandled rejection at:", promise, "reason:", reason);
  process.exit(1);
});

main().catch((error) => {
  console.error("Failed to start:", error);
  process.exit(1);
});
