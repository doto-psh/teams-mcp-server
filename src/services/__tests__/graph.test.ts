import { promises as fs } from "node:fs";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mockUser, server } from "../../test-utils/setup.js";
import { GraphService } from "../graph.js";

// Mock the filesystem
vi.mock("node:fs", () => ({
  promises: {
    readFile: vi.fn(),
    writeFile: vi.fn(),
    unlink: vi.fn(),
  },
}));

// Mock @microsoft/microsoft-graph-client
vi.mock("@microsoft/microsoft-graph-client", () => ({
  Client: {
    initWithMiddleware: vi.fn(),
  },
}));

describe("GraphService", () => {
  let graphService: GraphService;

  beforeEach(() => {
    // Start MSW server
    server.listen({ onUnhandledRequest: "error" });

    // Reset GraphService singleton
    (GraphService as any).instance = undefined;
    graphService = GraphService.getInstance();

    // Clear all mocks
    vi.clearAllMocks();
  });

  afterEach(() => {
    server.resetHandlers();
    server.close();
  });

  describe("getInstance", () => {
    it("should return singleton instance", () => {
      const instance1 = GraphService.getInstance();
      const instance2 = GraphService.getInstance();

      expect(instance1).toBe(instance2);
    });
  });

  describe("getAuthStatus", () => {
    it("should return unauthenticated status when no token file exists", async () => {
      vi.mocked(fs.readFile).mockRejectedValue(new Error("File not found"));

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });

    it("should return unauthenticated status when token file is invalid", async () => {
      vi.mocked(fs.readFile).mockResolvedValue("invalid json");

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });

    it("should return unauthenticated status when token is expired", async () => {
      const expiredTokenData = JSON.stringify({
        clientId: "test-client-id",
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: new Date(Date.now() - 3600000).toISOString(), // 1 hour ago
        token: "expired-token",
      });

      vi.mocked(fs.readFile).mockResolvedValue(expiredTokenData);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });

    it("should return authenticated status with valid token", async () => {
      const validTokenData = JSON.stringify({
        clientId: "test-client-id",
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: new Date(Date.now() + 3600000).toISOString(), // 1 hour from now
        token: "valid-token",
      });

      vi.mocked(fs.readFile).mockResolvedValue(validTokenData);

      // Mock the Graph Client
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: true,
        userPrincipalName: mockUser.userPrincipalName,
        displayName: mockUser.displayName,
        expiresAt: expect.any(String),
      });
    });

    it("should handle Graph API errors gracefully", async () => {
      const validTokenData = JSON.stringify({
        clientId: "test-client-id",
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: new Date(Date.now() + 3600000).toISOString(),
        token: "valid-token",
      });

      vi.mocked(fs.readFile).mockResolvedValue(validTokenData);

      // Mock the Graph Client to throw an error
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockRejectedValue(new Error("API Error")),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: false,
      });
    });
  });

  describe("getClient", () => {
    it("should throw error when not authenticated", async () => {
      vi.mocked(fs.readFile).mockRejectedValue(new Error("File not found"));

      await expect(graphService.getClient()).rejects.toThrow(
        "Not authenticated. Please run the authentication CLI tool first"
      );
    });

    it("should return client when authenticated", async () => {
      const validTokenData = JSON.stringify({
        clientId: "test-client-id",
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: new Date(Date.now() + 3600000).toISOString(),
        token: "valid-token",
      });

      vi.mocked(fs.readFile).mockResolvedValue(validTokenData);

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const client = await graphService.getClient();

      expect(client).toBe(mockClient);
    });
  });

  describe("isAuthenticated", () => {
    it("should return false when not initialized", () => {
      expect(graphService.isAuthenticated()).toBe(false);
    });

    it("should return true when client is initialized", async () => {
      const validTokenData = JSON.stringify({
        clientId: "test-client-id",
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: new Date(Date.now() + 3600000).toISOString(),
        token: "valid-token",
      });

      vi.mocked(fs.readFile).mockResolvedValue(validTokenData);

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      // Initialize the client
      await graphService.getAuthStatus();

      expect(graphService.isAuthenticated()).toBe(true);
    });
  });

  describe("token refresh scenarios", () => {
    it("should handle token refresh when expires soon", async () => {
      const soonExpiringTokenData = JSON.stringify({
        clientId: "test-client-id",
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: new Date(Date.now() + 300000).toISOString(), // 5 minutes from now
        token: "soon-expiring-token",
      });

      vi.mocked(fs.readFile).mockResolvedValue(soonExpiringTokenData);

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(true);
      expect(status.expiresAt).toBeDefined();
    });
  });

  describe("concurrent initialization", () => {
    it("should handle concurrent calls to initializeClient", async () => {
      const validTokenData = JSON.stringify({
        clientId: "test-client-id",
        authenticated: true,
        timestamp: new Date().toISOString(),
        expiresAt: new Date(Date.now() + 3600000).toISOString(),
        token: "valid-token",
      });

      vi.mocked(fs.readFile).mockResolvedValue(validTokenData);

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(mockUser),
        }),
      };

      const { Client } = await import("@microsoft/microsoft-graph-client");
      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      // Make multiple concurrent calls
      const promises = [
        graphService.getAuthStatus(),
        graphService.getAuthStatus(),
        graphService.getAuthStatus(),
      ];

      const results = await Promise.all(promises);

      // All should return the same authenticated status
      for (const result of results) {
        expect(result.isAuthenticated).toBe(true);
      }

      // readFile should be called for each concurrent call since we reset the singleton
      expect(vi.mocked(fs.readFile)).toHaveBeenCalled();
    });
  });
});
