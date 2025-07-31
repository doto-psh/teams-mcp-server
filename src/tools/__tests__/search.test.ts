import type { Client } from "@microsoft/microsoft-graph-client";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import { registerSearchTools } from "../search.js";

// Mock the Graph service
const mockGraphService = {
  getClient: vi.fn(),
} as unknown as GraphService;

// Mock the MCP server
const mockServer = {
  tool: vi.fn(),
} as unknown as McpServer;

// Mock client responses
const mockClient = {
  api: vi.fn(),
} as unknown as Client;

describe("Search Tools", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockGraphService.getClient = vi.fn().mockResolvedValue(mockClient);
  });

  describe("registerSearchTools", () => {
    it("should register all search tools", () => {
      registerSearchTools(mockServer, mockGraphService);

      expect(mockServer.tool).toHaveBeenCalledTimes(3);
      expect(mockServer.tool).toHaveBeenCalledWith(
        "search_messages",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_recent_messages",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_my_mentions",
        expect.any(String),
        expect.any(Object),
        expect.any(Function)
      );
    });
  });

  describe("search_messages", () => {
    let searchMessagesHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerSearchTools(mockServer, mockGraphService);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "search_messages");
      searchMessagesHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should search messages with default parameters", async () => {
      const mockSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    rank: 1,
                    summary: "Found message",
                    resource: {
                      id: "msg1",
                      body: { content: "Hello world" },
                      from: { user: { displayName: "John Doe" } },
                      createdDateTime: "2023-01-01T10:00:00Z",
                      chatId: "chat123",
                    },
                  },
                ],
                total: 1,
                moreResultsAvailable: false,
              },
            ],
          },
        ],
      };

      const mockApiChain = {
        post: vi.fn().mockResolvedValue(mockSearchResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await searchMessagesHandler({
        query: "hello",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/search/query");
      expect(mockApiChain.post).toHaveBeenCalledWith({
        requests: [
          {
            entityTypes: ["chatMessage"],
            query: {
              queryString: "hello",
            },
            from: 0,
            size: undefined,
            enableTopResults: undefined,
          },
        ],
      });

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.query).toBe("hello");
      expect(parsedResponse.scope).toBe(undefined);
      expect(parsedResponse.results).toHaveLength(1);
      expect(parsedResponse.results[0].content).toBe("Hello world");
    });

    it("should apply channel scope filter", async () => {
      const mockSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                hits: [],
                total: 0,
                moreResultsAvailable: false,
              },
            ],
          },
        ],
      };

      const mockApiChain = {
        post: vi.fn().mockResolvedValue(mockSearchResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      await searchMessagesHandler({
        query: "test",
        scope: "channels",
        limit: 10,
        enableTopResults: false,
      });

      expect(mockApiChain.post).toHaveBeenCalledWith({
        requests: [
          {
            entityTypes: ["chatMessage"],
            query: {
              queryString: "test AND (channelIdentity/channelId:*)",
            },
            from: 0,
            size: 10,
            enableTopResults: false,
          },
        ],
      });
    });

    it("should handle no search results", async () => {
      const mockSearchResponse = {
        value: [],
      };

      const mockApiChain = {
        post: vi.fn().mockResolvedValue(mockSearchResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await searchMessagesHandler({
        query: "nonexistent",
      });

      expect(result.content[0].text).toBe("No messages found matching your search criteria.");
    });

    it("should handle search errors", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Search API error")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await searchMessagesHandler({
        query: "error",
      });

      expect(result.content[0].text).toBe("❌ Error searching messages: Search API error");
    });
  });

  describe("get_recent_messages", () => {
    let getRecentMessagesHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerSearchTools(mockServer, mockGraphService);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "get_recent_messages");
      getRecentMessagesHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should get recent messages with advanced search", async () => {
      const mockSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    rank: 1,
                    summary: "Recent message",
                    resource: {
                      id: "msg1",
                      body: { content: "Recent content" },
                      from: { user: { displayName: "User" } },
                      createdDateTime: "2023-01-01T10:00:00Z",
                    },
                  },
                ],
                total: 1,
                moreResultsAvailable: false,
              },
            ],
          },
        ],
      };

      const mockApiChain = {
        post: vi.fn().mockResolvedValue(mockSearchResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getRecentMessagesHandler({
        hours: 24,
        keywords: "test",
        mentionsUser: "user123",
        hasAttachments: true,
        importance: "high",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/search/query");

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.method).toBe("search_api");
      expect(parsedResponse.messages).toHaveLength(1);
    });

    it("should fall back to basic search when advanced search fails", async () => {
      const mockChats = [{ id: "chat1" }];
      const mockMessages = [
        {
          id: "msg1",
          body: { content: "Basic message" },
          from: { user: { displayName: "User" } },
          createdDateTime: new Date().toISOString(),
        },
      ];

      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Search failed")),
        get: vi
          .fn()
          .mockResolvedValueOnce({ value: mockChats })
          .mockResolvedValueOnce({ value: mockMessages }),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getRecentMessagesHandler({
        hours: 12,
        includeChats: true,
        includeChannels: false,
      });

      expect(mockClient.api).toHaveBeenCalledWith("/me/chats?$expand=members");

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.method).toBe("direct_chat_queries");
      expect(parsedResponse.messages).toHaveLength(1);
    });

    it("should handle errors in basic search", async () => {
      const mockApiChain = {
        post: vi.fn().mockRejectedValue(new Error("Search failed")),
        get: vi.fn().mockRejectedValue(new Error("Basic search failed")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getRecentMessagesHandler({
        hours: 24,
      });

      expect(result.content[0].text).toBe("❌ Error getting recent messages: Basic search failed");
    });
  });

  describe("get_my_mentions", () => {
    let getMyMentionsHandler: (args?: any) => Promise<any>;

    beforeEach(() => {
      registerSearchTools(mockServer, mockGraphService);
      const call = vi
        .mocked(mockServer.tool)
        .mock.calls.find(([name]) => name === "get_my_mentions");
      getMyMentionsHandler = call?.[3] as unknown as (args?: any) => Promise<any>;
    });

    it("should get mentions using search API", async () => {
      const mockUser = { id: "currentuser123" };
      const mockSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    rank: 1,
                    summary: "You were mentioned",
                    resource: {
                      id: "msg1",
                      body: { content: "@currentuser123 hello" },
                      from: { user: { displayName: "Colleague" } },
                      createdDateTime: "2023-01-01T10:00:00Z",
                      chatId: "chat123",
                    },
                  },
                ],
                total: 1,
                moreResultsAvailable: false,
              },
            ],
          },
        ],
      };

      const mockApiChain = {
        get: vi.fn().mockResolvedValue(mockUser),
        post: vi.fn().mockResolvedValue(mockSearchResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getMyMentionsHandler({
        hours: 24,
        scope: "all",
      });

      expect(mockClient.api).toHaveBeenCalledWith("/me");
      expect(mockClient.api).toHaveBeenCalledWith("/search/query");

      const parsedResponse = JSON.parse(result.content[0].text);
      expect(parsedResponse.mentions).toHaveLength(1);
    });

    it("should handle no mentions found", async () => {
      const mockUser = { id: "currentuser123" };
      const mockSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                hits: [],
                total: 0,
                moreResultsAvailable: false,
              },
            ],
          },
        ],
      };

      const mockApiChain = {
        get: vi.fn().mockResolvedValue(mockUser),
        post: vi.fn().mockResolvedValue(mockSearchResponse),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getMyMentionsHandler({ hours: 24 });

      expect(result.content[0].text).toBe("No recent mentions found.");
    });

    it("should handle errors", async () => {
      const mockApiChain = {
        get: vi.fn().mockRejectedValue(new Error("User lookup failed")),
      };
      mockClient.api = vi.fn().mockReturnValue(mockApiChain);

      const result = await getMyMentionsHandler({ hours: 24 });

      expect(result.content[0].text).toBe("❌ Error getting mentions: User lookup failed");
    });
  });
});
