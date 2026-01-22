import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphService } from "../services/graph.js";
import type {
  Channel,
  ChannelSummary,
  ChatMessage,
  ConversationMember,
  GraphApiResponse,
  MemberSummary,
  MessageSummary,
  Team,
  TeamSummary,
} from "../types/graph.js";
import type { GraphMention, MentionMapping } from "../types/mentions.js";
import {
  type ImageAttachment,
  imageUrlToBase64,
  isValidImageType,
  uploadImageAsHostedContent,
} from "../utils/attachments.js";
import { markdownToHtml } from "../utils/markdown.js";
import { processMentionsInHtml, searchUsers, type UserInfo } from "../utils/users.js";

/**
 * Validates channel ID format and logs a warning if it doesn't match expected Teams format.
 * Teams channel IDs typically follow the pattern: 19:xxx@thread.tacv2
 */
function validateChannelIdFormat(channelId: string, mentionText: string): void {
  const channelIdPattern = /^19:[a-zA-Z0-9_-]+@thread\.tacv2$/;
  if (!channelIdPattern.test(channelId)) {
    console.warn(
      `Channel mention "${mentionText}": Channel ID "${channelId}" may not be in the expected format. ` +
        `Expected format: "19:xxx@thread.tacv2". The mention may not work correctly.`
    );
  }
}

/**
 * Shared Zod schema for mentions that supports both user and channel mentions.
 * Each mention must have either userId (for user mentions) OR channelId (for channel mentions), not both.
 */
const mentionsSchema = z
  .array(
    z.object({
      mention: z
        .string()
        .describe("The @mention text as it appears in the message (e.g., 'John Doe' or 'General')"),
      userId: z
        .string()
        .optional()
        .describe("Azure AD User ID for user mentions - mutually exclusive with channelId"),
      channelId: z
        .string()
        .optional()
        .describe("Channel ID for channel mentions (e.g., '19:abc123@thread.tacv2') - mutually exclusive with userId"),
    })
  )
  .refine(
    (mentions) =>
      mentions.every((m) => {
        const hasUserId = m.userId !== undefined && m.userId !== "";
        const hasChannelId = m.channelId !== undefined && m.channelId !== "";
        return (hasUserId && !hasChannelId) || (!hasUserId && hasChannelId);
      }),
    {
      message:
        "Invalid mention configuration: Each mention must specify exactly one of 'userId' (for user mentions) or 'channelId' (for channel mentions). " +
        "Found mention with both or neither. Check that each mention in the array has either userId OR channelId, not both and not neither.",
    }
  )
  .optional()
  .describe("Array of mentions - each must specify either userId (for user mentions) or channelId (for channel mentions)");

export function registerTeamsTools(server: McpServer, graphService: GraphService) {
  // List user's teams
  server.tool(
    "list_teams",
    "List all Microsoft Teams that the current user is a member of. Returns team names, descriptions, and IDs.",
    {},
    async () => {
      try {
        const client = await graphService.getClient();
        const response = (await client.api("/me/joinedTeams").get()) as GraphApiResponse<Team>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No teams found.",
              },
            ],
          };
        }

        const teamList: TeamSummary[] = response.value.map((team: Team) => ({
          id: team.id,
          displayName: team.displayName,
          description: team.description,
          isArchived: team.isArchived,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(teamList, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // List channels in a team
  server.tool(
    "list_channels",
    "List all channels in a specific Microsoft Team. Returns channel names, descriptions, types, and IDs for the specified team.",
    {
      teamId: z.string().describe("Team ID"),
    },
    async ({ teamId }) => {
      try {
        const client = await graphService.getClient();
        const response = (await client
          .api(`/teams/${teamId}/channels`)
          .get()) as GraphApiResponse<Channel>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No channels found in this team.",
              },
            ],
          };
        }

        const channelList: ChannelSummary[] = response.value.map((channel: Channel) => ({
          id: channel.id,
          displayName: channel.displayName,
          description: channel.description,
          membershipType: channel.membershipType,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(channelList, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Get channel messages
  server.tool(
    "get_channel_messages",
    "Retrieve recent messages from a specific channel in a Microsoft Team. Returns message content, sender information, and timestamps.",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(20)
        .describe("Number of messages to retrieve (default: 20)"),
    },
    async ({ teamId, channelId, limit }) => {
      try {
        const client = await graphService.getClient();

        // Build query parameters - Teams channel messages API has limited query support
        // Only $top is supported, no $orderby, $filter, etc.
        const queryParams: string[] = [`$top=${limit}`];
        const queryString = queryParams.join("&");

        const response = (await client
          .api(`/teams/${teamId}/channels/${channelId}/messages?${queryString}`)
          .get()) as GraphApiResponse<ChatMessage>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No messages found in this channel.",
              },
            ],
          };
        }

        const messageList: MessageSummary[] = response.value.map((message: ChatMessage) => ({
          id: message.id,
          content: message.body?.content,
          from: message.from?.user?.displayName,
          createdDateTime: message.createdDateTime,
          importance: message.importance,
        }));

        // Sort messages by creation date (newest first) since API doesn't support orderby
        messageList.sort((a, b) => {
          const dateA = new Date(a.createdDateTime || 0).getTime();
          const dateB = new Date(b.createdDateTime || 0).getTime();
          return dateB - dateA;
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  totalReturned: messageList.length,
                  hasMore: !!response["@odata.nextLink"],
                  messages: messageList,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Send message to channel
  server.tool(
    "send_channel_message",
    "Send a message to a specific channel in a Microsoft Team. Supports text and markdown formatting, mentions, and importance levels.",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      message: z.string().describe("Message content"),
      importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
      format: z.enum(["text", "markdown"]).optional().describe("Message format (text or markdown)"),
      mentions: mentionsSchema,
      imageUrl: z.string().optional().describe("URL of an image to attach to the message"),
      imageData: z.string().optional().describe("Base64 encoded image data to attach"),
      imageContentType: z
        .string()
        .optional()
        .describe("MIME type of the image (e.g., 'image/jpeg', 'image/png')"),
      imageFileName: z.string().optional().describe("Name for the attached image file"),
    },
    async ({
      teamId,
      channelId,
      message,
      importance = "normal",
      format = "text",
      mentions,
      imageUrl,
      imageData,
      imageContentType,
      imageFileName,
    }) => {
      try {
        const client = await graphService.getClient();

        // Process message content based on format
        let content: string;
        let contentType: "text" | "html";

        if (format === "markdown") {
          content = await markdownToHtml(message);
          contentType = "html";
        } else {
          content = message;
          contentType = "text";
        }

        // Process @mentions if provided (supports both user and channel mentions)
        const mentionMappings: MentionMapping[] = [];
        if (mentions && mentions.length > 0) {
          // Convert provided mentions to mappings with display names
          for (const mention of mentions) {
            if (mention.userId) {
              // User mention - look up display name
              try {
                const userResponse = await client
                  .api(`/users/${mention.userId}`)
                  .select("displayName")
                  .get();
                mentionMappings.push({
                  mention: mention.mention,
                  userId: mention.userId,
                  displayName: userResponse.displayName || mention.mention,
                });
              } catch (_error) {
                console.warn(
                  `Could not resolve user ${mention.userId}, using mention text as display name`
                );
                mentionMappings.push({
                  mention: mention.mention,
                  userId: mention.userId,
                  displayName: mention.mention,
                });
              }
            } else if (mention.channelId) {
              // Channel mention - validate format and use mention text as display name
              validateChannelIdFormat(mention.channelId, mention.mention);
              mentionMappings.push({
                mention: mention.mention,
                channelId: mention.channelId,
                displayName: mention.mention,
              });
            }
          }
        }

        // Process mentions in HTML content
        let finalMentions: GraphMention[] = [];
        if (mentionMappings.length > 0) {
          const result = processMentionsInHtml(content, mentionMappings);
          content = result.content;
          finalMentions = result.mentions;

          // Ensure we're using HTML content type when mentions are present
          contentType = "html";
        }

        // Handle image attachment
        const attachments: ImageAttachment[] = [];
        if (imageUrl || imageData) {
          let imageInfo: { data: string; contentType: string } | null = null;

          if (imageUrl) {
            imageInfo = await imageUrlToBase64(imageUrl);
            if (!imageInfo) {
              return {
                content: [
                  {
                    type: "text" as const,
                    text: `❌ Failed to download image from URL: ${imageUrl}`,
                  },
                ],
                isError: true,
              };
            }
          } else if (imageData && imageContentType) {
            if (!isValidImageType(imageContentType)) {
              return {
                content: [
                  {
                    type: "text" as const,
                    text: `❌ Unsupported image type: ${imageContentType}`,
                  },
                ],
                isError: true,
              };
            }
            imageInfo = { data: imageData, contentType: imageContentType };
          }

          if (imageInfo) {
            const uploadResult = await uploadImageAsHostedContent(
              graphService,
              teamId,
              channelId,
              imageInfo.data,
              imageInfo.contentType,
              imageFileName
            );

            if (uploadResult) {
              attachments.push(uploadResult.attachment);
            } else {
              return {
                content: [
                  {
                    type: "text" as const,
                    text: "❌ Failed to upload image attachment",
                  },
                ],
                isError: true,
              };
            }
          }
        }

        // Build message payload
        const messagePayload: any = {
          body: {
            content,
            contentType,
          },
          importance,
        };

        if (finalMentions.length > 0) {
          messagePayload.mentions = finalMentions;
        }

        if (attachments.length > 0) {
          messagePayload.attachments = attachments;
        }

        const result = (await client
          .api(`/teams/${teamId}/channels/${channelId}/messages`)
          .post(messagePayload)) as ChatMessage;

        // Build success message
        const successText = `✅ Message sent successfully. Message ID: ${result.id}${
          finalMentions.length > 0
            ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(", ")}`
            : ""
        }${attachments.length > 0 ? `\n🖼️ Image attached: ${attachments[0].name}` : ""}`;

        return {
          content: [
            {
              type: "text" as const,
              text: successText,
            },
          ],
        };
      } catch (error: any) {
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to send message: ${error.message}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Get replies to a message in a channel
  server.tool(
    "get_channel_message_replies",
    "Get all replies to a specific message in a channel. Returns reply content, sender information, and timestamps.",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      messageId: z.string().describe("Message ID to get replies for"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(20)
        .describe("Number of replies to retrieve (default: 20)"),
    },
    async ({ teamId, channelId, messageId, limit }) => {
      try {
        const client = await graphService.getClient();

        // Only $top is supported for message replies
        const queryParams: string[] = [`$top=${limit}`];
        const queryString = queryParams.join("&");

        const response = (await client
          .api(
            `/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies?${queryString}`
          )
          .get()) as GraphApiResponse<ChatMessage>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No replies found for this message.",
              },
            ],
          };
        }

        const repliesList: MessageSummary[] = response.value.map((reply: ChatMessage) => ({
          id: reply.id,
          content: reply.body?.content,
          from: reply.from?.user?.displayName,
          createdDateTime: reply.createdDateTime,
          importance: reply.importance,
        }));

        // Sort replies by creation date (oldest first for replies)
        repliesList.sort((a, b) => {
          const dateA = new Date(a.createdDateTime || 0).getTime();
          const dateB = new Date(b.createdDateTime || 0).getTime();
          return dateA - dateB;
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  parentMessageId: messageId,
                  totalReplies: repliesList.length,
                  hasMore: !!response["@odata.nextLink"],
                  replies: repliesList,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Reply to a message in a channel
  server.tool(
    "reply_to_channel_message",
    "Reply to a specific message in a channel. Supports text and markdown formatting, mentions, and importance levels.",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      messageId: z.string().describe("Message ID to reply to"),
      message: z.string().describe("Reply content"),
      importance: z.enum(["normal", "high", "urgent"]).optional().describe("Message importance"),
      format: z.enum(["text", "markdown"]).optional().describe("Message format (text or markdown)"),
      mentions: mentionsSchema,
      imageUrl: z.string().optional().describe("URL of an image to attach to the reply"),
      imageData: z.string().optional().describe("Base64 encoded image data to attach"),
      imageContentType: z
        .string()
        .optional()
        .describe("MIME type of the image (e.g., 'image/jpeg', 'image/png')"),
      imageFileName: z.string().optional().describe("Name for the attached image file"),
    },
    async ({
      teamId,
      channelId,
      messageId,
      message,
      importance = "normal",
      format = "text",
      mentions,
      imageUrl,
      imageData,
      imageContentType,
      imageFileName,
    }) => {
      try {
        const client = await graphService.getClient();

        // Process message content based on format
        let content: string;
        let contentType: "text" | "html";

        if (format === "markdown") {
          content = await markdownToHtml(message);
          contentType = "html";
        } else {
          content = message;
          contentType = "text";
        }

        // Process @mentions if provided (supports both user and channel mentions)
        const mentionMappings: MentionMapping[] = [];
        if (mentions && mentions.length > 0) {
          // Convert provided mentions to mappings with display names
          for (const mention of mentions) {
            if (mention.userId) {
              // User mention - look up display name
              try {
                const userResponse = await client
                  .api(`/users/${mention.userId}`)
                  .select("displayName")
                  .get();
                mentionMappings.push({
                  mention: mention.mention,
                  userId: mention.userId,
                  displayName: userResponse.displayName || mention.mention,
                });
              } catch (_error) {
                console.warn(
                  `Could not resolve user ${mention.userId}, using mention text as display name`
                );
                mentionMappings.push({
                  mention: mention.mention,
                  userId: mention.userId,
                  displayName: mention.mention,
                });
              }
            } else if (mention.channelId) {
              // Channel mention - validate format and use mention text as display name
              validateChannelIdFormat(mention.channelId, mention.mention);
              mentionMappings.push({
                mention: mention.mention,
                channelId: mention.channelId,
                displayName: mention.mention,
              });
            }
          }
        }

        // Process mentions in HTML content
        let finalMentions: GraphMention[] = [];
        if (mentionMappings.length > 0) {
          const result = processMentionsInHtml(content, mentionMappings);
          content = result.content;
          finalMentions = result.mentions;

          // Ensure we're using HTML content type when mentions are present
          contentType = "html";
        }

        // Handle image attachment
        const attachments: ImageAttachment[] = [];
        if (imageUrl || imageData) {
          let imageInfo: { data: string; contentType: string } | null = null;

          if (imageUrl) {
            imageInfo = await imageUrlToBase64(imageUrl);
            if (!imageInfo) {
              return {
                content: [
                  {
                    type: "text" as const,
                    text: `❌ Failed to download image from URL: ${imageUrl}`,
                  },
                ],
                isError: true,
              };
            }
          } else if (imageData && imageContentType) {
            if (!isValidImageType(imageContentType)) {
              return {
                content: [
                  {
                    type: "text" as const,
                    text: `❌ Unsupported image type: ${imageContentType}`,
                  },
                ],
                isError: true,
              };
            }
            imageInfo = { data: imageData, contentType: imageContentType };
          }

          if (imageInfo) {
            const uploadResult = await uploadImageAsHostedContent(
              graphService,
              teamId,
              channelId,
              imageInfo.data,
              imageInfo.contentType,
              imageFileName
            );

            if (uploadResult) {
              attachments.push(uploadResult.attachment);
            } else {
              return {
                content: [
                  {
                    type: "text" as const,
                    text: "❌ Failed to upload image attachment",
                  },
                ],
                isError: true,
              };
            }
          }
        }

        // Build message payload
        const messagePayload: any = {
          body: {
            content,
            contentType,
          },
          importance,
        };

        if (finalMentions.length > 0) {
          messagePayload.mentions = finalMentions;
        }

        if (attachments.length > 0) {
          messagePayload.attachments = attachments;
        }

        const result = (await client
          .api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`)
          .post(messagePayload)) as ChatMessage;

        // Build success message
        const successText = `✅ Reply sent successfully. Reply ID: ${result.id}${
          finalMentions.length > 0
            ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(", ")}`
            : ""
        }${attachments.length > 0 ? `\n🖼️ Image attached: ${attachments[0].name}` : ""}`;

        return {
          content: [
            {
              type: "text" as const,
              text: successText,
            },
          ],
        };
      } catch (error: any) {
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Failed to send reply: ${error.message}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // List team members
  server.tool(
    "list_team_members",
    "List all members of a specific Microsoft Team. Returns member names, email addresses, roles, and IDs.",
    {
      teamId: z.string().describe("Team ID"),
    },
    async ({ teamId }) => {
      try {
        const client = await graphService.getClient();
        const response = (await client
          .api(`/teams/${teamId}/members`)
          .get()) as GraphApiResponse<ConversationMember>;

        if (!response?.value?.length) {
          return {
            content: [
              {
                type: "text",
                text: "No members found in this team.",
              },
            ],
          };
        }

        const memberList: MemberSummary[] = response.value.map((member: ConversationMember) => ({
          id: member.id,
          displayName: member.displayName,
          roles: member.roles,
        }));

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(memberList, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // Search users for @mentions
  server.tool(
    "search_users_for_mentions",
    "Search for users to mention in messages. Returns users with their display names, email addresses, and mention IDs.",
    {
      query: z.string().describe("Search query (name or email)"),
      limit: z
        .number()
        .min(1)
        .max(50)
        .optional()
        .default(10)
        .describe("Maximum number of results to return"),
    },
    async ({ query, limit }) => {
      try {
        const users = await searchUsers(graphService, query, limit);

        if (users.length === 0) {
          return {
            content: [
              {
                type: "text",
                text: `No users found matching "${query}".`,
              },
            ],
          };
        }

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  query,
                  totalResults: users.length,
                  users: users.map((user: UserInfo) => ({
                    id: user.id,
                    displayName: user.displayName,
                    userPrincipalName: user.userPrincipalName,
                    mentionText:
                      user.userPrincipalName?.split("@")[0] ||
                      user.displayName.toLowerCase().replace(/\s+/g, ""),
                  })),
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text",
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );
}
