import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { PublicClientApplication, type Configuration } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { cachePlugin } from "../msal-cache.js";

// Microsoft Graph CLI app ID (default public client)
const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const AUTHORITY = "https://login.microsoftonline.com/common";

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string | undefined;
  displayName?: string | undefined;
  expiresAt?: string | undefined;
}

interface StoredAuthInfo {
  clientId: string;
  authenticated: boolean;
  timestamp: string;
  expiresAt?: string;
  account?: string;
}

// Scopes for delegated (user) authentication
const DELEGATED_SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "ChannelMessage.Send",
  "TeamMember.Read.All",
  "Chat.ReadBasic",
  "Chat.ReadWrite",
];

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private msalClient: PublicClientApplication | undefined;
  private readonly authPath = join(homedir(), ".msgraph-mcp-auth.json");
  private isInitialized = false;
  private authInfo: StoredAuthInfo | undefined;

  static getInstance(): GraphService {
    if (!GraphService.instance) {
      GraphService.instance = new GraphService();
    }
    return GraphService.instance;
  }

  private async initializeClient(): Promise<void> {
    if (this.isInitialized) return;

    try {
      // Check if we have stored auth info
      const authData = await fs.readFile(this.authPath, "utf8");
      this.authInfo = JSON.parse(authData);

      if (this.authInfo?.authenticated) {
        // Create MSAL client with persistent cache
        const msalConfig: Configuration = {
          auth: {
            clientId: CLIENT_ID,
            authority: AUTHORITY,
          },
          cache: {
            cachePlugin,
          },
        };

        this.msalClient = new PublicClientApplication(msalConfig);

        // Create Graph client using MSAL for token acquisition
        this.client = Client.initWithMiddleware({
          authProvider: {
            getAccessToken: async () => {
              if (!this.msalClient) {
                throw new Error("MSAL client not initialized");
              }

              // Try to get token silently using cached refresh token
              const accounts = await this.msalClient.getTokenCache().getAllAccounts();

              if (accounts.length > 0) {
                try {
                  const result = await this.msalClient.acquireTokenSilent({
                    account: accounts[0],
                    scopes: DELEGATED_SCOPES,
                  });

                  if (result?.accessToken) {
                    return result.accessToken;
                  }
                } catch (error) {
                  console.error("Silent token acquisition failed, re-authentication required:", error);
                  throw new Error("Token refresh failed. Please re-authenticate with: npx teams-mcp authenticate");
                }
              }

              throw new Error("No cached account found. Please authenticate with: npx teams-mcp authenticate");
            },
          },
        });

        this.isInitialized = true;
      }
    } catch (error) {
      // If no auth file exists, that's okay - just not authenticated
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") {
        console.error("Failed to initialize Graph client:", error);
      }
    }
  }

  async getAuthStatus(): Promise<AuthStatus> {
    await this.initializeClient();

    if (!this.client) {
      return { isAuthenticated: false };
    }

    try {
      const me = await this.client.api("/me").get();
      return {
        isAuthenticated: true,
        userPrincipalName: me?.userPrincipalName ?? undefined,
        displayName: me?.displayName ?? undefined,
        expiresAt: this.authInfo?.expiresAt,
      };
    } catch (error) {
      console.error("Error getting user info:", error);
      return { isAuthenticated: false };
    }
  }

  async getClient(): Promise<Client> {
    await this.initializeClient();

    if (!this.client) {
      throw new Error(
        "Not authenticated. Please run the authentication CLI tool first: npx teams-mcp authenticate"
      );
    }
    return this.client;
  }

  isAuthenticated(): boolean {
    return !!this.client && this.isInitialized;
  }
}
