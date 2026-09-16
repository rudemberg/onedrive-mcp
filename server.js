import express from "express";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

const app = express();
app.use(express.json());

// Liberação de CORS para permitir a conexão do Gemini
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }
  next();
});

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI;
let REFRESH_TOKEN = process.env.REFRESH_TOKEN || null;

let currentAccessToken = null;
let tokenExpiresAt = 0;

// Renovação automática de Token
async function getAccessToken() {
  if (currentAccessToken && Date.now() < tokenExpiresAt - 60000) {
    return currentAccessToken;
  }
  const tokenToUse = process.env.REFRESH_TOKEN || REFRESH_TOKEN;
  if (!tokenToUse) {
    throw new Error("OneDrive não autorizado. Acesse a rota /auth no navegador para autenticar.");
  }
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    grant_type: "refresh_token",
    refresh_token: tokenToUse,
  });

  const response = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString(),
  });

  const data = await response.json();
  if (data.error) {
    throw new Error(`Erro ao renovar token: ${data.error_description || data.error}`);
  }

  currentAccessToken = data.access_token;
  if (data.refresh_token) {
    REFRESH_TOKEN = data.refresh_token;
  }
  tokenExpiresAt = Date.now() + (data.expires_in * 1000);
  return currentAccessToken;
}

// Rotas de Autenticação inicial (OAuth)
app.get("/auth", (req, res) => {
  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(REDIRECT_URI)}&response_mode=query&scope=offline_access%20Files.ReadWrite%20User.Read`;
  res.redirect(authUrl);
});

app.get("/callback", async (req, res) => {
  const code = req.query.code;
  if (!code) return res.status(400).send("Código não fornecido.");

  try {
    const params = new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      code,
      redirect_uri: REDIRECT_URI,
      grant_type: "authorization_code",
    });

    const resp = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    });

    const data = await resp.json();
    if (data.error) {
      return res.status(500).send(`Erro na autenticação: ${data.error_description || data.error}`);
    }

    REFRESH_TOKEN = data.refresh_token;
    currentAccessToken = data.access_token;
    tokenExpiresAt = Date.now() + (data.expires_in * 1000);

    res.send(`
      <html>
        <body style="font-family: sans-serif; text-align: center; padding: 40px;">
          <h2 style="color: #2e7d32;">OneDrive conectado com sucesso!</h2>
          <p>O servidor já está autorizado a acessar seus arquivos.</p>
          <p>Copie este Refresh Token para salvar nas variáveis do Render:</p>
          <textarea style="width: 80%; height: 100px; font-size: 11px;">${REFRESH_TOKEN}</textarea>
        </body>
      </html>
    `);
  } catch (err) {
    res.status(500).send("Erro: " + err.message);
  }
});

// Fábrica do servidor MCP por sessão
function setupMcpServer() {
  const server = new Server(
    { name: "onedrive-mcp", version: "1.0.0" },
    { capabilities: { tools: {} } }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: [
        {
          name: "onedrive_list_files",
          description: "Lista arquivos e pastas no seu OneDrive.",
          inputSchema: {
            type: "object",
            properties: {
              folder_path: { type: "string", description: "Caminho da pasta (opcional, vazio para a raiz)" },
            },
          },
        },
        {
          name: "onedrive_search_files",
          description: "Pesquisa arquivos por palavra-chave no OneDrive.",
          inputSchema: {
            type: "object",
            properties: {
              query: { type: "string", description: "Termo de busca" },
            },
            required: ["query"],
          },
        },
        {
          name: "onedrive_read_file",
          description: "Lê o conteúdo em texto de um arquivo no OneDrive.",
          inputSchema: {
            type: "object",
            properties: {
              item_id: { type: "string", description: "ID do item/arquivo retornado pela listagem ou busca" },
            },
            required: ["item_id"],
          },
        },
        {
          name: "onedrive_upload_file",
          description: "Cria ou atualiza um arquivo de texto no OneDrive.",
          inputSchema: {
            type: "object",
            properties: {
              file_path: { type: "string", description: "Caminho e nome do arquivo (ex: 'Notas/resumo.txt')" },
              content: { type: "string", description: "Conteúdo em texto a ser salvo" },
            },
            required: ["file_path", "content"],
          },
        },
      ],
    };
  });

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const token = await getAccessToken();
    const { name, arguments: args } = request.params;

    if (name === "onedrive_list_files") {
      const path = args?.folder_path ? `root:/${encodeURIComponent(args.folder_path)}:/children` : "root/children";
      const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/${path}`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const data = await res.json();
      return { content: [{ type: "text", text: JSON.stringify(data.value || data, null, 2) }] };
    }

    if (name === "onedrive_search_files") {
      const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root/search(q='${encodeURIComponent(args.query)}')`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const data = await res.json();
      return { content: [{ type: "text", text: JSON.stringify(data.value || data, null, 2) }] };
    }

    if (name === "onedrive_read_file") {
      const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${args.item_id}/content`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const text = await res.text();
      return { content: [{ type: "text", text }] };
    }

    if (name === "onedrive_upload_file") {
      const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(args.file_path)}:/content`, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "text/plain",
        },
        body: args.content,
      });
      const data = await res.json();
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }

    throw new Error(`Ferramenta desconhecida: ${name}`);
  });

  return server;
}

// Endpoints SSE
const transports = new Map();

app.get("/sse", async (req, res) => {
  const transport = new SSEServerTransport("/messages", res);
  transports.set(transport.sessionId, transport);
  const server = setupMcpServer();
  await server.connect(transport);
  req.on("close", () => transports.delete(transport.sessionId));
});

app.post("/messages", async (req, res) => {
  const sessionId = req.query.sessionId;
  const transport = transports.get(sessionId);
  if (!transport) return res.status(404).send("Sessão não encontrada.");
  await transport.handlePostMessage(req, res);
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`Servidor OneDrive MCP rodando na porta ${PORT}`);
});
