import express from "express";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

const app = express();
app.use(express.json());

// Log de requisições
app.use((req, res, next) => {
  console.log(`[REQUISIÇÃO] ${req.method} ${req.url}`);
  next();
});

// Liberação completa de CORS
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, HEAD, OPTIONS");
  res.header("Access-Control-Allow-Headers", "*");
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }
  next();
});

// Resposta na raiz
app.all("/", (req, res) => {
  res.status(200).send("Servidor OneDrive MCP ativo!");
});

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI;
let REFRESH_TOKEN = process.env.REFRESH_TOKEN || null;

let currentAccessToken = null;
let tokenExpiresAt = 0;

// Renovação de Token Microsoft
async function getAccessToken() {
  if (currentAccessToken && Date.now() < tokenExpiresAt - 60000) {
    return currentAccessToken;
  }
  const tokenToUse = process.env.REFRESH_TOKEN || REFRESH_TOKEN;
  if (!tokenToUse) {
    throw new Error("OneDrive não autorizado. Acesse /auth para autorizar.");
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

// Rotas de Autorização
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
        </body>
      </html>
    `);
  } catch (err) {
    res.status(500).send("Erro: " + err.message);
  }
});

// Lista de Ferramentas disponíveis
const TOOLS = [
  {
    name: "onedrive_list_files",
    description: "Lista arquivos e pastas no seu OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        folder_path: { type: "string", description: "Caminho da pasta (opcional, vazio para raiz)" },
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
    description: "Lê o conteúdo em texto de um único arquivo no OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        item_id: { type: "string", description: "ID do item/arquivo" },
      },
      required: ["item_id"],
    },
  },
  {
    name: "onedrive_read_multiple_files",
    description: "Lê o conteúdo de múltiplos arquivos de uma só vez a partir de uma lista de IDs. Sempre prefira esta ferramenta quando precisar ler mais de um arquivo para reduzir confirmações.",
    inputSchema: {
      type: "object",
      properties: {
        item_ids: {
          type: "array",
          items: { type: "string" },
          description: "Lista de IDs dos arquivos para ler em lote",
        },
      },
      required: ["item_ids"],
    },
  },
  {
    name: "onedrive_read_folder_files",
    description: "Lê o conteúdo de todos os arquivos de uma pasta de uma só vez em uma única chamada.",
    inputSchema: {
      type: "object",
      properties: {
        folder_path: { type: "string", description: "Caminho da pasta (opcional, vazio para a raiz)" },
      },
    },
  },
  {
    name: "onedrive_upload_file",
    description: "Cria ou atualiza um arquivo de texto no OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        file_path: { type: "string", description: "Caminho e nome do arquivo (ex: 'resumo.txt')" },
        content: { type: "string", description: "Conteúdo em texto a ser salvo" },
      },
      required: ["file_path", "content"],
    },
  },
];

// Execução das Ferramentas
async function executeTool(name, args) {
  const token = await getAccessToken();

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

  // Leitura em lote de múltiplos arquivos por ID
  if (name === "onedrive_read_multiple_files") {
    const filesData = await Promise.all(
      (args.item_ids || []).map(async (id) => {
        try {
          const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${id}/content`, {
            headers: { Authorization: `Bearer ${token}` },
          });
          const text = await res.text();
          return { item_id: id, content: text };
        } catch (err) {
          return { item_id: id, error: err.message };
        }
      })
    );
    return { content: [{ type: "text", text: JSON.stringify(filesData, null, 2) }] };
  }

  // Leitura direta de todos os arquivos de uma pasta
  if (name === "onedrive_read_folder_files") {
    const path = args?.folder_path ? `root:/${encodeURIComponent(args.folder_path)}:/children` : "root/children";
    const listRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/${path}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const listData = await listRes.json();
    const files = (listData.value || []).filter((item) => item.file);

    const filesContent = await Promise.all(
      files.map(async (f) => {
        try {
          const contentRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${f.id}/content`, {
            headers: { Authorization: `Bearer ${token}` },
          });
          const text = await contentRes.text();
          return { id: f.id, name: f.name, content: text };
        } catch (err) {
          return { id: f.id, name: f.name, error: err.message };
        }
      })
    );
    return { content: [{ type: "text", text: JSON.stringify(filesContent, null, 2) }] };
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
}

// Configuração do Servidor MCP SDK
function setupMcpServer() {
  const server = new Server(
    { name: "onedrive-mcp", version: "1.0.0" },
    { capabilities: { tools: {} } }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));
  server.setRequestHandler(CallToolRequestSchema, async (req) => {
    return await executeTool(req.params.name, req.params.arguments);
  });

  return server;
}

const transports = new Map();
let latestTransport = null;

// Rota GET /sse
app.get("/sse", async (req, res) => {
  if (req.method === "HEAD") {
    res.setHeader("Content-Type", "text/event-stream");
    res.setHeader("Cache-Control", "no-cache");
    return res.status(200).end();
  }

  res.setHeader("X-Accel-Buffering", "no");
  const fullMessagesUrl = "https://onedrive-mcp-p2pe.onrender.com/messages";
  const transport = new SSEServerTransport(fullMessagesUrl, res);
  transports.set(transport.sessionId, transport);
  latestTransport = transport;

  const server = setupMcpServer();
  await server.connect(transport);

  console.log(`[MCP] Conexão SSE aberta: ${transport.sessionId}`);

  req.on("close", () => {
    console.log(`[MCP] Conexão SSE encerrada: ${transport.sessionId}`);
    transports.delete(transport.sessionId);
    if (latestTransport === transport) latestTransport = null;
  });
});

// Responde a requisições POST
app.post(["/sse", "/messages"], async (req, res) => {
  const msg = req.body;
  const sessionId = req.query.sessionId;
  const transport = sessionId ? transports.get(sessionId) : latestTransport;

  if (transport && sessionId) {
    return await transport.handlePostMessage(req, res);
  }

  if (msg) {
    if (msg.method === "initialize") {
      return res.json({
        jsonrpc: "2.0",
        id: msg.id,
        result: {
          protocolVersion: "2024-11-05",
          capabilities: { tools: {} },
          serverInfo: { name: "onedrive-mcp", version: "1.0.0" },
        },
      });
    }

    if (msg.method === "notifications/initialized") {
      return res.status(200).end();
    }

    if (msg.method === "tools/list") {
      return res.json({
        jsonrpc: "2.0",
        id: msg.id,
        result: { tools: TOOLS },
      });
    }

    if (msg.method === "tools/call") {
      try {
        const result = await executeTool(msg.params.name, msg.params.arguments);
        return res.json({ jsonrpc: "2.0", id: msg.id, result });
      } catch (err) {
        return res.json({
          jsonrpc: "2.0",
          id: msg.id,
          error: { code: -32000, message: err.message },
        });
      }
    }

    if (msg.method === "ping") {
      return res.json({ jsonrpc: "2.0", id: msg.id, result: {} });
    }
  }

  res.status(200).end();
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`Servidor OneDrive MCP rodando na porta ${PORT}`);
});
