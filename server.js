
import express from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';

const app = express();

app.use(express.json({ limit: '1mb' }));

const port = Number(process.env.PORT || 3000);

// Preserva o token de verificação do domínio.
app.get('/.well-known/openai-apps-challenge', (_req, res) => {
  res.type('text/plain').send(
    '0FKoAHyHuiCTFG-gWVF5CvyPPJq0GQ4YmfujdB5KsfA'
  );
});

// Cria um servidor MCP com apenas uma ferramenta de teste.
function createServer() {
  const server = new McpServer({
    name: 'arena-executiva-mcp',
    version: '1.0.5'
  });

  server.registerTool(
    'verificar_conexao',
    {
      title: 'Verificar conexão Arena Executiva',
      description:
        'Verifica se o servidor MCP da Arena Executiva está funcionando.',
      inputSchema: {}
    },
    async () => ({
      content: [
        {
          type: 'text',
          text: 'Conexão MCP da Arena Executiva funcionando.'
        }
      ]
    })
  );

  return server;
}

// Página de disponibilidade.
app.get('/', (_req, res) => {
  res.json({
    name: 'Arena Executiva MCP',
    status: 'ok',
    endpoint: '/mcp'
  });
});

// Endpoint MCP.
app.post('/mcp', async (req, res) => {
  const server = createServer();

  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined
  });

  let closed = false;

  async function cleanup() {
    if (closed) return;
    closed = true;

    await Promise.allSettled([
      transport.close(),
      server.close()
    ]);
  }

  res.once('finish', () => void cleanup());
  res.once('close', () => void cleanup());

  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (error) {
    console.error('MCP request failed:', error);

    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: '2.0',
        error: {
          code: -32603,
          message: 'Internal server error'
        },
        id: null
      });
    } else if (!res.writableEnded) {
      res.end();
    }

    await cleanup();
  }
});

app.get('/mcp', (_req, res) => {
  res.status(405).send('Use POST for MCP requests');
});

app.delete('/mcp', (_req, res) => {
  res.status(405).send('Stateless server');
});

app.listen(port, '0.0.0.0', () => {
  console.log(`Arena Executiva MCP on ${port}`);
});
