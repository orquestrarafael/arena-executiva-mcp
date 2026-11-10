
import express from 'express';
import ExcelJS from 'exceljs';
import { randomUUID } from 'node:crypto';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { z } from 'zod';

const app = express();
app.use(express.json({ limit: '1mb' }));

const port = Number(process.env.PORT || 3000);
const PUBLIC_URL = (
  process.env.PUBLIC_URL ||
  'https://arena-executiva-mcp-1.onrender.com'
).replace(/\/$/, '');

const DOWNLOAD_TTL_MS = 10 * 60 * 1000;
const MAX_DOWNLOADS = 100;
const downloads = new Map();

app.get('/.well-known/openai-apps-challenge', (_req, res) => {
  res.type('text/plain').send(
    '8tHwZfTxR3S4KfJ3dyM_nqJjEv5pVXhVnDY2oU1WYiO'
  );
});

function safeCell(value) {
  const s = String(value ?? '');
  return /^[=+\-@\t\r]/.test(s) ? `'${s}` : s;
}

function purgeDownloads() {
  const now = Date.now();

  for (const [id, entry] of downloads) {
    if (entry.expiresAt <= now) {
      downloads.delete(id);
    }
  }
}

function createServer() {
  const server = new McpServer({
    name: 'arena-executiva-mcp',
    version: '1.0.1'
  });

  server.registerTool('gerar_excel_okrs', {
    title: 'Gerar Excel de OKRs',
    description:
      'Gera uma planilha Excel com objetivos, indicadores, metas, responsáveis e prazos. Retorna uma URL HTTPS temporária para baixar o arquivo.',

    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      openWorldHint: false
    },

    inputSchema: {
      nome_plano: z.string()
        .min(1)
        .max(120)
        .describe('Nome do plano estratégico'),

      objetivos: z.array(
        z.object({
          objetivo: z.string().min(1),
          indicador: z.string().default(''),
          meta: z.string().default(''),
          responsavel: z.string().default(''),
          prazo: z.string().default('')
        })
      ).min(1).max(100)
    }
  }, async ({ nome_plano, objetivos }) => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('OKRs');

    ws.columns = [
      { header: 'Plano', key: 'plano', width: 30 },
      { header: 'Objetivo', key: 'objetivo', width: 45 },
      { header: 'Indicador (KR)', key: 'indicador', width: 38 },
      { header: 'Meta', key: 'meta', width: 25 },
      { header: 'Responsável', key: 'responsavel', width: 28 },
      { header: 'Prazo', key: 'prazo', width: 22 }
    ];

    ws.getRow(1).font = {
      bold: true,
      color: { argb: 'FFFFFFFF' }
    };

    ws.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF17365D' }
    };

    for (const o of objetivos) {
      ws.addRow({
        plano: safeCell(nome_plano),
        objetivo: safeCell(o.objetivo),
        indicador: safeCell(o.indicador),
        meta: safeCell(o.meta),
        responsavel: safeCell(o.responsavel),
        prazo: safeCell(o.prazo)
      });
    }

    ws.autoFilter = 'A1:F1';
    ws.views = [{
      state: 'frozen',
      ySplit: 1
    }];

    const buffer = Buffer.from(
      await wb.xlsx.writeBuffer()
    );

    purgeDownloads();

    if (downloads.size >= MAX_DOWNLOADS) {
      return {
        isError: true,
        content: [{
          type: 'text',
          text: 'Limite temporário de arquivos atingido. Tente novamente mais tarde.'
        }]
      };
    }

    const id = randomUUID();

    downloads.set(id, {
      buffer,
      expiresAt: Date.now() + DOWNLOAD_TTL_MS
    });

    const url = `${PUBLIC_URL}/downloads/${id}`;

    return {
      content: [{
        type: 'text',
        text:
          `Excel de OKRs gerado. Baixe o arquivo neste link HTTPS (válido por 10 minutos): ${url}`
      }]
    };
  });

  return server;
}

app.get('/', (_req, res) => {
  res.json({
    name: 'Arena Executiva MCP',
    status: 'ok',
    endpoint: '/mcp'
  });
});

app.get('/downloads/:id', (req, res) => {
  const entry = downloads.get(req.params.id);

  if (!entry || entry.expiresAt <= Date.now()) {
    downloads.delete(req.params.id);

    return res.status(404).send(
      'Arquivo não encontrado ou link expirado.'
    );
  }

  res.set({
    'Content-Type':
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition':
      'attachment; filename="arena-okrs.xlsx"',
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff'
  });

  res.send(entry.buffer);
});

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

  res.once('finish', () => {
    void cleanup();
  });

  res.once('close', () => {
    void cleanup();
  });

  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err) {
    console.error('MCP request failed:', err);

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
