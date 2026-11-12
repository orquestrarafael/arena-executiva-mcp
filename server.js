
import express from 'express';
import PDFDocument from 'pdfkit';
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

// TOKEN DE VERIFICAÇÃO DO DOMÍNIO

app.get('/.well-known/openai-apps-challenge', (_req, res) => {
  res.type('text/plain').send(
    '0FKoAHyHuiCTFG-gWVF5CvyPPJq0GQ4YmfujdB5KsfA'
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

function saveDownload(buffer, kind) {
  purgeDownloads();

  if (downloads.size >= MAX_DOWNLOADS) {
    throw new Error('Limite temporário de arquivos atingido.');
  }

  const id = randomUUID();

  downloads.set(id, {
    buffer,
    kind,
    expiresAt: Date.now() + DOWNLOAD_TTL_MS
  });

  return `${PUBLIC_URL}/downloads/${id}`;
}

// GERAÇÃO DE PDF

function makePdf({
  nome_plano,
  provocacao,
  comprovacao,
  objetivos,
  acompanhamento,
  proximos_passos
}) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: 'A4',
      margin: 50,
      bufferPages: true,
      info: {
        Title: `Arena Executiva - ${nome_plano}`,
        Author: 'Rafael Luis dos Santos'
      }
    });

    const chunks = [];

    doc.on('data', chunk => chunks.push(chunk));
    doc.on('error', reject);
    doc.on('end', () => resolve(Buffer.concat(chunks)));

    const blue = '#17365D';
    const gray = '#455468';

    const section = title => {
      doc.moveDown(1.1)
        .font('Helvetica-Bold')
        .fontSize(15)
        .fillColor(blue)
        .text(title);

      doc.moveDown(0.35);
    };

    const paragraph = value => {
      doc.font('Helvetica')
        .fontSize(10.5)
        .fillColor('#202936')
        .text(value || 'Não informado.', {
          lineGap: 4
        });

      doc.moveDown(0.45);
    };

    doc.font('Helvetica-Bold')
      .fontSize(24)
      .fillColor(blue)
      .text('ARENA EXECUTIVA');

    doc.moveDown(0.25)
      .font('Helvetica')
      .fontSize(11)
      .fillColor(gray)
      .text(
        'Planejamento Estratégico | Método de Rafael Luis dos Santos'
      );

    doc.moveDown(1.5)
      .font('Helvetica-Bold')
      .fontSize(18)
      .fillColor(blue)
      .text(nome_plano);

    doc.moveDown(0.5)
      .font('Helvetica')
      .fontSize(10)
      .fillColor(gray)
      .text(
        `Documento gerado em ${new Date().toLocaleDateString(
          'pt-BR',
          { timeZone: 'America/Sao_Paulo' }
        )}`
      );

    section('1. PROVOCAR - Propósito estratégico');
    paragraph(provocacao);

    section('2. COMPROVAR - Diagnóstico e evidências');
    paragraph(comprovacao);

    section('3. DEFINIR - Objetivos e indicadores (OKRs)');

    objetivos.forEach((item, i) => {
      doc.font('Helvetica-Bold')
        .fontSize(11)
        .fillColor(blue)
        .text(`${i + 1}. ${item.objetivo}`);

      doc.moveDown(0.2);

      paragraph(
        `Indicador (KR): ${item.indicador || 'Não informado'}\n` +
        `Meta: ${item.meta || 'Não informada'}\n` +
        `Responsável: ${item.responsavel || 'Não informado'}\n` +
        `Prazo: ${item.prazo || 'Não informado'}`
      );
    });

    section('4. ACOMPANHAR - Governança e decisões');
    paragraph(acompanhamento);

    section('Próximos passos');
    paragraph(proximos_passos);

    const pages = doc.bufferedPageRange();

    for (let i = 0; i < pages.count; i++) {
      doc.switchToPage(i);

      doc.font('Helvetica')
        .fontSize(8)
        .fillColor(gray)
        .text(
          `Arena Executiva | Página ${i + 1} de ${pages.count}`,
          50,
          doc.page.height - 34,
          {
            align: 'center',
            width: doc.page.width - 100
          }
        );
    }

    doc.end();
  });
}

// SERVIDOR MCP

function createServer() {
  const server = new McpServer({
    name: 'arena-executiva-mcp',
    version: '1.0.6'
  });

  // FERRAMENTA 1: PDF

  server.registerTool(
    'gerar_pdf_planejamento',
    {
      title: 'Gerar PDF do planejamento estratégico',

      description:
        'Gera o planejamento estratégico da Arena Executiva ' +
        'em PDF, incluindo as quatro etapas, OKRs, ' +
        'governança e próximos passos.',

      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      },

      inputSchema: {
        nome_plano: z.string().min(1).max(120),
        provocacao: z.string().min(1).max(12000),
        comprovacao: z.string().min(1).max(12000),

        objetivos: z.array(z.object({
          objetivo: z.string().min(1).max(1000),
          indicador: z.string().max(1000).default(''),
          meta: z.string().max(1000).default(''),
          responsavel: z.string().max(1000).default(''),
          prazo: z.string().max(1000).default('')
        })).min(1).max(30),

        acompanhamento: z.string().min(1).max(12000),
        proximos_passos: z.string().min(1).max(12000)
      }
    },
    async args => {
      try {
        const buffer = await makePdf(args);
        const url = saveDownload(buffer, 'pdf');

        return {
          content: [{
            type: 'text',
            text:
              'PDF gerado com sucesso. ' +
              `Link válido por 10 minutos: ${url}`
          }]
        };
      } catch (error) {
        console.error('PDF generation failed:', error);

        return {
          isError: true,
          content: [{
            type: 'text',
            text: 'Não foi possível gerar o PDF.'
          }]
        };
      }
    }
  );

  // FERRAMENTA 2: EXCEL

  server.registerTool(
    'gerar_excel_okrs',
    {
      title: 'Gerar Excel de OKRs',

      description:
        'Gera planilha Excel com objetivos, indicadores, ' +
        'metas, responsáveis e prazos.',

      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      },

      inputSchema: {
        nome_plano: z.string().min(1).max(120),

        objetivos: z.array(z.object({
          objetivo: z.string().min(1),
          indicador: z.string().default(''),
          meta: z.string().default(''),
          responsavel: z.string().default(''),
          prazo: z.string().default('')
        })).min(1).max(100)
      }
    },
    async ({ nome_plano, objetivos }) => {
      try {
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

        const url = saveDownload(buffer, 'xlsx');

        return {
          content: [{
            type: 'text',
            text:
              'Excel gerado com sucesso. ' +
              `Link válido por 10 minutos: ${url}`
          }]
        };
      } catch (error) {
        console.error('Excel generation failed:', error);

        return {
          isError: true,
          content: [{
            type: 'text',
            text: 'Não foi possível gerar o Excel.'
          }]
        };
      }
    }
  );

  return server;
}

// STATUS DO SERVIDOR

app.get('/', (_req, res) => {
  res.json({
    name: 'Arena Executiva MCP',
    status: 'ok',
    endpoint: '/mcp',
    tools: [
      'gerar_pdf_planejamento',
      'gerar_excel_okrs'
    ]
  });
});

// DOWNLOAD DOS ARQUIVOS

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
      entry.kind === 'xlsx'
        ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        : 'application/pdf',

    'Content-Disposition':
      entry.kind === 'xlsx'
        ? 'attachment; filename="arena-okrs.xlsx"'
        : 'attachment; filename="arena-executiva-planejamento.pdf"',

    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff'
  });

  res.send(entry.buffer);
});

// COMUNICAÇÃO MCP

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
