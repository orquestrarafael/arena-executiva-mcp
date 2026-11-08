import express from 'express';
import ExcelJS from 'exceljs';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { z } from 'zod';

const app = express();
app.get('/.well-known/openai-apps-challenge', (req, res) => {
  res.type('text/plain').send('8tHwZfTxR3S4KfJ3dyM_nqJjEv5pVXhVnDY2oU1WYiO');
});
app.get('/.well-known/openai-apps-challenge', (req, res) => {
  res.type('text/plain').send('8tHwZfTxR3S4KfJ3dyM_nqJjEv5pVXhVnDY2oUlWYio');
});
app.use(express.json({ limit: '1mb' }));
const port = Number(process.env.PORT || 3000);

function safeCell(value) {
  const s = String(value ?? '');
  return /^[=+\-@\t\r]/.test(s) ? `'${s}` : s;
}

function createServer() {
  const server = new McpServer({name: 'arena-executiva-mcp', version: '1.0.0'});
  server.registerTool('gerar_excel_okrs', {
    title: 'Gerar Excel de OKRs',
    description: 'Gera uma planilha Excel com objetivos, indicadores, metas, responsáveis e prazos. Retorna um link para baixar o arquivo.',
    inputSchema: {
      nome_plano: z.string().min(1).max(120).describe('Nome do plano estratégico'),
      objetivos: z.array(z.object({
        objetivo: z.string().min(1),
        indicador: z.string().default(''),
        meta: z.string().default(''),
        responsavel: z.string().default(''),
        prazo: z.string().default('')
      })).min(1).max(100)
    }
  }, async ({ nome_plano, objetivos }) => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('OKRs');
    ws.columns = [
      {header: 'Plano', key: 'plano', width: 30},
      {header: 'Objetivo', key: 'objetivo', width: 45},
      {header: 'Indicador (KR)', key: 'indicador', width: 38},
      {header: 'Meta', key: 'meta', width: 25},
      {header: 'Responsável', key: 'responsavel', width: 28},
      {header: 'Prazo', key: 'prazo', width: 22}
    ];
    ws.getRow(1).font = {bold: true, color: {argb: 'FFFFFFFF'}};
    ws.getRow(1).fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF17365D'}};
    for (const o of objetivos) ws.addRow({plano: safeCell(nome_plano), objetivo: safeCell(o.objetivo), indicador: safeCell(o.indicador), meta: safeCell(o.meta), responsavel: safeCell(o.responsavel), prazo: safeCell(o.prazo)});
    ws.autoFilter = 'A1:F1';
    ws.views = [{state: 'frozen', ySplit: 1}];
    const buffer = await wb.xlsx.writeBuffer();
    const base64 = Buffer.from(buffer).toString('base64');
    return {content: [{type: 'resource_link', uri: `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`, name: 'arena-okrs.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', description: 'Arquivo Excel com os OKRs'}]};
  });
  return server;
}

app.get('/', (_req, res) => res.json({name: 'Arena Executiva MCP', status: 'ok', endpoint: '/mcp'}));
app.post('/mcp', async (req, res) => {
  const server = createServer();
  const transport = new StreamableHTTPServerTransport({sessionIdGenerator: undefined});
  res.on('close', () => {transport.close().catch(() => {}); server.close().catch(() => {});});
  try {await server.connect(transport); await transport.handleRequest(req, res, req.body);}
  catch (err) {console.error(err); if (!res.headersSent) res.status(500).json({jsonrpc: '2.0', error: {code: -32603, message: 'Internal server error'}, id: null});}
});
app.get('/mcp', (_req, res) => res.status(405).send('Use POST for MCP requests'));
app.delete('/mcp', (_req, res) => res.status(405).send('Stateless server'));
app.listen(port, '0.0.0.0', () => console.log(`Arena Executiva MCP on ${port}`));
