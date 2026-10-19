import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { createRequire } from 'node:module';

import { SECTIONS, type Section, readWorkbook } from './workbook/read.js';
import { scanWorkbooks, whereUsed } from './workbook/scan.js';

const { version } = createRequire(import.meta.url)('../package.json') as { version: string };

const INSTRUCTIONS = `SAP Analysis for Office (AfO) workbooks, read offline: no Excel, no SAP logon.
The tools unzip .xlsx/.xlsm files and parse the AfO definition Analysis for Office stores in them.

- afo_read_workbook: everything about one workbook.
- afo_scan_workbooks: inventory of a folder: profiles, which query is used where, recurring layouts, outliers.
- afo_where_used: every place a query, planning function, variable or characteristic occurs across a folder.

What a workbook stores and what it does not:
- Data sources carry their saved navigation state (filters, axes, hierarchies) only when AfO saved one;
  stateStored=false means the data source opens in the query's default view.
- Query variable values are stored only when the workbook setting StorePromptsInDocument is True.
  Planning function variables are always stored with the planning object.
- Structure members and some filter values appear as BW-internal IDs (e.g. 00O2TP…); resolving them
  needs the query definition from the BW system (bw-modeling-mcp: bw_get_query).
- Sheet references use the visible sheet name; zero-based AfO coordinates are converted to A1.`;

const TOOLS = [
  {
    name: 'afo_read_workbook',
    description:
      'Describe one Analysis for Office workbook without Excel: settings (system, AfO version, change mode on open, prompt settings), sheets, data sources with their saved state (InfoProvider, query, filter per characteristic, axes, hierarchies), planning functions and sequences with bound data source, aggregation level, filter, variable values and cell/filter references, crosstabs and formula items with sheet and range, SAP cell formulas (SAPGetData, SAPGetVariable, SAPSetFilterComponent, …) with parsed arguments, VBA modules with procedures and the AfO API calls in them (SAPExecuteCommand, SAPExecutePlanningFunction, SAPSetVariable, …), buttons with the macro they run. Lists broken links as issues: a formula or macro naming an undefined data source alias, a button whose macro does not exist, a crosstab without its data source.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path of the workbook (.xlsx, .xlsm).' },
        sections: {
          type: 'array',
          items: { type: 'string', enum: [...SECTIONS] },
          description: `Sections to include (default: all). The summary and the issues are always included. One of: ${SECTIONS.join(', ')}.`,
        },
        macro_source: { type: 'boolean', description: 'Include the full VBA source of every module (default false: procedures and AfO calls only).' },
        max_formulas_per_sheet: { type: 'number', description: 'Cap on SAP formulas listed per sheet (default 200). Counts stay complete.' },
      },
      required: ['path'],
    },
  },
  {
    name: 'afo_scan_workbooks',
    description:
      'Inventory a folder of Analysis for Office workbooks without Excel: a short profile per workbook (system, AfO version, data sources by alias, planning objects, crosstabs, SAP formula cells, macros, issues), then across all workbooks: which query/data source is used in which workbooks, the same for planning functions and variables, recurring layouts (workbooks binding the same aliases to the same shared queries), and outliers (different system or AfO version than the majority, unreadable files, workbooks without AfO definition).',
    inputSchema: {
      type: 'object',
      properties: {
        folder: { type: 'string', description: 'Absolute path of the folder.' },
        recursive: { type: 'boolean', description: 'Include subfolders (default false).' },
        limit: { type: 'number', description: 'Maximum number of workbooks to read (default 2000).' },
        detail: {
          type: 'string',
          enum: ['full', 'compact'],
          description: 'full: a profile per workbook and complete usage lists. compact: no profiles, usage lists capped at five workbooks each. Default: full up to 50 workbooks, compact above. Use afo_read_workbook or afo_where_used to drill down.',
        },
      },
      required: ['folder'],
    },
  },
  {
    name: 'afo_where_used',
    description:
      'Find every place a technical name occurs across a folder of Analysis for Office workbooks: as data source (query or InfoProvider), data source alias, planning function or sequence, aggregation level, planning filter, variable (prompt, planning variable, cell reference, SAPGetVariable), characteristic in filters and axes, argument of SAP cell formulas and formula items, and in VBA code. Use it to see what breaks when a query, variable or planning function changes. Matching is exact on the whole name, case-insensitive.',
    inputSchema: {
      type: 'object',
      properties: {
        folder: { type: 'string', description: 'Absolute path of the folder.' },
        name: { type: 'string', description: 'Technical name: query, InfoProvider, planning function or sequence, variable, characteristic, or a data source alias such as DS_1.' },
        recursive: { type: 'boolean', description: 'Include subfolders (default false).' },
      },
      required: ['folder', 'name'],
    },
  },
] as const;

type Args = Record<string, unknown>;

function str(args: Args, key: string): string {
  const v = args[key];
  if (typeof v !== 'string' || !v.trim()) throw new Error(`Argument "${key}" is required and must be a non-empty string.`);
  return v.trim();
}

function runTool(name: string, args: Args): unknown {
  switch (name) {
    case 'afo_read_workbook': {
      const sections = Array.isArray(args.sections) ? (args.sections as string[]) : undefined;
      const unknown = sections?.filter((s) => !(SECTIONS as readonly string[]).includes(s));
      if (unknown?.length) throw new Error(`Unknown section(s): ${unknown.join(', ')}. Valid: ${SECTIONS.join(', ')}.`);
      return readWorkbook(str(args, 'path'), {
        sections: sections as Section[] | undefined,
        macroSource: args.macro_source === true,
        maxFormulasPerSheet: typeof args.max_formulas_per_sheet === 'number' ? args.max_formulas_per_sheet : undefined,
      });
    }
    case 'afo_scan_workbooks':
      return scanWorkbooks(str(args, 'folder'), {
        recursive: args.recursive === true,
        limit: typeof args.limit === 'number' ? args.limit : undefined,
        detail: args.detail === 'full' || args.detail === 'compact' ? args.detail : undefined,
      });
    case 'afo_where_used':
      return whereUsed(str(args, 'folder'), str(args, 'name'), { recursive: args.recursive === true });
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

export function createServer(): Server {
  const server = new Server({ name: 'afo-mcp', version }, { capabilities: { tools: {} }, instructions: INSTRUCTIONS });
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS.map((t) => ({ ...t })) }));
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
      const result = runTool(request.params.name, (request.params.arguments ?? {}) as Args);
      return { content: [{ type: 'text', text: JSON.stringify(result) }] };
    } catch (err) {
      return { content: [{ type: 'text', text: `Error: ${(err as Error).message}` }], isError: true };
    }
  });
  return server;
}

export async function startStdio(): Promise<void> {
  await createServer().connect(new StdioServerTransport());
  process.stderr.write(`afo-mcp ${version} started\n`);
}
