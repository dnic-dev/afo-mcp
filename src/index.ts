import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { createRequire } from 'node:module';

import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

import { type BuildSpec, type MinimalLayout, buildWorkbook } from './build/build.js';
import { type UpdateOperations, updateWorkbook } from './build/update.js';
import { SECTIONS, type Section, readWorkbook } from './workbook/read.js';
import { scanWorkbooks, whereUsed } from './workbook/scan.js';

const { version } = createRequire(import.meta.url)('../package.json') as { version: string };

const INSTRUCTIONS = `SAP Analysis for Office (AfO) workbooks, read and built offline: no Excel, no SAP logon.
The tools unzip .xlsx/.xlsm files and parse or write the AfO definition Analysis for Office stores in them.

- afo_read_workbook: everything about one workbook.
- afo_scan_workbooks: inventory of a folder: profiles, which query is used where, recurring layouts, outliers.
- afo_where_used: every place a query, planning function, variable or characteristic occurs across a folder.
- afo_build_workbook: write a new AfO workbook from a spec (data sources, crosstabs, planning functions,
  macros) onto a layout template or a generated minimal layout.
- afo_update_workbook: change an existing AfO workbook (add planning functions, retrofit the agent bridge,
  add VBA modules), keeping everything else.

Building: this server has no BW access. Take the facts a spec needs from the BW system first
(bw-modeling-mcp): query variables and InfoProvider (bw_get_query), the characteristics of the
aggregation level (bw_get_aggregation_level), planning functions and their variables, and the runtime
UID of a key figure structure for stored axes (bw_query_data, raw format: the entry with isStructure).
Stored prompt values use the internal format (0CALMONTH 202501); cells a planning function reads
hold the external format (01.2025). The result is a file only: AfO accepts it on the next open in
Excel with Analysis for Office, where the user logs on.

What a workbook stores and what it does not:
- Data sources carry their saved navigation state (filters, axes, hierarchies) only when AfO saved one;
  stateStored=false means the data source opens in the query's default view.
- Query variable values are stored only when the workbook setting StorePromptsInDocument is True.
  Planning function variables are always stored with the planning object.
- Structure members and some filter values appear as BW-internal IDs (e.g. 00O2TP…); resolving them
  needs the query definition from the BW system (bw-modeling-mcp: bw_get_query).
- Sheet references use the visible sheet name; zero-based AfO coordinates are converted to A1.`;

const PLANNING_FUNCTIONS_SCHEMA = {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  key: { type: 'string', description: 'Alias used by SAPExecutePlanningFunction, e.g. PF_1.' },
                  from: { type: 'string', description: 'Take this planning function or sequence over from the source workbook (its key there), with variables, filter and references; the other fields are then optional.' },
                  name: { type: 'string', description: 'Technical name of the planning function.' },
                  description: { type: 'string' },
                  dataProviderKey: { type: 'string', description: 'Data source the function runs on.' },
                  aggregationLevel: { type: 'string' },
                  characteristics: { type: 'array', items: { type: 'string' }, description: 'All characteristics of the aggregation level.' },
                  filter: { type: 'object', additionalProperties: { type: 'array', items: { type: 'string' } }, description: 'Fixed filter values per characteristic, internal keys.' },
                  variables: {
                    type: 'array',
                    items: { type: 'object', properties: { name: { type: 'string' }, value: { type: 'string' } }, required: ['name', 'value'] },
                    description: 'Stored variable values, internal format.',
                  },
                  variableCells: {
                    type: 'array',
                    items: { type: 'object', properties: { name: { type: 'string' }, cell: { type: 'string' } }, required: ['name', 'cell'] },
                    description: 'Variables read from a cell at execution time, e.g. Planning!$C$9. The cell holds the external format.',
                  },
                },
                required: ['key'],
              },
            } as const;

const AGENT_BRIDGE_SCHEMA = {
              type: 'object',
              description: 'Adds a command cell for agents such as Claude in Excel, whose cell writes Analysis for Office does not register: a bridge sheet (default AgentLog) with command cell B1 (defined name AfoAgentCommand), result B3, status B4, help, AfO messages and a step log, plus VBA that runs SET Sheet!A1=value, TRANSFER, PLAN key, PVAR key NAME=value, VAR DS NAME=value, FILTER DS DIM=value, REFRESH, SAVE, RESET, API SAPFunction arg1|arg2 (any AfO API function, SAPLogon excluded) and workbook commands. New AfO messages are logged after every step and errors fail it. Reading needs no command: the sheet explains the AfO read formulas and keeps live lists of data sources and planning objects. Needs an .xlsm output.',
              properties: {
                sheet: { type: 'string', description: 'Name of the bridge sheet (default AgentLog).' },
                hidden: { type: 'boolean', description: 'Hide the bridge sheet; agents find it through the defined names AfoAgentCommand, AfoAgentResult, AfoAgentStatus.' },
                commands: {
                  type: 'array',
                  description: 'Workbook commands, each running a macro, e.g. the macro behind a button.',
                  items: {
                    type: 'object',
                    properties: {
                      name: { type: 'string', description: 'Upper case command word, e.g. COPY_ACTUALS.' },
                      macro: { type: 'string', description: 'Module.Procedure among vbaModules.' },
                      description: { type: 'string' },
                    },
                    required: ['name', 'macro'],
                  },
                },
              },
            } as const;

const VBA_MODULES_SCHEMA = {
              type: 'array',
              description: 'Standard modules. Source as code, or as file (absolute, or relative to spec_path).',
              items: {
                type: 'object',
                properties: { name: { type: 'string' }, code: { type: 'string' }, file: { type: 'string' } },
                required: ['name'],
              },
            } as const;

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
  {
    name: 'afo_build_workbook',
    description:
      'Write a new Analysis for Office workbook without Excel: data sources on BW queries with stored prompts, axes and input readiness; crosstabs at given cells; planning functions bound to a data source with aggregation level, fixed filter values, stored variables and variables read from cells; VBA modules and button macros. The base is either a layout template (.xlsx/.xlsm with the sheets, texts, formats and the shapes used as buttons) or a generated minimal layout (sheets plus optional cell values and defined names, no buttons). The spec is checked first and every problem is reported at once; nothing is written then. After writing, the file is read back and broken links are returned as issues. The workbook opens in Excel with AfO, which asks for the logon; refresh and planning need that step.',
    inputSchema: {
      type: 'object',
      properties: {
        output: { type: 'string', description: 'Absolute path of the workbook to write, .xlsm (needed for VBA) or .xlsx. The folder must exist.' },
        spec: {
          type: 'object',
          description: 'The workbook spec. Alternatively spec_path.',
          properties: {
            source: { type: 'string', description: 'Existing AfO workbook (absolute, or relative to spec_path) to take data sources, planning objects, crosstab and formula item settings, the system connection (incl. SAProuter/SNC) and the workbook settings from, for entries with from. Use it to give an existing workbook a new layout without losing any AfO feature. Settings the spec sets win.' },
            formulaItems: {
              type: 'array',
              description: 'AfO formula items such as filter components. The builder writes the formula into the cell as well.',
              items: {
                type: 'object',
                properties: {
                  key: { type: 'string', description: 'e.g. Fdd1.' },
                  from: { type: 'string', description: 'Take the item over from the source workbook; its formula follows renamed data sources.' },
                  sheet: { type: 'string' },
                  cell: { type: 'string', description: 'A1.' },
                  formula: { type: 'string', description: 'e.g. =SAPSetFilterComponent("DS_1", "0CALMONTH", ""). Required unless from.' },
                },
                required: ['key', 'sheet', 'cell'],
              },
            },
            system: {
              type: 'object',
              description: 'BW connection as in the SAP Logon landscape (%APPDATA%\\SAP\\Common\\SAPUILandscape.xml, Service entry with this systemid). Not needed with a source workbook: its connection is used.',
              properties: {
                systemId: { type: 'string', description: 'SID, e.g. ABC.' },
                systemName: { type: 'string', description: 'The landscape entry name, e.g. "ABC - BW/4HANA".' },
                client: { type: 'string', description: 'e.g. 001.' },
                language: { type: 'string', description: 'Logon language, e.g. DE or EN.' },
                server: { type: 'string', description: 'Application server host, without port.' },
                systemNumber: { type: 'string', description: 'Instance number, two digits (port 32NN → NN).' },
                uuid: { type: 'string', description: 'The landscape entry uuid.' },
              },
              required: ['systemId', 'systemName', 'client', 'language'],
            },
            refreshOnOpen: { type: 'boolean', description: 'Refresh all data sources when the workbook opens.' },
            changeModeOnOpen: { type: 'boolean', description: 'Planning: open input-ready data sources in change mode.' },
            storePrompts: { type: 'boolean', description: 'Store prompt values in the workbook (StorePromptsInDocument). Set true when prompts are given.' },
            mergeVariables: { type: 'boolean', description: 'Share variables of the same name across data sources and planning objects (AfO default false).' },
            dataProviders: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  key: { type: 'string', description: 'Alias, e.g. DS_1.' },
                  from: { type: 'string', description: 'Take this data source over from the source workbook (its key there) with its whole stored state: filters incl. hierarchy nodes, display settings, design rules, exceptions, conditions, prompts. query and the other fields are then not needed.' },
                  query: { type: 'string', description: 'Technical name of the query.' },
                  infoProvider: { type: 'string', description: 'The query\'s InfoProvider (for input-ready queries the aggregation level).' },
                  inputReady: { type: 'boolean', description: 'Default true.' },
                  prompts: {
                    type: 'array',
                    description: 'Stored variable values, internal format. Each entry has exactly one of member, literal, low(+high).',
                    items: {
                      type: 'object',
                      properties: {
                        name: { type: 'string' },
                        member: { type: 'string', description: 'Single value of a characteristic variable, internal key.' },
                        literal: { type: 'string', description: 'Literal value, e.g. a version typed in.' },
                        low: { type: 'string', description: 'Interval start, internal key (0CALMONTH: 202501).' },
                        high: { type: 'string', description: 'Interval end.' },
                      },
                      required: ['name'],
                    },
                  },
                  axes: {
                    type: 'object',
                    description: 'Stored layout. Characteristics by technical name, a key figure structure by its runtime UID. Omitted axes keep the query default.',
                    properties: {
                      rows: { type: 'array', items: { type: 'string' } },
                      columns: { type: 'array', items: { type: 'string' } },
                      free: { type: 'array', items: { type: 'string' } },
                    },
                  },
                },
                required: ['key'],
              },
            },
            grids: {
              type: 'array',
              description: 'Crosstabs. AfO does not move crosstabs apart: leave room for the rows each one grows to.',
              items: {
                type: 'object',
                properties: {
                  key: { type: 'string', description: 'e.g. Crosstab1.' },
                  from: { type: 'string', description: 'Take the crosstab settings over from this crosstab of the source workbook; only the position is new. dataProviderKey follows the taken-over data source.' },
                  dataProviderKey: { type: 'string' },
                  sheet: { type: 'string', description: 'Visible sheet name.' },
                  cell: { type: 'string', description: 'Top-left cell, A1.' },
                },
                required: ['key', 'sheet', 'cell'],
              },
            },
            planningFunctions: PLANNING_FUNCTIONS_SCHEMA,
            shapeMacros: { type: 'object', additionalProperties: { type: 'string' }, description: 'Shape name in the template → Module.Procedure run on click.' },
            agentBridge: AGENT_BRIDGE_SCHEMA,
            documentModules: {
              type: 'object',
              additionalProperties: { type: 'string' },
              description: 'Code of document modules by visible sheet name or ThisWorkbook: event handlers such as Worksheet_Change or Workbook_SheetChange.',
            },
            vbaModules: VBA_MODULES_SCHEMA,
          },
          required: ['dataProviders', 'grids'],
        },
        spec_path: { type: 'string', description: 'Absolute path of a JSON file with the spec, instead of spec. Relative VBA module files resolve against its folder.' },
        template: { type: 'string', description: 'Absolute path of a plain layout workbook (not yet an AfO workbook). Every crosstab sheet must exist in it; button shapes are found by name.' },
        layout: {
          type: 'object',
          description: 'Generated minimal layout, used when there is no template. Sheets named by crosstabs and cell references are created automatically.',
          properties: {
            sheets: { type: 'array', items: { type: 'string' }, description: 'Sheets in tab order.' },
            cells: {
              type: 'array',
              description: 'Values to write, e.g. labels and the parameter cells planning functions read. Text is stored with text format.',
              items: {
                type: 'object',
                properties: {
                  sheet: { type: 'string' },
                  cell: { type: 'string', description: 'A1.' },
                  value: { type: ['string', 'number'] },
                  name: { type: 'string', description: 'Optional defined name for the cell.' },
                },
                required: ['sheet', 'cell', 'value'],
              },
            },
          },
        },
        overwrite: { type: 'boolean', description: 'Replace an existing output file (default false).' },
      },
      required: ['output'],
    },
  },
  {
    name: 'afo_update_workbook',
    description:
      'Change an existing Analysis for Office workbook without Excel instead of rebuilding it: add planning functions to its AfO definition (bound to a data source already in the workbook), retrofit the agent bridge so Claude in Excel can operate it, add VBA modules. Layout, formulas, data sources, crosstabs and existing macros stay as they are; existing Workbook_Open/Workbook_SheetCalculate handlers get one call added. Writes a new file (output), never the input; the workbook must be closed in Excel. Everything is checked first and every problem reported at once. Signed VBA projects are refused.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string', description: 'Absolute path of the existing AfO workbook.' },
        output: { type: 'string', description: 'Absolute path of the changed copy (.xlsm when macros are involved).' },
        planningFunctions: PLANNING_FUNCTIONS_SCHEMA,
        agentBridge: AGENT_BRIDGE_SCHEMA,
        vbaModules: VBA_MODULES_SCHEMA,
        base_dir: { type: 'string', description: 'Folder that relative VBA module files resolve against.' },
        overwrite: { type: 'boolean', description: 'Replace an existing output file (default false).' },
      },
      required: ['input', 'output'],
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
    case 'afo_build_workbook': {
      const hasSpec = args.spec !== undefined, hasPath = args.spec_path !== undefined;
      if (hasSpec === hasPath) throw new Error('Give either spec or spec_path.');
      let spec: BuildSpec;
      let baseDir: string | undefined;
      if (hasPath) {
        const path = resolve(str(args, 'spec_path'));
        try {
          spec = JSON.parse(readFileSync(path, 'utf8').replace(/^﻿/, '')) as BuildSpec;
        } catch (err) {
          throw new Error(`Cannot read the spec ${path}: ${(err as Error).message}`);
        }
        baseDir = dirname(path);
      } else {
        if (typeof args.spec !== 'object' || args.spec === null) throw new Error('spec must be an object.');
        spec = args.spec as BuildSpec;
      }
      return buildWorkbook(spec, {
        output: str(args, 'output'),
        template: typeof args.template === 'string' && args.template.trim() ? args.template.trim() : undefined,
        layout: typeof args.layout === 'object' && args.layout !== null ? (args.layout as MinimalLayout) : undefined,
        overwrite: args.overwrite === true,
        baseDir,
      });
    }
    case 'afo_update_workbook':
      return updateWorkbook(
        {
          planningFunctions: args.planningFunctions as UpdateOperations['planningFunctions'],
          agentBridge: args.agentBridge as UpdateOperations['agentBridge'],
          vbaModules: args.vbaModules as UpdateOperations['vbaModules'],
        },
        {
          input: str(args, 'input'),
          output: str(args, 'output'),
          overwrite: args.overwrite === true,
          baseDir: typeof args.base_dir === 'string' && args.base_dir.trim() ? args.base_dir.trim() : undefined,
        },
      );
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
