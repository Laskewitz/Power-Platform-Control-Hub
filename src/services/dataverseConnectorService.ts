/**
 * Cross-environment Dataverse queries via the Microsoft Dataverse connector
 * (shared_commondataserviceforapps). Using the `GetItemWithOrganization` and
 * `ListRecordsWithOrganization` operations passes an `organization` header that
 * routes the call to the correct environment's Dataverse — the connector handles
 * authentication automatically without any internal SDK hacks.
 */
import { getClient } from '@microsoft/power-apps/data';
import type { IOperationResult } from '@microsoft/power-apps/data';
import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo.ts';
import type { Bots } from '../generated/models/BotsModel.ts';

const client = getClient(dataSourcesInfo);
const DATA_SOURCE = 'commondataserviceforapps';

/** A bot component record from Dataverse (topics, skills, knowledge sources, etc.) */
export interface BotComponent {
  botcomponentid?: string;
  name?: string;
  /** See COMPONENT_TYPE_LABELS for human-readable labels */
  componenttype?: number;
  statecode?: number;
  statuscode?: number;
  category?: string;
  description?: string;
}

export interface DataverseSystemUser {
  systemuserid?: string;
  azureactivedirectoryobjectid?: string;
  fullname?: string;
  internalemailaddress?: string;
  domainname?: string;
}

/** Human-readable labels for botcomponent ComponentType picklist values */
export const COMPONENT_TYPE_LABELS: Record<number, string> = {
  0: 'Topic',
  1: 'Skill',
  2: 'Bot Variable',
  3: 'Bot Entity',
  4: 'Dialog',
  5: 'Trigger',
  6: 'Language Understanding',
  7: 'Language Generation',
  8: 'Dialog Schema',
  9: 'Topic (V2)',
  10: 'Bot Translations',
  11: 'Bot Entity (V2)',
  12: 'Bot Variable (V2)',
  13: 'Skill (V2)',
  14: 'File Attachment',
  15: 'Custom GPT',
  16: 'Knowledge Source',
  17: 'External Trigger',
  18: 'Copilot Settings',
  19: 'Test Case',
};

/**
 * Fetch a single bot record from any Dataverse environment by its ID.
 * @param instanceUrl  The Dataverse org URL, e.g. "https://orgXXX.crm.dynamics.com"
 * @param botId        The bot's GUID (the `name` field on the Inventory resource)
 * @param $select      Comma-separated list of OData fields to retrieve
 */
export async function getDataverseBotById(
  instanceUrl: string,
  botId: string,
  $select: string,
): Promise<IOperationResult<Bots>> {
  const org = instanceUrl.endsWith('/') ? instanceUrl.slice(0, -1) : instanceUrl;
  return client.executeAsync<Record<string, unknown>, Bots>({
    connectorOperation: {
      tableName: DATA_SOURCE,
      operationName: 'GetItemWithOrganization',
      parameters: {
        organization: org,
        entityName: 'bots',
        recordId: botId,
        prefer: 'odata.include-annotations=*',
        accept: 'application/json',
        '$select': $select,
      },
    },
  });
}

/**
 * List bot records from any Dataverse environment using an OData filter.
 * Returns the `value` array from the OData response.
 * @param instanceUrl  The Dataverse org URL
 * @param $filter      OData filter expression, e.g. "name eq 'abc123'"
 * @param $select      Comma-separated list of OData fields to retrieve
 */
export async function listDataverseBots(
  instanceUrl: string,
  $filter: string,
  $select: string,
): Promise<IOperationResult<{ value?: Bots[] }>> {
  const org = instanceUrl.endsWith('/') ? instanceUrl.slice(0, -1) : instanceUrl;
  return client.executeAsync<Record<string, unknown>, { value?: Bots[] }>({
    connectorOperation: {
      tableName: DATA_SOURCE,
      operationName: 'ListRecordsWithOrganization',
      parameters: {
        organization: org,
        entityName: 'bots',
        '$filter': $filter,
        '$select': $select,
        '$top': 1,
        accept: 'application/json',
      },
    },
  });
}

/**
 * List bot component records for a specific bot from any Dataverse environment.
 * Fetches only top-level components (no parent component) to avoid deep nesting noise.
 * @param instanceUrl  The Dataverse org URL
 * @param botId        The bot's primary key GUID (bot.botid)
 * @param $select      Comma-separated list of OData fields to retrieve
 */
export async function listBotComponents(
  instanceUrl: string,
  botId: string,
  $select: string,
): Promise<IOperationResult<{ value?: BotComponent[] }>> {
  const org = instanceUrl.endsWith('/') ? instanceUrl.slice(0, -1) : instanceUrl;
  return client.executeAsync<Record<string, unknown>, { value?: BotComponent[] }>({
    connectorOperation: {
      tableName: DATA_SOURCE,
      operationName: 'ListRecordsWithOrganization',
      parameters: {
        organization: org,
        entityName: 'botcomponents',
        '$filter': `_parentbotid_value eq ${botId} and _parentbotcomponentid_value eq null`,
        '$select': $select,
        '$top': 500,
        '$orderby': 'componenttype asc,name asc',
        accept: 'application/json',
      },
    },
  });
}

/** Resolve Dataverse system-user IDs or Entra object IDs within one environment. */
export async function listDataverseSystemUsers(
  instanceUrl: string,
  ids: string[],
): Promise<IOperationResult<{ value?: DataverseSystemUser[] }>> {
  const org = instanceUrl.endsWith('/') ? instanceUrl.slice(0, -1) : instanceUrl;
  const filter = ids
    .flatMap((id) => [
      `systemuserid eq ${id}`,
      `azureactivedirectoryobjectid eq ${id}`,
    ])
    .join(' or ');

  return client.executeAsync<Record<string, unknown>, { value?: DataverseSystemUser[] }>({
    connectorOperation: {
      tableName: DATA_SOURCE,
      operationName: 'ListRecordsWithOrganization',
      parameters: {
        organization: org,
        entityName: 'systemusers',
        '$filter': filter,
        '$select': 'systemuserid,azureactivedirectoryobjectid,fullname,internalemailaddress,domainname',
        '$top': ids.length * 2,
        accept: 'application/json',
      },
    },
  });
}

/** Load a compact environment user directory for client-side identity matching. */
export async function listAllDataverseSystemUsers(
  instanceUrl: string,
): Promise<IOperationResult<{ value?: DataverseSystemUser[] }>> {
  const org = instanceUrl.endsWith('/') ? instanceUrl.slice(0, -1) : instanceUrl;
  return client.executeAsync<Record<string, unknown>, { value?: DataverseSystemUser[] }>({
    connectorOperation: {
      tableName: DATA_SOURCE,
      operationName: 'ListRecordsWithOrganization',
      parameters: {
        organization: org,
        entityName: 'systemusers',
        '$select': 'systemuserid,azureactivedirectoryobjectid,fullname,internalemailaddress,domainname',
        '$top': 5000,
        accept: 'application/json',
      },
    },
  });
}
