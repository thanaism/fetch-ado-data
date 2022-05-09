import axios from 'axios';
import { EnvironmentVariables } from '../environment/env';
import { destinationProfile, ProjectProfile, sourceProfile } from '../environment/profile';
import { User } from '../users/user';
import { WorkItemJsonPatch } from './workItemJsonPatch';

export class SourceWorkItem {
  private static profile: ProjectProfile = sourceProfile;
  private static project: string = EnvironmentVariables.instance.SourceProject;

  public static async ids(): Promise<number[]> {
    const queryAllWorkItems =
      `Select [System.Id] From WorkItems ` + `Where [System.AreaPath] Under '${this.project}'`;
    const response = await WorkItemAPI.queryByWiql(this.profile, queryAllWorkItems);
    return response.data.workItems.map((e) => e.id);
  }

  public static async list(ids: number[]): Promise<WorkItem[]> {
    const limit = 200;
    const idChunks = chunk(ids, limit);
    const data: WorkItem[] = [];

    for (const chunk of idChunks) {
      const response = await WorkItemAPI.list(this.profile, chunk);
      data.push(...response.data.value);
    }

    return data;
  }
}

export class DestinationWorkItem {
  private static profile: ProjectProfile = destinationProfile;
  private static project: string = EnvironmentVariables.instance.DestinationProject;

  public static async ids(): Promise<number[]> {
    const queryAllWorkItems =
      `Select [System.Id] From WorkItems ` + `Where [System.AreaPath] Under '${this.project}'`;
    const response = await WorkItemAPI.queryByWiql(this.profile, queryAllWorkItems);
    return response.data.workItems.map((e) => e.id);
  }

  public static async list(ids: number[]): Promise<WorkItem[]> {
    const limit = 200;
    const idChunks = chunk(ids, limit);
    const data: WorkItem[] = [];

    for (const chunk of idChunks) {
      const response = await WorkItemAPI.list(this.profile, chunk);
      data.push(...response.data.value);
    }

    return data;
  }

  public static async get(id: string | number) {
    const response = await WorkItemAPI.get(this.profile, String(id));
    return response.data;
  }

  public static async delete(id: string | number) {
    const response = await WorkItemAPI.delete(this.profile, String(id), false);
    return response.data;
  }

  public static async deleteParmanently(id: string | number) {
    const response = await WorkItemAPI.delete(this.profile, String(id), true);
    return response.data;
  }

  public static async create(type: string, body: WorkItemJsonPatch[]) {
    console.info(`DestinationWorkItem.create: ${type}`);
    const response = await WorkItemAPI.create(this.profile, type, body);
    return response.data;
  }

  public static async update(id: string | number, body: WorkItemJsonPatch[]) {
    const response = await WorkItemAPI.update(this.profile, String(id), body);
    return response.data;
  }
}

class WorkItemAPI {
  public static queryByWiql(profile: ProjectProfile, query: string) {
    const { organization, project, team, username, password } = profile;
    const url = `https://dev.azure.com/${organization}/${project}/${team}/_apis/wit/wiql`;
    type QueryByWiqlResponse = { workItems: { id: number; url: string }[] };
    return axios.post<QueryByWiqlResponse>(
      url,
      { query: query },
      {
        auth: { username: username, password: password },
        params: { 'api-version': '7.1-preview.2' },
      },
    );
  }

  public static get(profile: ProjectProfile, id: string) {
    const { organization, project, username, password } = profile;
    const url = `https://dev.azure.com/${organization}/${project}/_apis/wit/workitems/${id}`;
    return axios.get<WorkItem>(url, {
      auth: { username: username, password: password },
      params: { 'api-version': '6.0', $expand: 'all' },
    });
  }

  public static list(profile: ProjectProfile, ids: number[]) {
    const { organization, project, username, password } = profile;
    const url = `https://dev.azure.com/${organization}/${project}/_apis/wit/workitems`;
    const commaSeparatedIds = ids.map((i) => i.toString()).join(',');
    return axios.get<WorkItemListResponse>(url, {
      auth: { username: username, password: password },
      params: { ids: commaSeparatedIds, 'api-version': '7.1-preview.3', $expand: 'all' },
    });
  }

  public static create(profile: ProjectProfile, type: string, body: WorkItemJsonPatch[]) {
    const { organization, project, username, password } = profile;
    const url = `https://dev.azure.com/${organization}/${project}/_apis/wit/workitems/$${type}`;
    return axios.post<WorkItem>(url, body, {
      auth: { username: username, password: password },
      params: { 'api-version': '6.0' },
      headers: { 'Content-Type': 'application/json-patch+json' },
    });
  }

  public static update(profile: ProjectProfile, id: string, body: WorkItemJsonPatch[]) {
    const { organization, project, username, password } = profile;
    const url = `https://dev.azure.com/${organization}/${project}/_apis/wit/workitems/${id}`;
    return axios.patch<WorkItem>(url, body, {
      auth: { username: username, password: password },
      params: { 'api-version': '6.0' },
      headers: { 'Content-Type': 'application/json-patch+json' },
    });
  }

  public static delete(profile: ProjectProfile, id: string, destroy: boolean) {
    const { organization, project, username, password } = profile;
    const url = `https://dev.azure.com/${organization}/${project}/_apis/wit/workitems/${id}`;
    return axios.delete<WorkItemDeleteResponse>(url, {
      auth: { username: username, password: password },
      params: { 'api-version': '6.0', destroy: destroy },
    });
  }
}
const chunk = <T>(arr: T[], size: number): T[][] =>
  arr.flatMap((e, i) => (i % size ? [] : [arr.slice(i, i + size)]));

type WorkItemDeleteResponse = {
  id: number;
  type: string;
  name: string;
  project: string;
  deletedDate: string;
  deletedBy: string;
  code: number;
  url: string;
  resource: WorkItem;
};

type WorkItemListResponse = {
  // count: number;
  value: WorkItem[];
};

export type WorkItem = {
  id: number;
  rev: number;
  fields: WorkItemFields;
  relations: WorkItemRelation[];
  _links: WorkItemLink;
  url: string;
};

type WorkItemRelationType = 'System.LinkTypes.Hierarchy-Forward' | 'AttachedFile';

export type WorkItemRelation = {
  rel: WorkItemRelationType;
  url: string;
  attributes: {
    name: string;
    isLocked?: false;
    authorizedDate?: string;
    id?: number;
    resourceCreatedDate?: string;
    resourceModifiedDate?: string;
    revisedDate?: string;
    resourceSize?: number;
  };
};

type WorkItemLinkKey =
  | 'self'
  | 'workItemUpdates'
  | 'workItemRevisions'
  | 'workItemComments'
  | 'html'
  | 'workItemType'
  | 'fields';

type WorkItemLink = {
  [key in WorkItemLinkKey]: {
    href: string;
  };
};

type WorkItemFields = {
  'System.Id': number;
  'System.AreaPath': string;
  'System.TeamProject': string;
  'System.NodeName': string;
  'System.IterationId': number;
  'System.IterationPath': string;
  'System.WorkItemType': string;
  'System.State': string;
  'System.Reason': string;
  'System.AssignedTo': User;
  'System.CreatedDate': string;
  'System.CreatedBy': User;
  'System.ChangedDate': string;
  'System.ChangedBy': User;
  'System.CommentCount': number;
  'System.Title': string;
  'System.BoardColumn': string;
  'System.BoardColumnDone': boolean;
  'System.Description': string;
};
