/*
TODO
- Copying attachments
- Associations between tasks (e.g., parent-child relationship between PBI and Task) -> probably done
- Copy date of Itrations
- Restore the original assignee, speaker, time of the statement, etc. (add information so that it can be tracked since full restoration is not possible)
- Citing other items in comments, queried only by number, so it is necessary to change it to the ID of the copy destination.
 */

const main = () => {
  duplicateClassificationPaths();
  duplicateAllProjectWorkItems();
};

// --------------------------------------------------------------------------------------------------------------------
// Copy UserProps to ScriptProps --------------------------------------------------------------------------------------
class KeyValueStore {
  userStore = PropertiesService.getUserProperties();
  scriptStore = PropertiesService.getScriptProperties();
  getProperty(key: string) {
    /*
     If the key consists of anything other than numbers,
     it is not stored in the script side because it is confidential information.
     */
    if (!key.match(/[0-9]+/)) return this.userStore.getProperty(key);
    const userRetVal = this.userStore.getProperty(key);
    if (userRetVal === null) {
      console.log(`key: ${key} is null, skipped.`);
      return userRetVal;
    } // If null, do not store.
    this.scriptStore.setProperty(key, userRetVal);
    return this.scriptStore.getProperty(key);
  }
  setProperty(key: string, value: string) {
    this.userStore.setProperty(key, value);
    this.scriptStore.setProperty(key, value);
  }
  deleteProperty(key: string) {
    this.userStore.deleteProperty(key);
    this.scriptStore.deleteProperty(key);
  }
}

// --------------------------------------------------------------------------------------------------------------------
// Global variables for requests --------------------------------------------------------------------------------------
const USER = new KeyValueStore();
const user_name = USER.getProperty('user_name');
const personal_access_token = USER.getProperty('personal_access_token');
const organization = USER.getProperty('organization');
const project = USER.getProperty('project');
const team = USER.getProperty('team');
const token = Utilities.base64Encode(
  `${user_name}:${personal_access_token}`,
  Utilities.Charset.UTF_8,
);
const personal_access_token_copyto = USER.getProperty('personal_access_token_copyto');
const organization_copyto = USER.getProperty('organization_copyto');
const project_copyto = USER.getProperty('project_copyto');
const team_copyto = USER.getProperty('team_copyto');
const token_copyto = Utilities.base64Encode(
  `${user_name}:${personal_access_token_copyto}`,
  Utilities.Charset.UTF_8,
);

// --------------------------------------------------------------------------------------------------------------------
// JSON Response Type Difinition --------------------------------------------------------------------------------------
type MailAddress = string;
type URL = string;
type GUID = string;
type IterationPath = string;
type WorkItemType = 'Product Backlog Item' | 'Task' | 'Epic' | 'Feature';
type State = 'Active' | 'In Progress' | 'Closed';

type User = {
  displayName: string;
  url: URL;
  _links: {
    avatar: {
      href: URL;
    };
  };
  id: GUID;
  uniqueName: MailAddress;
  imageUrl: URL;
  descriptor: string;
};

type DateLike = string;
interface WorkItemFields {
  'System.Id': number;
  'System.AreaId': number;
  'System.AreaPath': string;
  'System.TeamProject': string;
  'System.NodeName': string;
  'System.AreaLevel1': string;
  'System.Rev': number;
  'System.AuthorizedDate': DateLike;
  'System.RevisedDate': DateLike;
  'System.IterationId': number;
  'System.IterationPath': IterationPath;
  'System.IterationLevel1': string;
  'System.IterationLevel2': string;
  'System.WorkItemType': WorkItemType;
  'System.State': State;
  'System.Reason': string;
  'System.AssignedTo': User;
  'System.CreatedDate': DateLike;
  'System.CreatedBy': User;
  'System.ChangedDate': DateLike;
  'System.ChangedBy': User;
  'System.AuthorizedAs': User;
  'System.PersonId': number;
  'System.Watermark': number;
  'System.CommentCount': number;
  'System.Title': string;
  'System.BoardColumn': string;
  'System.BoardColumnDone': boolean;
  'Microsoft.VSTS.Common.StateChangeDate': DateLike;
  'Microsoft.VSTS.Common.ClosedDate': DateLike;
  'Microsoft.VSTS.Common.ClosedBy': User;
  'Microsoft.VSTS.Common.Priority': number;
  'Microsoft.VSTS.Common.ValueArea': string;
  'Microsoft.VSTS.Common.BacklogPriority': number;
  // 'WEF_8F5044CDB6F649D5BD09C4B52C8D802F_System.ExtensionMarker': boolean;
  // 'WEF_8F5044CDB6F649D5BD09C4B52C8D802F_Kanban.Column': string;
  // 'WEF_8F5044CDB6F649D5BD09C4B52C8D802F_Kanban.Column.Done': boolean;
  'System.Description': string;
}

type WorkItemRelationType = 'System.LinkTypes.Hierarchy-Forward' | 'AttachedFile';

type WorkItemRelation = {
  rel: WorkItemRelationType;
  url: URL;
  attributes: {
    name: string;
    isLocked?: false;
    authorizedDate?: DateLike;
    id?: number;
    resourceCreatedDate?: DateLike;
    resourceModifiedDate?: DateLike;
    revisedDate?: DateLike;
    resourceSize?: number;
  };
};

interface WorkItemResponse {
  id: number;
  rev: number;
  fields: WorkItemFields;
  relations: WorkItemRelation[];
  _links: WorkItemLink;
  url: URL;
}

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
    href: URL;
  };
};

interface WorkItem {
  id: string;
  areaPath: string;
  state: string;
  iterationPath: string;
  workItemType: string;
  assignedTo: string;
  createdDate: DateLike;
  changedDate: DateLike;
  createdBy: string;
  title: string;
  description: string;
  relations: Relation[] | undefined;
  comments: commentReduced[];
}

interface comment {
  mentions: Array<Object>;
  workItemId: number;
  id: number;
  version: number;
  text: string;
  createdBy: {
    displayName: string;
  };
  createdDate: string;
  modifiedBy: object;
  modifiedDate: string;
  url: string;
}

// --------------------------------------------------------------------------------------------------------------------
// Functions for sending API requests ---------------------------------------------------------------------------------

const deleteToCopied = (
  requestUrl: string,
  payload: string | undefined = undefined,
  contentType: string = 'application/json',
) => {
  return sendRequest('copied', 'delete', requestUrl, payload, contentType);
};
const patchToCopied = (
  requestUrl: string,
  payload: string | undefined = undefined,
  contentType: string = 'application/json-patch+json',
) => {
  return sendRequest('copied', 'patch', requestUrl, payload, contentType);
};
const postToCopied = (
  requestUrl: string,
  payload: string | undefined = undefined,
  contentType: string = 'application/json',
) => {
  return sendRequest('copied', 'post', requestUrl, payload, contentType);
};
const postToOriginal = (
  requestUrl: string,
  payload: string | undefined = undefined,
  contentType: string = 'application/json',
) => {
  return sendRequest('original', 'post', requestUrl, payload, contentType);
};
const getToCopied = (
  requestUrl: string,
  payload: string | undefined = undefined,
  contentType: string = 'application/json',
) => {
  return sendRequest('copied', 'get', requestUrl, payload, contentType);
};
const getToOriginal = (
  requestUrl: string,
  payload: string | undefined = undefined,
  contentType: string = 'application/json',
) => {
  return sendRequest('original', 'get', requestUrl, payload, contentType);
};
const sendRequest = (
  destination: 'copied' | 'original',
  method: 'post' | 'get' | 'patch' | 'delete',
  requestUrl: string,
  payload: string | undefined,
  contentType: string,
) => {
  const isDestinationOriginal = destination == 'original';
  const requestOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    muteHttpExceptions: true,
    method: method,
    headers: { Authorization: `Basic ${isDestinationOriginal ? token : token_copyto}` },
    contentType: contentType,
    payload: payload,
  };
  // console.log(method + ' ' + requestUrl);
  const response = UrlFetchApp.fetch(requestUrl, requestOptions);
  if (method != 'delete') {
    const responseText = response.getContentText();
    const responseJson = JSON.parse(responseText);
    return responseJson;
  }
};

// --------------------------------------------------------------------------------------------------------------------
// Duplicate Areas and Iterations -------------------------------------------------------------------------------------
type ClassificationPath = {
  id: number;
  identifier: string;
  name: string;
  structureType: 'area' | 'iteration';
  hasChildren: boolean;
  children?: ClassificationPath[];
  path: string;
  url: string;
};

type RootClassificationPath = {
  count: number;
  value: ClassificationPath[];
};

const duplicateClassificationPaths = () => {
  const requestUrl = (destination: 'copied' | 'original') => {
    const org = destination == 'copied' ? organization_copyto : organization;
    const prj = destination == 'copied' ? project_copyto : project;
    return `https://dev.azure.com/${org}/${prj}/_apis/wit/classificationnodes?api-version=5.0&$depth=100`;
  };
  const originalResponse: RootClassificationPath = getToOriginal(requestUrl('original'));
  const copiedResponse: RootClassificationPath = getToCopied(requestUrl('copied'));

  // Duplicate under Copied in the destination project
  const originalRootArea = originalResponse.value[0];
  const originalRootIteration = originalResponse.value[1];
  const copiedRootArea = copiedResponse.value[0].children?.find(
    element => element.name == 'Copied',
  );
  const copiedRootIteration = copiedResponse.value[1].children?.find(
    element => element.name == 'Copied',
  );
  const getSanitizedRootPath = (rootPath: string) =>
    ['Copied'].concat(rootPath.split('\\').slice(4)).join('/');
  const iterationIds: string[] = [];
  const rec = (originalPath: ClassificationPath, copiedPath: ClassificationPath) => {
    if (!originalPath.hasChildren) return;
    originalPath.children?.forEach(originalChild => {
      let counterpart = copiedPath.children?.find(
        copiedChild => copiedChild.name == originalChild.name,
      );
      if (counterpart == undefined)
        counterpart = createClassificationPath(
          originalChild.structureType == 'area' ? 'Areas' : 'Iterations',
          getSanitizedRootPath(originalChild.path),
          originalChild.name,
        );
      iterationIds.push(counterpart.identifier);
      rec(originalChild, counterpart);
    });
  };
  if (copiedRootArea != undefined) rec(originalRootArea, copiedRootArea);
  if (copiedRootIteration != undefined) rec(originalRootIteration, copiedRootIteration);
  addIterationsToTeam(iterationIds);
};

const addIterationsToTeam = (iterationIds: string[]) => {
  const requestUrl = `https://dev.azure.com/${organization_copyto}/${project_copyto}/${team_copyto}/_apis/work/teamsettings/iterations?api-version=6.0`;
  iterationIds.forEach((id: string) => {
    const payload = JSON.stringify({ id: id });
    postToCopied(requestUrl, payload);
  });
};

type StructureGroup = 'Areas' | 'Iterations';
const createClassificationPath = (
  structureGroup: StructureGroup,
  newClassificationPathRoot: string,
  newClassificationPathBaseName: string,
) => {
  // Request Parameter Creation
  const requestUrl = encodeURI(
    `https://dev.azure.com/${organization_copyto}/${project_copyto}/_apis/wit/` +
      `classificationnodes/${structureGroup}/${newClassificationPathRoot}?api-version=5.0`,
  );
  const requestPayload = JSON.stringify({ name: newClassificationPathBaseName });

  // Send Request
  const response: ClassificationPath = postToCopied(requestUrl, requestPayload);
  return response;
};

// --------------------------------------------------------------------------------------------------------------------
// Creation of WorkItems ----------------------------------------------------------------------------------------------
const propertyExists = (key: string) => !!USER.getProperty(key);

const getNewClassificationPath = (originalClassificationPath: string) => {
  const originalClassificationPathArray = originalClassificationPath.split('\\').splice(1);
  return [project_copyto, 'Copied'].concat(originalClassificationPathArray).join('\\');
};

interface workItemJsonPatch {
  op: 'add';
  path: string;
  from: string | null;
  value: string | object;
}

interface Relation {
  rel: string;
  url: string;
  attributes: object;
}

const generateJsonPatch = (originalWorkItem: WorkItem) => {
  const newAreaPath = getNewClassificationPath(originalWorkItem.areaPath);
  const newIterationPath = getNewClassificationPath(originalWorkItem.iterationPath);
  const requestPayloadObject: workItemJsonPatch[] = [
    {
      op: 'add',
      path: '/fields/System.AreaPath',
      from: null,
      value: newAreaPath,
    },
    {
      op: 'add',
      path: '/fields/System.State',
      from: null,
      value: encodeURI(originalWorkItem.state),
    },
    {
      op: 'add',
      path: '/fields/System.IterationPath',
      from: null,
      value: newIterationPath,
    },
    {
      op: 'add',
      path: '/fields/System.Title',
      from: null,
      value: originalWorkItem.title,
    },
    {
      op: 'add',
      path: '/fields/System.Description',
      from: null,
      value: originalWorkItem.description ?? '',
    },
  ];
  if (originalWorkItem.relations != undefined) {
    const baseUrl = `https://dev.azure.com/${organization_copyto}/_apis/wit/workItems/`;
    const relations: workItemJsonPatch[] = originalWorkItem.relations.map((item: Relation) => {
      const counterpartId = USER.getProperty(item.url.split('/').splice(-1)[0]);
      return {
        op: 'add',
        path: '/relations/-',
        from: null,
        value: {
          rel: item.rel,
          url: baseUrl + counterpartId,
        },
      };
    });
    const requestPayload = JSON.stringify(requestPayloadObject.concat(relations));
    return requestPayload;
  } else {
    const requestPayload = JSON.stringify(requestPayloadObject);
    return requestPayload;
  }
};

interface CommentResponse {
  workItemId: number;
  id: number;
  version: number;
  text: string;
  createdBy: object;
  createdDate: string;
  modifiedBy: object;
  modifiedDate: string;
  url: string;
}

const updateComments = (originalWorkItem: WorkItem) => {
  const workItemId = USER.getProperty(originalWorkItem.id);
  originalWorkItem.comments.forEach((comment: commentReduced) => {
    const originalCommentId = comment.id.toString();
    if (!originalCommentId.match(/[0-9]+/)) {
      console.log(originalCommentId);
      throw Error;
    }
    if (propertyExists(originalCommentId)) {
      const copiedCommentId = USER.getProperty(originalCommentId);
      if (copiedCommentId == null) throw Error;
      const requestUrl = `https://dev.azure.com/${organization_copyto}/${project_copyto}/_apis/wit/workItems/${workItemId}/comments/${copiedCommentId}?api-version=6.0-preview.3`;
      deleteToCopied(requestUrl);
      console.log('Deleted Comment Id: ' + copiedCommentId);
    }
    const requestUrl = `https://dev.azure.com/${organization_copyto}/${project_copyto}/_apis/wit/workItems/${workItemId}/comments?api-version=7.1-preview.3`;
    const payload = JSON.stringify({ text: comment.text });
    const response: CommentResponse = postToCopied(requestUrl, payload);
    const copiedCommentId = response.id.toString();
    USER.setProperty(originalCommentId, copiedCommentId);
    console.log('Updated Comment ID: ' + copiedCommentId);
  });
};

const updateWorkItem: (arg0: string, arg1: WorkItem) => void = (copiedItemId, originalItem) => {
  const requestUrl = `https://dev.azure.com/${organization_copyto}/${project_copyto}/_apis/wit/workitems/${copiedItemId}?api-version=6.0`;
  const requestPayload = generateJsonPatch(originalItem);
  patchToCopied(requestUrl, requestPayload);
  console.log('Updated WorkItem ID: ' + copiedItemId);
  updateComments(originalItem);
};

const createWorkItem: (originalWorkItem: WorkItem) => string = originalWorkItem => {
  const type = encodeURI(originalWorkItem.workItemType);
  const requestUrl = `https://dev.azure.com/${organization_copyto}/${project_copyto}/_apis/wit/workitems/$${type}?api-version=6.0`;
  const requestPayload = generateJsonPatch(originalWorkItem);
  const response: WorkItemResponse = postToCopied(
    requestUrl,
    requestPayload,
    'application/json-patch+json',
  );
  const id = (response.id * 1).toString();
  console.log(`Newly Created Item Id: ${id}`);
  updateComments(originalWorkItem);
  return id;
};

interface commentReduced {
  id: number;
  createdBy: string;
  createdDate: string;
  text: string;
}

interface itemQueriedByWiql {
  id: number;
  url: string;
}

const sanitizeComment = (commentText: string) => commentText;

const duplicateAllProjectWorkItems = () => {
  // Query for WorkItem IDs that exist in the source project
  const queriedIds = (() => {
    const requestPayload = JSON.stringify({ query: 'Select [System.Id] From WorkItems' });
    const requestUrl = `https://dev.azure.com/${organization}/${project}/${team}/_apis/wit/wiql?api-version=7.1-preview.2`;
    return postToOriginal(requestUrl, requestPayload);
  })();

  const workItems: itemQueriedByWiql[] = queriedIds.workItems;
  const ids = workItems.map((item: itemQueriedByWiql) => item.id.toString());
  console.log(`Original project's workitem count: ${ids.length}`);

  // Copy each WorkItem
  ids.forEach((id: string) => {
    console.log('Queried WorkItem ID: ' + id);
    const workItem = generateWorkItemObjectFromId(id);
    duplicateSingleWorkItem(workItem);
  });
};

const generateWorkItemObjectFromId = (id: string) => {
  // Query WorkItems
  const queriedWorkItem: WorkItemResponse = (() => {
    const requestUrl = `https://dev.azure.com/${organization}/${project}/_apis/wit/workitems/${id}?api-version=7.1-preview.3&$expand=all`;
    return getToOriginal(requestUrl);
  })();

  // Query Comments
  const queried_comments = (() => {
    const requestUrl = `https://dev.azure.com/${organization}/${project}/_apis/wit/workItems/${id}/comments?api-version=7.1-preview.3`;
    return getToOriginal(requestUrl);
  })();

  // Reorder in ascending order since it is in descending order
  const comments: comment[] = queried_comments.comments.reverse();

  // Extract only necessary items from comment information
  const commentsReduced: commentReduced[] = comments.map((comment: comment) => ({
    id: comment.id,
    createdBy: comment.createdBy.displayName,
    createdDate: comment.createdDate,
    text: sanitizeComment(comment.text),
  }));

  // Description of original WorkItem (including comments)
  const fields: WorkItemFields = queriedWorkItem.fields;
  const workItem: WorkItem = {
    id: id,
    areaPath: fields['System.AreaPath'],
    state: fields['System.State'],
    iterationPath: fields['System.IterationPath'],
    workItemType: fields['System.WorkItemType'],
    assignedTo:
      'System.AssignedTo' in fields ? fields['System.AssignedTo'].displayName : 'Unassigned',
    createdDate: fields['System.CreatedDate'],
    changedDate: fields['System.ChangedDate'],
    createdBy: fields['System.CreatedBy'].displayName,
    title: fields['System.Title'],
    description: fields['System.Description'],
    relations: queriedWorkItem.relations ?? undefined,
    comments: commentsReduced,
  };
  return workItem;
};

const duplicateSingleWorkItem: (originalItem: WorkItem) => void = originalItem => {
  if (!propertyExists(originalItem.id)) {
    const copiedItemId = createWorkItem(originalItem);
    const sanitizedId = copiedItemId.replace(/([0-9]+)\.0/, '$1');
    USER.setProperty(originalItem.id, sanitizedId);
  } else {
    const copiedItemId = USER.getProperty(originalItem.id);
    if (copiedItemId == null) throw Error;
    if (copiedItemId == 'NaN') {
      USER.deleteProperty(originalItem.id);
      duplicateSingleWorkItem(originalItem);
      return;
    }
    const sanitizedId = copiedItemId.replace(/([0-9]+)\.0/, '$1');
    USER.setProperty(originalItem.id, sanitizedId);
    updateWorkItem(sanitizedId, originalItem);
  }
};
