import { attachmentDownloadAndUpload } from '../attachments/attachmentDuplicator';
import { addMetadata, convertLinks } from '../converter';
import { EnvironmentVariables } from '../environment/env';
import { Memo } from '../keyValueStore/memo';
import { destinationProfile } from '../environment/profile';
import { WorkItem, WorkItemRelation } from './workItemAPI';

export type WorkItemJsonPatch = {
  op: string;
  path: string;
  from: string | null;
  value:
    | string
    | {
        rel: string;
        url: string;
        attributes?: {
          name: string;
        };
      };
};

export const toJsonPatchForCreate = (sourceWorkItem: WorkItem): WorkItemJsonPatch[] => {
  console.info(`toJsonPatchForCreate: ${sourceWorkItem.id}`);
  const jsonPatch = [
    {
      op: 'add',
      path: '/fields/System.AreaPath',
      from: null,
      value: convertClassificationPath(sourceWorkItem.fields['System.AreaPath']),
    },
    {
      op: 'add',
      path: '/fields/System.IterationPath',
      from: null,
      value: convertClassificationPath(sourceWorkItem.fields['System.IterationPath']),
    },
    {
      op: 'add',
      path: '/fields/System.Title',
      from: null,
      value: sourceWorkItem.fields['System.Title'],
    },
  ];
  return jsonPatch;
};

export const toJsonPatchForUpdate = async (
  sourceWorkItem: WorkItem,
  destinationWorkItems: WorkItem[],
): Promise<WorkItemJsonPatch[]> => {
  console.log(`toJsonPatchForUpdate: ${sourceWorkItem.fields['System.Title']}`);
  const jsonPatch = [
    {
      op: 'add',
      path: '/fields/System.AreaPath',
      from: null,
      value: convertClassificationPath(sourceWorkItem.fields['System.AreaPath']),
    },
    {
      op: 'add',
      path: '/fields/System.IterationPath',
      from: null,
      value: convertClassificationPath(sourceWorkItem.fields['System.IterationPath']),
    },
    {
      op: 'add',
      path: '/fields/System.Title',
      from: null,
      value: sourceWorkItem.fields['System.Title'],
    },
    {
      op: 'add',
      path: '/fields/System.State',
      from: null,
      value: sourceWorkItem.fields['System.State'],
    },
    {
      op: 'add',
      path: '/fields/System.Description',
      from: null,
      value: await convertLinks(
        addMetadata(
          sourceWorkItem.fields['System.Description'],
          sourceWorkItem.fields['System.CreatedBy'].displayName,
          sourceWorkItem.fields['System.CreatedDate'],
        ),
      ),
    },
    ...(sourceWorkItem.relations == null
      ? []
      : await relationsJsonPatch(sourceWorkItem.relations, destinationWorkItems)),
  ];

  console.info(JSON.stringify(jsonPatch));
  return jsonPatch;
};

const convertClassificationPath = (sourceClassificationPath: string) => {
  const sourceClassificationPathArray = sourceClassificationPath.split('\\').slice(1);
  const env = EnvironmentVariables.instance;
  const newClassificationPath = [env.DestinationProject, env.DestinationTeam]
    .concat(sourceClassificationPathArray)
    .join('\\');
  console.info(
    `convertClassificationPath: ${sourceClassificationPath} -> ${newClassificationPath}`,
  );
  return newClassificationPath;
};

const relationsJsonPatch = async (
  relations: WorkItemRelation[],
  destinationWorkItems: WorkItem[],
): Promise<WorkItemJsonPatch[]> => {
  const memo = await Memo.build();
  const attachmentsNotYetDuplicated = relations.filter(
    (relation) =>
      relation.rel === 'AttachedFile' && !memo.has('counterpartUrl', toId(relation.url)),
  );
  for (const attachment of attachmentsNotYetDuplicated) {
    const id = toId(attachment.url);
    const response = await attachmentDownloadAndUpload(
      id,
      encodeURIComponent(attachment.attributes.name),
    );
    memo.upsert('counterpartUrl', id, response.url);
  }

  const attachmentsExistingInDestinationWorkItems = new Map(
    destinationWorkItems.flatMap((workItem) =>
      workItem.relations == null
        ? []
        : workItem.relations.flatMap((relation) =>
            relation.rel === 'AttachedFile' ? [[relation.url, true]] : [],
          ),
    ),
  );

  console.info(JSON.stringify(attachmentsExistingInDestinationWorkItems));

  const attachments = relations.flatMap((relation) => {
    if (relation.rel !== 'AttachedFile') return [];
    const id = toId(relation.url);
    const counterpartUrl = memo.get('counterpartUrl', id) as string;
    if (attachmentsExistingInDestinationWorkItems.has(counterpartUrl.split('?')[0])) {
      console.info(`Attachment is already added to relations...`);
      return [];
    }
    return [
      {
        op: 'add',
        path: '/relations/-',
        from: null,
        value: {
          rel: relation.rel,
          url: counterpartUrl,
          attributes: {
            name: relation.attributes.name,
          },
        },
      },
    ];
  });

  const { organization } = destinationProfile;
  const baseUrl = `https://dev.azure.com/${organization}/_apis/wit/workItems/`;

  const others = relations.flatMap((relation) => {
    const id = toId(relation.url);
    if (relation.rel === 'AttachedFile' || !memo.has('counterpartId', id)) return [];
    const counterpartId = memo.get('counterpartId', id);
    return [
      {
        op: 'add',
        path: '/relations/-',
        from: null,
        value: {
          rel: relation.rel,
          url: [baseUrl, counterpartId].join(''),
        },
      },
    ];
  });

  const combinedJsonPatch = [...attachments, ...others];
  console.info(`Combined JsonPatch: ${combinedJsonPatch.length}`);
  return combinedJsonPatch;
};

const toId = (relationUrl: string): string => relationUrl.split('/').slice(-1)[0];
