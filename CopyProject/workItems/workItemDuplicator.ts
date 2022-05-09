import { getLogger } from 'log4js';
import { duplicateComments } from '../comments/commentDupulicator';
import { Memo } from '../keyValueStore/memo';
import { DestinationWorkItem, SourceWorkItem, WorkItem } from './workItemAPI';
import { toJsonPatchForCreate, toJsonPatchForUpdate } from './workItemJsonPatch';

const logger = getLogger();
logger.level = 'debug';

export const duplicateWorkItems = async () => {
  const sourceWorkItemIds = await SourceWorkItem.ids();
  const sourceWorkItems = await SourceWorkItem.list(sourceWorkItemIds);

  console.info(`Original project's workitem count: ${sourceWorkItemIds.length}`);

  for (const sourceWorkItem of sourceWorkItems) {
    await createWorkItemIfNotExists(sourceWorkItem);
  }

  const destinationWorkItemIds = await DestinationWorkItem.ids();
  const destinationWorkItems = await DestinationWorkItem.list(destinationWorkItemIds);

  for (const sourceWorkItem of sourceWorkItems) {
    await updateWorkItem(sourceWorkItem, destinationWorkItems);
  }

  const memo = await Memo.build();
  const counterpartIds = new Map(
    sourceWorkItemIds.map((workItemId) => [memo.get('counterpartId', workItemId) as number, true]),
  );
  const expiredWorkItemIds = destinationWorkItems
    .filter((workItem) => !counterpartIds.has(workItem.id))
    .map((workItem) => workItem.id);

  for (const id of expiredWorkItemIds) await DestinationWorkItem.delete(id);
};

const createWorkItemIfNotExists = async (sourceWorkItem: WorkItem): Promise<void> => {
  const memo = await Memo.build();
  if (memo.has('previousChangedDate', sourceWorkItem.id)) {
    console.log(`WorkItem Already Exists: ${sourceWorkItem.id}`);
    return;
  }

  console.log(`Create WorkItem: ${sourceWorkItem.id}`);

  const counterpart = await DestinationWorkItem.create(
    encodeURIComponent(sourceWorkItem.fields['System.WorkItemType']),
    toJsonPatchForCreate(sourceWorkItem),
  );
  memo.upsert(
    'previousChangedDate',
    sourceWorkItem.id,
    sourceWorkItem.fields['System.ChangedDate'],
  );
  memo.upsert('counterpartId', sourceWorkItem.id, counterpart.id);
  console.log(`WorkItem Creation Succeded: ${sourceWorkItem.id}`);
};

const updateWorkItem = async (
  sourceWorkItem: WorkItem,
  destinationWorkItems: WorkItem[],
): Promise<void> => {
  const memo = await Memo.build();
  const counterpartId = memo.get('counterpartId', sourceWorkItem.id) as number;
  const previousChangedDate = memo.get('previousChangedDate', sourceWorkItem.id);

  console.info(`Update WorkItem: ${sourceWorkItem.id}`);

  if (counterpartId == null || previousChangedDate == null) throw Error;

  const isChanged =
    Date.parse(sourceWorkItem.fields['System.ChangedDate']) >=
    Date.parse(previousChangedDate as string);

  if (!isChanged) return;

  console.log(`Update WorkItem: from ${sourceWorkItem.id} to ${String(counterpartId)}`);

  await DestinationWorkItem.update(
    counterpartId,
    await toJsonPatchForUpdate(sourceWorkItem, destinationWorkItems),
  );

  console.log(`duplicateComments: ${sourceWorkItem.fields['System.Title']}`);
  await duplicateComments(sourceWorkItem);

  memo.upsert('previousChangedDate', sourceWorkItem.id, new Date().toISOString());
};
